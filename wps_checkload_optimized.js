// ============================================================
// checkload() 加速优化版 —— 仅替换此函数即可
// 存放位置：复制到 v3 代码中替换原 checkload() 函数
// 涉及改动：仅 checkload()，不碰任何其他函数
// ============================================================

// ============================================================
// 配置项（可在此调整）
// ============================================================
// COMMENT_MODE 可选值：
//   'full'    — 问题单元格批注 + A列行汇总批注（原行为）
//   'rowOnly' — 只写A列行汇总批注，不标单元格级批注（速度最快）
//   'none'    — 不写任何批注，只标红 + 统计（极致速度）
//
// 注意：在 WPS JS 中，顶层不允许 var/const/let，
//       所以配置项放在函数开头，每次调用时生效。
//       如需全局切换，请在 checkload() 内第1行修改。
// ============================================================

function checkload() {

    // ==================== 配置 ====================
    var CHECKLOAD_COMMENT_MODE = 'full';  // 'full' | 'rowOnly' | 'none'

    // ==================== 状态保存 ====================
    var _screenUpdating = Application.ScreenUpdating;
    var _calculation = Application.Calculation;
    var _enableEvents = Application.EnableEvents;

    try {
        // ==================== 准备工作 ====================
        var shObj = Sheets.Item(1);
        var laRng = shObj.Range('A' + shObj.Rows.Count).End(xlUp);
        var lastRow = laRng.Row;
        var rowCount = lastRow - 1; // 数据行数（第2行到 lastRow）

        // 性能优化
        Application.ScreenUpdating = false;
        Application.Calculation = xlCalculationManual;
        Application.EnableEvents = false;

        // ==================== 1. 预编译正则和常量（循环外一次性创建）====================
        var reClass = /(\d+|[一二三四五六七八九十大中小]+)(班)?(年级)?(年)?(\d+|[一二三四五六七八九十]+|\(\d+\)|\([一二三四五六七八九十]+\))(班)?/;
        var reDisaster = /鼠疫|霍乱|传染性非典|肺炭疽/g;
        var rePositive = /阳性|\+/g;
        var reKindergarten = /幼儿园/g;
        var reMalaria = /疟/g;
        var reLocalInput = /本地|输入/g;

        var occpNeedUnit = ['民工', '教师', '医务人员', '工人', '干部职员'];
        var disNeedConfirmed = ['新型冠状病毒感染', '登革热'];

        // ==================== 2. 一次性计算当前时间 ====================
        // 使用 JS Date 计算 Excel 日期序列值（避免写入 =NOW() 公式触发全列计算）
        // Excel 序列号基准：1899-12-30 = 0（含 1900-02-29 Lotus 兼容偏移）
        function calcNow() {
            var now = new Date();
            var epoch = new Date(1899, 11, 30); // 1899-12-30 本地时
            var msDiff = now.getTime() - epoch.getTime();
            return msDiff / (24 * 60 * 60 * 1000);
        }
        var nowTime = calcNow();

        // ==================== 3. 批量读取关键列到内存 ====================
        // 需要读取的列：D(姓名) E(监护人) F(身份证) I(年龄) J(工作单位)
        //               L(地区) O(人群分类) P(病例分类) Q(疾病分型)
        //               S(诊断日期) T(死亡日期) U(疾病名称)
        //               AD(录入时间) AG(审核时间) AK(修改时间) AT(备注)
        var colKeys = ['D','E','F','I','J','L','O','P','Q','S','T','U','AD','AG','AK','AT'];
        var colCount = colKeys.length;
        var cache = {};
        // 预分配数组
        for (var ci = 0; ci < colCount; ci++) {
            cache[colKeys[ci]] = new Array(rowCount);
        }
        // 逐列批量读取
        for (var ci = 0; ci < colCount; ci++) {
            var col = colKeys[ci];
            var rng = shObj.Range(col + '2:' + col + lastRow);
            var vals = rng.Value2;
            // WPS JS 返回的 Value2 可能是一维数组、二维数组或单值
            if (vals != null) {
                if (typeof vals === 'object' && vals.length !== undefined) {
                    if (vals.length > 0 && typeof vals[0] === 'object' && vals[0] !== null && vals[0].length !== undefined) {
                        // 二维 [[v1], [v2], ...]
                        var len = vals.length < rowCount ? vals.length : rowCount;
                        for (var ri = 0; ri < len; ri++) {
                            cache[col][ri] = vals[ri][0];
                        }
                    } else {
                        // 一维 [v1, v2, ...]
                        var len = vals.length < rowCount ? vals.length : rowCount;
                        for (var ri = 0; ri < len; ri++) {
                            cache[col][ri] = vals[ri];
                        }
                    }
                } else {
                    cache[col][0] = vals;
                }
            }
        }
        // 便捷读取函数（基于 0-based 索引）
        function cv(col, idx) {
            var arr = cache[col];
            return (arr && idx < arr.length) ? arr[idx] : null;
        }

        // ==================== 4. 清理格式和批注（仅数据范围，不清理整列）====================
        shObj.Range('A1:AT' + lastRow).ClearFormats();
        shObj.Range('A1:AT' + lastRow).ClearComments();

        // ==================== 5. 写入 AG 列当前时间（静态值，非公式）====================
        // 原代码写入 =NOW() 公式，公式填充在 2 万行时显著慢
        // 改为写入静态值：先写第1行，再用 FillDown 扩展到全列
        shObj.Range('AG2').Value2 = nowTime;
        if (rowCount > 1) {
            shObj.Range('AG2:AG' + lastRow).FillDown();
        }
        shObj.Range('AG2:AG' + lastRow).NumberFormatLocal = 'G/通用格式';

        // ==================== 6. 第一阶段：规则判断（纯内存操作）====================
        // cellActions: 按 (row, col) 分组的问题消息
        // 格式: { 'row_col': { row: n, col: 'X', msgs: ['msg1', 'msg2'] } }
        var cellActions = {};
        // 记录只标红但不加批注的单元格（如 I 列年龄标记）
        var markOnlyCells = {};
        // rowSummary: 按行号的 A 列摘要
        var rowSummary = {};

        var stats = {
            certnum: 0, checkloadtime: 0, loadadjusttime: 0,
            aggnum1: 0, occpynum1: 0, occpynum2: 0,
            occpynum3: 0, occpynum4: 0,
            notenum: 0, yigannum: 0
        };

        function addAction(row, col, msg, statKey) {
            var key = row + '_' + col;
            if (!cellActions[key]) {
                cellActions[key] = { row: row, col: col, msgs: [] };
            }
            cellActions[key].msgs.push(msg);
            if (statKey && stats[statKey] !== undefined) {
                stats[statKey]++;
            }
        }

        function markOnly(row, col) {
            var key = row + '_' + col;
            markOnlyCells[key] = true;
        }

        for (var i = 2; i <= lastRow; i++) {
            var ri = i - 2; // 0-based 索引
            var rowquesall = 0;
            var rowquestion = '';

            // ---- (一) 身份证校验 ----
            var region = cv('L', ri);
            var cert = cv('F', ri);

            if ((region === '本县区' || region === '本市其它县区' || region === '其他省') && cert != null && cert !== '') {
                var certStr = String(cert);
                // 补 * 前缀
                if (certStr.length !== 20 && certStr.substring(0, 1) !== "'") {
                    if (certStr.substring(0, 1) !== '*') {
                        certStr = '*' + certStr;
                        // 必须写回单元格（用户期望看到 * 前缀）
                        shObj.Range('F' + i).Value2 = certStr;
                        // 同步更新缓存
                        cache['F'][ri] = certStr;
                    }
                }

                if (certStr.length !== 20) {
                    addAction(i, 'F', '非18位身份证', 'certnum');
                    rowquesall++;
                    rowquestion += rowquesall + '. 身份证不正确\n';
                } else {
                    var subString19 = certStr.substring(18, 19);
                    if (!isDigit(subString19)) {
                        if (subString19 === 'X' || subString19 === 'x') {
                            var subStr = certStr.substring(1, 18);
                            if (isNaN(subStr)) {
                                addAction(i, 'F', '身份证非17位数字+X', 'certnum');
                                rowquesall++;
                                rowquestion += rowquesall + '. 身份证不正确\n';
                            }
                        } else {
                            addAction(i, 'F', '身份证最后一位非X', 'certnum');
                            rowquesall++;
                            rowquestion += rowquesall + '. 身份证不正确\n';
                        }
                    } else {
                        var subStr = certStr.substring(1, 19);
                        if (isNaN(subStr)) {
                            addAction(i, 'F', '身份证非18位数字', 'certnum');
                            rowquesall++;
                            rowquestion += rowquesall + '. 身份证不正确\n';
                        }
                    }
                }
            }

            // ---- (二) 年龄与人群关系 ----
            var ageVal = cv('I', ri);
            var age = String(ageVal || '');
            var guardian = String(cv('E', ri) || '');
            var name = String(cv('D', ri) || '');
            var occp = cv('O', ri);
            var compyVal = cv('J', ri);
            var compy = String(compyVal || '').replace(/\（/g, '(').replace(/\）/g, ')');

            var ageHasMonth = age.indexOf('月') !== -1;
            var ageHasDay = age.indexOf('天') !== -1;
            var ageHasYear = age.indexOf('岁') !== -1;
            var ageHasRi = age.indexOf('日') !== -1;
            var ageNum = 0;
            if (ageHasYear) ageNum = Number(age.replace('岁', ''));

            // 2a: 14岁以下需监护人
            if (ageHasMonth || ageHasDay || ageNum <= 14) {
                if (guardian === '') {
                    addAction(i, 'E', '14岁以下患者需要填写监护人信息', 'aggnum1');
                    markOnly(i, 'I');
                    rowquesall++;
                    rowquestion += rowquesall + '. 人群年龄与其职业有问题\n';
                }
                if (guardian !== '' && guardian === name) {
                    addAction(i, 'E', '14岁以下患者与监护人名字一致', 'aggnum1');
                    markOnly(i, 'I');
                    rowquesall++;
                    rowquestion += rowquesall + '. 人群年龄与其职业有问题\n';
                }
            }

            // 2b: 学生/幼托儿童需填写数字班级
            if (occp === '学生' || occp === '幼托儿童') {
                if (compy === '') {
                    addAction(i, 'J', '学生或幼托儿童时，要求填写单位为数字班级', 'occpynum1');
                    rowquesall++;
                    rowquestion += rowquesall + '. 学生要求填写单位为数字班级\n';
                } else if (!reClass.test(compy)) {
                    addAction(i, 'J', '学生或幼托儿童时，要求填写单位为数字班级', 'occpynum1');
                    rowquesall++;
                    rowquestion += rowquesall + '. 学生要求填写单位为数字班级\n';
                }
            }

            // 2c: 特定职业需填写单位
            if (occpNeedUnit.indexOf(occp) !== -1) {
                if (!compy) {
                    addAction(i, 'J', occp + '须填写发病时所在的工作单位名称', 'occpynum2');
                    rowquesall++;
                    rowquestion += rowquesall + '.' + occp + '必须填写发病时所在的工作单位名称\n';
                }
            }

            // 2d: 年龄与职业对应
            if (ageHasYear) {
                if (ageNum > 3 && ageNum < 5 && (occp === '散居儿童' || occp === '学生')) {
                    addAction(i, 'O', '3-5岁不对应幼托儿童为可疑', 'occpynum3');
                    markOnly(i, 'I');
                    rowquesall++;
                    rowquestion += rowquesall + '. 3-5岁不对应幼托儿童为可疑\n';
                } else if (ageNum > 7 && ageNum < 16 && (occp === '散居儿童' || occp === '幼托儿童')) {
                    addAction(i, 'O', '7-18岁不对应学生为可疑', 'occpynum3');
                    markOnly(i, 'I');
                    rowquesall++;
                    rowquestion += rowquesall + '. 7-18岁不对应学生为可疑\n';
                } else if (ageNum < 16 && ['学生', '幼托儿童', '散居儿童'].indexOf(occp) === -1) {
                    addAction(i, 'O', '16岁以下从业人员不符合实际', 'occpynum3');
                    markOnly(i, 'I');
                    rowquesall++;
                    rowquestion += rowquesall + '. 16岁以下从业人员不符合实际\n';
                }
            } else if (ageHasMonth || ageHasRi || (ageHasYear && ageNum <= 3)) {
                // 注意：此处保留 checkload 原始条件 "日" 不变（与原始第192行一致）
                if (occp !== '幼托儿童' && occp !== '散居儿童') {
                    addAction(i, 'O', '3岁以下应为幼托儿童或散居儿童', 'occpynum3');
                    markOnly(i, 'I');
                    rowquesall++;
                    rowquestion += rowquesall + '. 3岁以下应为幼托儿童或散居儿童\n';
                }
            }

            // 2e: 幼儿园单位与职业对应
            if (reKindergarten.test(compy) && occp !== '幼托儿童' &&
                (ageHasMonth || ageHasDay || ageNum < 16)) {
                addAction(i, 'O', '工作单位为幼儿园但人群分类未对应幼托儿童', 'occpynum4');
                rowquesall++;
                rowquestion += rowquesall + '. 工作单位为幼儿园但人群分类未对应幼托儿童\n';
            }

            // ---- (三) 逻辑与疾病信息 ----
            var dis = cv('U', ri);
            if (dis != null) {
                var disStr = String(dis);
                var matches_dis = reDisaster.test(disStr);
                reDisaster.lastIndex = 0; // 重置正则 lastIndex（g 标志导致的状态残留）

                var load_time = cv('AD', ri);
                var diag_time = cv('S', ri);

                // 迟报检查
                if (load_time != null && diag_time != null) {
                    if (load_time - diag_time > 2 / 24 && matches_dis) {
                        addAction(i, 'S', '甲类报告超2小时', 'checkloadtime');
                        rowquesall++;
                        rowquestion += '\n' + rowquesall + '. 甲类报告超2小时';
                    } else if (load_time - diag_time > 1 && !matches_dis) {
                        addAction(i, 'S', '乙类报告超24小时', 'checkloadtime');
                        rowquesall++;
                        rowquestion += '\n' + rowquesall + '. 乙类报告超24小时';
                    }
                }

                // 迟审检查
                var modified_time = cv('AK', ri);
                if (load_time != null && nowTime != null) {
                    var timediff = nowTime - load_time;
                    var timediff2;
                    if (modified_time != null) {
                        timediff2 = nowTime - modified_time;
                    } else {
                        timediff2 = 999;
                    }
                    var isDot = (modified_time != null && String(modified_time) === '.');

                    if (timediff > 2 / 24 && matches_dis && (timediff2 > 2 / 24 || isDot)) {
                        addAction(i, 'AD', '甲类审核超2小时', 'loadadjusttime');
                        rowquesall++;
                        rowquestion += '\n' + rowquesall + '. 甲类审核超2小时';
                    }
                    if (timediff > 1 && !matches_dis && (timediff2 > 1 || isDot)) {
                        addAction(i, 'AD', '乙类审核超24小时', 'loadadjusttime');
                        rowquesall++;
                        rowquestion += '\n' + rowquesall + '. 乙类审核超24小时';
                    }
                }

                // 病例分类与疾病逻辑校验
                var chctype = cv('P', ri);
                var chcnote = String(cv('AT', ri) || '');
                var yigan_crt2 = cv('Q', ri);

                if (chctype === '确诊病例') {
                    if (!rePositive.test(chcnote)) {
                        addAction(i, 'AT', '确诊病例要求有病原学结果', 'notenum');
                        rowquesall++;
                        rowquestion += '\n' + rowquesall + '. 确诊病例要求有病原学结果';
                    }
                }

                if (disStr === '乙肝' && (chctype !== '确诊病例' || yigan_crt2 !== '慢性')) {
                    addAction(i, 'U', '乙肝必须是确诊病例，慢性', 'yigannum');
                    rowquesall++;
                    rowquestion += '\n' + rowquesall + '. 乙肝必须是确诊病例，慢性';
                }

                if (disStr === '丙肝') {
                    if (chctype === '确诊病例' && yigan_crt2 !== '慢性' && yigan_crt2 !== '急性') {
                        addAction(i, 'U', '丙肝确诊病例必须是慢性或者急性', 'yigannum');
                        rowquesall++;
                        rowquestion += '\n' + rowquesall + '. 丙肝确诊病例必须是慢性或者急性';
                    } else if (chctype === '临床诊断病例' && yigan_crt2 !== '未分型') {
                        addAction(i, 'U', '丙肝临床诊断病例必须是未分型', 'yigannum');
                        rowquesall++;
                        rowquestion += '\n' + rowquesall + '. 丙肝临床诊断病例必须是未分型';
                    }
                }

                if (disNeedConfirmed.indexOf(disStr) !== -1 && chctype !== '确诊病例') {
                    addAction(i, 'U', disStr + '必须要求病例分类为"确诊病例"', 'yigannum');
                    rowquesall++;
                    rowquestion += '\n' + rowquesall + '. ' + disStr + '须分类为"确诊病例"';
                }

                reMalaria.lastIndex = 0;
                if (reMalaria.test(disStr) && (chctype !== '确诊病例' || yigan_crt2 !== '未分型')) {
                    addAction(i, 'U', '疟疾分为间日疟、恶性疟和未分型三类，其余分型错误', 'yigannum');
                    rowquesall++;
                    rowquestion += '\n' + rowquesall + '. 疟疾分型错误';
                }

                if ((disStr === '间日疟' || disStr === '恶性疟' || disStr === '登革热') && !reLocalInput.test(chcnote)) {
                    addAction(i, 'U', '需要注明本地感染还是输入（外市输入、境外输入）', 'notenum');
                    rowquesall++;
                    rowquestion += '\n' + rowquesall + '. 需要注明本地感染还是输入';
                }

                if (disStr === '炭疽' && yigan_crt2 !== '肺炭疽' && yigan_crt2 !== '皮肤炭疽' && yigan_crt2 !== '未分型') {
                    addAction(i, 'U', '炭疽分为肺炭疽、皮肤炭疽和未分型三类，其余分型错误', 'yigannum');
                    rowquesall++;
                    rowquestion += '\n' + rowquesall + '. 炭疽分型错误';
                }
            }

            // 死亡标记
            var deathdate = cv('T', ri);
            if (deathdate != null && deathdate !== '.') {
                addAction(i, 'T', '死亡人群要进一步确认', 'yigannum');
                rowquesall++;
                rowquestion += '\n' + rowquesall + '. 死亡人群要进一步确认';
            }

            // 保存行摘要
            if (rowquestion) {
                rowSummary[i] = rowquestion;
            }
        }

        // ==================== 7. 第二阶段：统一应用格式和批注 ====================

        // 7a: 所有问题单元格标红 + 批注
        for (var key in cellActions) {
            var act = cellActions[key];
            var cell = shObj.Range(act.col + act.row);
            cell.Font.Color = 255;
            cell.Interior.Color = 65536;

            if (CHECKLOAD_COMMENT_MODE === 'full') {
                // 合并同单元格的多条消息
                var commentText = act.msgs.join('\n');
                cell.AddComment(commentText);
            }
        }

        // 7b: 只标红不加批注的单元格（如 I 列年龄标记）
        for (var key in markOnlyCells) {
            var parts = key.split('_');
            var rowNum = parseInt(parts[0]);
            var colLetter = parts[1];
            // 如果已经被 cellActions 覆盖（有批注），则跳过
            if (cellActions[key]) continue;
            var cell = shObj.Range(colLetter + rowNum);
            cell.Font.Color = 255;
            cell.Interior.Color = 65536;
        }

        // 7c: A 列写行级问题摘要批注
        if (CHECKLOAD_COMMENT_MODE === 'full' || CHECKLOAD_COMMENT_MODE === 'rowOnly') {
            for (var rn in rowSummary) {
                var aCell = shObj.Range('A' + rn);
                var commentText = rowSummary[rn];
                // 去除首尾换行符（与原始行为一致）
                commentText = commentText.replace(/^\n+|\n+$/g, '');
                aCell.AddComment(commentText);
            }
        }

        // ==================== 8. 格式化日期列 ====================
        shObj.Range('H:H,R:R,Z:Z,T:T').NumberFormatLocal = "yyyy/m/d;@";
        shObj.Range('S:S,AD:AD,W:X,AK:AL,AG2:AG' + lastRow).NumberFormatLocal = "yyyy/m/d h:mm;@";

        // ==================== 9. 弹出汇总 ====================
        alert('分析完毕！\n' +
              '身份证有问题：' + stats.certnum + '个。\n' +
              '迟报卡：' + stats.checkloadtime + '个。\n' +
              '迟审卡：' + stats.loadadjusttime + '个。\n' +
              '家长信息问题：' + stats.aggnum1 + '个。\n' +
              '班级问题：' + stats.occpynum1 + '个。\n' +
              '单位填报问题：' + stats.occpynum2 + '个。\n' +
              '年龄职业对应异常：' + stats.occpynum3 + '个。\n' +
              '单位与年龄不对应：' + stats.occpynum4 + '个。\n' +
              '病例类型问题：' + stats.yigannum + '个。\n' +
              '备注问题：' + stats.notenum + '个。');

    } catch (e) {
        alert('checkload 运行出错：' + e.message);

    } finally {
        // 无论成功还是出错，必须恢复 Application 状态
        Application.ScreenUpdating = _screenUpdating;
        Application.Calculation = _calculation;
        Application.EnableEvents = true;
    }
}

// ============================================================
// 注意：本函数依赖顶层 isDigit() 函数，该函数在 v3 中已存在：
//   function isDigit(char) { return /^[0-9]$/.test(char); }
// 如果原始代码中没有，请添加此函数至顶层。
// ============================================================
