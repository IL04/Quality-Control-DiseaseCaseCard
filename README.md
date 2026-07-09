# 报卡质控宏 - Quality Control for Disease Case Cards

WPS 宏工具，用于传染病报告卡的质量控制审核，支持 WPS Office (JSA)。

## 功能

### 日常报表审核 (checkload)

精简版质控，快速扫描三项核心指标：

| 检查项 | 说明 |
|--------|------|
| 身份证校验 | 检查身份证格式（18位、前缀\*、末位X校验） |
| 迟报检查 | 诊断时间 → 录入时间超阈值（甲类>2h / 乙类>24h） |
| 迟审检查 | 录入时间 → 审核时间超阈值，支持订正卡前后对比 |

### 订正卡逻辑

自动识别订正卡（AK列有数据），区分订正前后：

| 检查 | 公式 |
|------|------|
| 订正前迟报 | 录入时间(AD) - 订正前诊断时间(W) |
| 订正前迟审 | 订正前终审时间(X) - 录入时间(AD) |
| 订正后迟报 | 订正报告时间(AK) - 诊断时间(S) |
| 订正后迟审 | 县区审核时间(AG) - 订正报告时间(AK) |

### 报卡质控 (qualitycontrol)

完整质控扫描，包含全部规则：

- 身份证校验
- 现住址完整性（光明区需含街道信息）
- 年龄与监护人/人群分类对应
- 迟报迟审检查
- 疾病分类与分型逻辑（乙肝/丙肝/新冠/登革热/疟疾/炭疽）
- 病原学结果检查
- 死亡病例标记

### 重卡审核 (repeatcontrol)

基于身份证号的重复卡片检测与处理。

## 使用方式

1. 打开 WPS Office
2. `Alt+F11` 打开 JS 宏编辑器
3. 将 `wps_report_quality_control_v4.txt` 全部内容复制粘贴到模块中
4. 保存后重启 WPS
5. 工具栏将出现「报卡质控」选项卡，包含 4 个按钮

## 配置

在 `checkload()` 函数开头可切换批注模式：

```javascript
var CHECKLOAD_COMMENT_MODE = 'full'; // full | rowOnly | none
```

- `full`: 问题单元格批注 + A列行级汇总（默认）
- `rowOnly`: 仅A列汇总批注，单元格只标红
- `none`: 仅标红，不写任何批注

## 技术说明

- 语言：WPS JSA (JavaScript for Application)
- 兼容性：WPS Office 专业版/个人版
- 顶层代码仅允许 `function` 声明，变量通过 `init()` 惰性初始化
- 自动备份机制（qualitycontrol）：运行前复制当前工作表
- Application 状态（ScreenUpdating/Calculation/Events）使用 try/finally 确保恢复
