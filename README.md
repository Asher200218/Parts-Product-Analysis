# 零件 · 产品 · 工人 分析工具

> A fully local web app for analyzing parts, products, workers and orders: theoretical work hours, scheduling, retrospectives, and theory-vs-actual comparison against hour reports. Runs entirely in the browser — no network, no uploads.

一个**完全在本机运行**的分析网页：录入/导入 产品、零件(VC)、工人能力、订单和 hour report 工时台账，自动计算理论工时，给出**排期**（接下来怎么干）、**复盘**（过去干得怎么样）以及**理论 vs 实际**的对比。

🔒 **100% 本地**：所有数据只存在浏览器的 localStorage 里；服务器只监听 `localhost`；Excel 解析使用内置的 `web/vendor/xlsx.full.min.js`，断网也能完整使用。

## 快速开始

### Windows

1. 把整个文件夹拷到本地磁盘。
2. 双击 `零件产品分析.exe`。
   如出现 SmartScreen 提示：点「更多信息」→「仍要运行」（程序未购买签名证书，本身不联网）。
3. 保持黑色小窗口开着，浏览器会自动打开 <http://localhost:4173>。用完关掉黑窗口即可，数据不会丢。

exe 打不开时，双击 `备用启动(无需exe).bat`（用系统自带的 Python 或 PowerShell 起服务器，端口相同、数据互通）。

### macOS

双击 `web/launch-mac.command`（需要 Python 3；首次若提示「未验证的开发者」，右键 → 打开）。

### 任意系统（手动）

```bash
cd web
python3 -m http.server 4173
```

然后打开 <http://localhost:4173>。不建议直接双击 `index.html`（`file://` 下部分浏览器会禁用 localStorage）。

## 功能页面

| 页面 | 作用 |
|---|---|
| **产品** | 登记产品名（如 Alpha、Beta），作为工人能力表的列和订单的类型 |
| **零件 (VC)** | 每个 VC 的难度（高/低/无）、重复设计工时 t7、首单设计工时 t8；可设校对系数 |
| **工人能力** | 每人每产品的设计/校验等级（A/B/C）；请假登记 |
| **订单** | 智能导入公司系统导出的订单大表，自动识别表头并映射列 |
| **导入/导出** | 拖放自动识别文件；导出 Excel；整体备份/恢复（JSON） |
| **分析** | 理论工时汇总、理论 vs 实际报工（全员 / 每人 / 每单） |
| **排期** | 只排未完成订单：连续排 / 按交期 / 贴近交期三种方案，个人排期与重新分配建议 |
| **复盘** | 只看已完成订单：准时率、完成速度、工人负载、标准 vs 实际效率 |
| **工时台账** | 导入 `HourReport_{年}_Analysis.xlsx`，一键应用实际工时口径 |

## 核心计算规则

| 规则 | 内容 |
|---|---|
| 设计工时 | Σ(订单内各 VC 工时) × 数量系数：≤5 个 VC ×1，6–10 ×1.2，>10 ×1.5 |
| 首单/重复 | 同一「产品+VC」按创建日期最早的一单用 t8，之后用 t7 |
| 特殊值 | 工时为 0 → 按 0.03h；VC 999 → 固定 0.08h |
| 校对工时 | 设计 × 校对系数（默认 0.3，可用台账实际值校准） |
| 工作日 | 只算周一至周五；年假按每年 260 个工作日均摊扣产能 |
| 订单号 | Sales Order + "-" + Item（去前导 0），如 `123456789-1200` |
| 已完成/未完成 | 看「实际完成日期」是否填写：未完成 → 排期，已完成 → 复盘 |

## 数据与备份

数据跟着「这台电脑 + 这个浏览器 + `localhost:4173`」走。换电脑、换浏览器或清除浏览器数据都会看不到原数据，请定期在「导入/导出」页做**整体备份**（JSON），在新环境拖入即可恢复。

## 目录结构

```
零件产品分析.exe            Windows 主程序（网页已内嵌）
使用说明.html               完整使用说明
快速开始.txt                三行速览
备用启动(无需exe).bat       exe 被拦截时的备用启动方式
web/
  index.html / app.js / analysis.js / styles.css   网页源码
  vendor/xlsx.full.min.js   本地 Excel 解析库
  launch-windows.bat        Windows 启动器
  launch-mac.command        macOS 启动器
  launch.ps1                PowerShell 本地服务器
  build-windows-exe.ps1     用 ps2exe 把 launch.ps1 打包成 exe
```

## 自己构建 exe

在 Windows 上右键 `web/build-windows-exe.ps1` →「使用 PowerShell 运行」。它会从 PowerShell 库下载 `ps2exe`，把 `launch.ps1` 打包成 exe。

更多细节（工作流建议、蓝/橙对账标色、FAQ）见 [`使用说明.html`](使用说明.html) 和 [`web/README.md`](web/README.md)。
