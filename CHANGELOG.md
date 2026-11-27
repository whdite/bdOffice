# 更新日志

## Windows 1.0.4

- 新建 Word 文档会按要求设置标题字号、对齐方式和表格列宽。指定页数时，助手会实际分页，并在原对话中修正偏差。
- 局部修改会检查指定段落和标题的范围。润色时会核对原文中的数字、审批义务和金额上限。
- 多文档任务会逐份核对交付文件与来源，可对照 Excel 单元格、Word 段落和 PDF 页中的事实。
- 首页和 Word 侧栏可在本轮授权内纠正问题。工作记录保留核验结果与未达项，已选的写作预设会传递给 Word 写作器。

[安装包与校验](https://github.com/whdite/bdOffice/releases/tag/v1.0.4)


## Windows 1.0.3

- Word 编辑遇到空选区或目标不存在时，可在原对话中重新定位。需要核对写入结果时，也可直接在原对话继续任务。
- 首页与 Word 侧栏共用审批规则。读取和预演可直接执行，可撤销的编辑可对同一文件的当前版本批准本轮操作。
- Word 生成和编辑保留标题、段落、表格结构。保存、另建文件、修改页眉和样式后，助手会核对相应文件或文档部件。
- 「关于」更新了产品说明。稳定版和 Beta 通道从公开发布仓库检查更新，GitHub API 限流时改用发布订阅和签名清单。

[安装包与校验](https://github.com/whdite/bdOffice/releases/tag/v1.0.3)


## Windows 1.0.2

- 手机协作可在设置中启动和停止 Remote，也可保存为随大狗自动启动。
- 连接路径支持自动选路和服务器转发。自动模式依次尝试 LAN、P2P、TURN 和加密 Relay，面板显示实际手机链路。
- 修复桌面重启后使用已保存设备凭据的身份校验。已授权的设备可以继续连接。
- 配对界面标示本机调试邀请，并提供端口占用、服务不可达和授权失败的反馈。

[安装包与校验](https://github.com/whdite/bdOffice/releases/tag/v1.0.2)

## Windows 1.0.1

- 「文件与记忆」在输入框旁打开，支持直接管理附件、自动记忆和来源。记忆摘要可先看两行，再展开全文。
- 「每次审批」和「自动执行」直接切换并保存，首页与文件助手同步。
- 大狗将同一次执行的工具步骤归为一组，完成后收起。
- Word 可在文档内审阅本轮修改，并按文件批准本轮编辑和保存。成果记录支持回退整轮改动。

[安装包与校验](https://github.com/whdite/bdOffice/releases/tag/v1.0.1)


## Windows 1.0.0

### 更新内容

- 旧版 .ppt 演示文稿可以转换为可编辑的 .pptx 副本。
- Excel 切片器和时间线可随文件恢复，改进排序时的共享公式与导出计算。
- 改进 Word 表格编辑、文档保存恢复和 PDF 页面顺序校验。
- MCP 增加文档检查、按块读取和转换工具。大狗优化长会话上下文与摘要用量统计。
- 升级 Electron 运行时，加强文件操作校验与旧格式导入限制。

### 下载与校验

- [Windows x64 安装包](https://github.com/whdite/bdOffice/releases/download/v1.0.0/BigDogOffice-1.0.0-windows-x64.exe)
- [官网下载](https://bdoffice.deepthink-x.com/updates#downloads)
- [Android 0.1.4 APK](https://github.com/whdite/bdOffice/releases/download/v0.9.11/AI-BigDog-0.1.4-android.apk)
- [SHA-256 校验文件](https://github.com/whdite/bdOffice/releases/download/v1.0.0/SHA256SUMS.txt)

安装包提供独立 Minisign 签名，沿用大狗Office 的发布公钥。


## Windows 0.9.11 / Android 0.1.4（2026-11-17）

大狗Office Windows **0.9.11** · AI 大狗 Android **0.1.4**

本次提供 Windows x64 安装器与独立 Android APK。macOS/Linux 本轮未重建，请使用明确标为 0.9.3 的历史下载。

### 主要更新

- 全局 AI 大狗整合对话、Session 记忆、任务状态与成果记录，简化设置和审批入口。
- 文件拖入作为引用并登记路径；补齐全局文件发现和搜索，修复 GPT-6 工具调用兼容。
- Cua SDK 电脑操作工具通过独立辅助进程执行，沿用现有模型、权限和任务管理。
- 手机支持二维码、图片或完整配对码连接；会话、任务确认、重试及成果与桌面保持一致。
- Android 按 dsh-Go 的界面方式重做会话侧栏、聊天输入栏、任务弹层和分组设置，支持深浅主题、字号与内置 Office/PDF 预览。

### 下载

- [Windows x64 0.9.11](https://github.com/whdite/bdOffice/releases/download/v0.9.11/BigDogOffice-0.9.11-windows-x64.exe)
- [Android 0.1.4 APK](https://github.com/whdite/bdOffice/releases/download/v0.9.11/AI-BigDog-0.1.4-android.apk)
- [macOS/Linux 0.9.3 历史版本](https://github.com/whdite/bdOffice/releases/tag/v0.9.3)

手机通过自部署连接服务协作，由电脑执行任务。Windows 安装器和 Android APK 的签名与校验方式见下载说明。[手机协作配置](docs/REMOTE.md) · [下载与校验](DOWNLOADS.md)

### 文件校验

两个安装包、版本清单和 SHA256SUMS 均附 Minisign 签名，沿用现有公钥，未轮换。

| 文件 | SHA-256 |
| --- | --- |
| Windows EXE | `315891100678f8726aec6165d8504c06bb176ad42779ddf25003ed979b16fe0e` |
| Android APK | `e4700b5e2ef470d23506a1707e3b554c4b0eb8cab4bebd65b8ddc0e4b9d94667` |

[下载与校验说明](https://github.com/whdite/bdOffice/blob/main/DOWNLOADS.md)。Minisign 不替代操作系统的发布者认证。



## 0.9.3 · 编辑引擎与 Home Agent 更新（2026-11-05）

- 合入经审查的官方编辑引擎更新，覆盖 Word 解析、排版、导入导出，表格公式与旧格式转换，以及演示文稿、PDF、Markdown 和 HTML 的相关修复。
- Home Agent 支持真实办公 CLI 文档生成及图片／文档附件协作；独立 Agent、编辑器分离和回挂继续保留会话、工作区、草稿与文件能力。
- 保留大狗现有登录、账户与模型路由、审批、撤销及未保存内容保护；外部 CLI 不能覆盖编辑器正在持有的文件。
- 修复 Windows 安装脚本中停用函数造成的严格打包失败，命令行 PATH 仍由用户在「设置 → 集成」中显式启用。
- 桌面版本为 0.9.3，内置办公 CLI 仍独立使用 0.10.0，操作 Skill 为 2.9.0。本次不包含企业版功能。
- 沿用 0.9.2 起的 Minisign 发布身份 `60FD3AF1A46DC5BC`，本次没有轮换；旧版和历史公钥均保留。

提供 Windows x64 EXE、Linux x64 AppImage／DEB／RPM、macOS Apple Silicon ARM64 DMG／ZIP，以及各包 Minisign 签名、签名清单和 SHA-256 校验文件。Windows 尚无 Authenticode；macOS 使用 ad-hoc 签名、未公证，Minisign 不替代系统发布者认证。


## 0.9.2 · 上游整合与独立编辑器（2026-09-23）

- Word 新增 Worker 解析、分批加载期间编辑和增量分页；在保存、导出、草稿和 Agent 读取路径保留文档完整性保护。
- Word 与表格支持独立编辑器窗口，与大狗独立 Agent 共存；适配实时文档访问、未保存提示及取消退出后继续编辑。
- 文档图片分析复用现有媒体 API；Parallel 作为显式选择的搜索服务；模型发现复用大狗入口，保留默认服务和账户流程。
- 整合表格原生计算、各编辑器与办公 CLI 的功能改进和修复；修复取消退出后原生表格引擎失效，以及异步退出无法完成的问题。
- 保留大狗品牌、审批、撤销、账户与 GitHub 更新通道；本次不包含企业版功能。
- **Minisign 公钥轮换**：原私钥无法恢复，0.9.2 使用新身份，旧公钥继续用于历史版本。轮换没有旧身份交叉签名，下载节点须审查[公告](SIGNING-KEYS.md)后更新固定公钥。

Windows/Linux x64、macOS Apple Silicon ARM64 安装包附 Minisign 签名。Windows 无 Authenticode，macOS 为 ad-hoc 签名、未公证；文件签名不替代系统发布者认证。


## 0.9.0 · 办公 CLI 与流式生成（2026-09-14）

- 内置 `bigdogoffice` 命令行，提供结构化办公操作、批量处理与无界面导出；需要排版的转换复用应用编辑引擎。
- 设置页新增集成入口，可向已有 Claude Code、Codex、Cursor 等助手安装办公操作 Skill，并显式配置命令行入口；不会自动安装外部 Agent 或修改 PATH。
- Word 长文与 HTML 规划／页面生成使用流式展示，支持取消、超时和分批刷新；中断内容由用户决定采用或丢弃。
- 草稿生成期间保护保存、导出及其他 Agent 访问；丢弃草稿不触发自动保存。
- 大狗 Agent 新增文档导出工具，保留审批、工作区限制、源文件变更检查和撤销。需先保存并关闭源文档；不会静默读取旧内容。
- 复用表格、演示文稿的共享操作层；保留现有独立 Agent 窗口、模型设置、账户与 GitHub 更新通道。不包含企业版。

Windows/Linux x64、macOS Apple Silicon ARM64 安装包均附 Minisign 签名。Windows 未配置 Authenticode，macOS 为 ad-hoc 签名、未公证；文件签名不替代系统发布者认证。

## 0.8.7 · 独立 Agent 与功能更新（2026-09-14）

- 大狗 Agent 可在编辑器侧栏和独立窗口之间切换，保留草稿、回复和待审批状态；独立窗口默认显示完整工作区、会话与文件。
- 审批模式放到发送区，状态改为紧凑显示，工具详情可展开；收起窗口不等同于停止任务。
- 修复引用准备阶段的停止、插话、删除会话后迟到请求、并发发送和超时边界。
- 新增 Zotero 引用与参考文献集成，需要本机 Zotero 服务；新增 Requesty API 厂商及自动模型列表读取，并修复旧配置迁移。
- 接入官方 47 项独立补丁，改善 CSV 数字精度与多行引号、表格替换与公式引用、DOCX 保护与引用字段、PDF 打印与样式验证、多语言字数和打印范围。
- 保留现有 Agent、审批、账户、原编辑引擎及 GitHub 更新通道；不包含尚在开发的企业版，也不在本次移入大范围 CLI／引擎拆分改造。

0.8.7 高于 0.8.6，可触发正常版本更新提示。安装仍需用户主动完成。Windows/Linux x64 与 macOS Apple Silicon 下载见本版本 Release；Minisign 校验不替代系统发布者证书。

## 0.8.6 · Agent 性能优化重发（2026-09-12）

- 六编辑器 AI 流式文本按 32ms 批量刷新，完成、工具、停止和错误边界即时更新。
- 历史会话按需加载，并发恢复去重；工作簿解析按文档版本复用和失效，运行结束释放派生缓存。
- PDF 按页处理待保存文本，未修改页复用只读索引，保留 OCR、重复文字和原位置检查。
- 流式事件异步写入，保持审批/结束状态顺序，修复迟到写入和历史恢复边界。
- 摘要增加输入/输出硬上限与本地降级。工具审批、撤销、工作区与账户权限保持不变。

版本号仍是 0.8.6，安装包、哈希和签名重新生成。旧 0.8.6 用户需手动重新下载安装；下载服务器需要按已签名清单中的 SHA-256 刷新同版本文件，不应仅按版本号跳过同步。此版本不包含企业版或完整跨会话长期记忆。

### 0.8.6 首次通道迁移

- 「关于」的仓库入口、更新检查与下载入口统一为 `whdite/bdOffice` GitHub Releases。
- 新版本在顶部标签栏内提示，不弹窗、不打断编辑，可在本次会话隐藏。
- 检查版本通道、系统与 CPU 架构，跳过草稿、残缺上传与降级；离线或限流不会误报已是最新。
- 修复 macOS 在异步关闭文档之后无法完成显式退出的问题，保留取消关闭与保存确认。
- 下载门户增加 GitHub 下载节点，移除“筹备中”及公开源码的旧说明。
- Windows 实际安装包已通过更新检查→顶部提示→对应版本页跳转→关闭提示的完整链路，以及 HTML 保存和 PDF／DOCX 导出验证。
- Linux 三种安装包通过版本、架构、686 项编辑器输出、WASM、许可文件和原生引擎核对；不将资源检查视为完整 Linux GUI 验收。
- macOS Apple Silicon 安装包通过原生组件、ad-hoc 签名、更新提示与版本页跳转、HTML 保存、PDF／DOCX 导出及正常退出验证。

旧版本内置的更新地址不会被仓库改动自动重写，首次迁移需手动安装 0.8.6。
本次不包含尚未完成的企业版功能，不启用静默安装，不公开桌面源码或构建流程。

## 0.8.5

- 合并并适配官方 Docs、Sheets、Slides、PDF 和 Markdown 的修复。
- 新增 HTML 编辑器：打开文件、编辑源代码、沙箱预览、保存、导出 PDF／DOCX。
- HTML 编辑器接入内嵌 AI 的审批机制。
- 保留大狗品牌、全局 Agent、账户连接、模型设置和自有下载通道。
- 对编辑器输出、原生表格引擎、PDF WASM 资源及许可证进行安装包检查。
- Windows 安装包通过实际启动与 HTML 打开、编辑保存、PDF／DOCX 导出测试。
- Linux 原生表格引擎通过测试，安装包完成架构、版本、资源与文件完整性核对；未把这些检查表述为完整的 Linux GUI 验收。

此版本记录此前的 Windows/Linux 发布及 HTML 引擎适配。GitHub 最新安装包以最新 Release 为准，不以旧包替换新版缺失平台。
