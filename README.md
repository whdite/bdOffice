<p align="center"><img src="assets/bigdog-office.png" width="112" alt="大狗Office"></p>

# 大狗Office

让 AI 和你一起完成文档，而不只是回答问题。

大狗Office 是一款基于 GenOffice 构建的桌面办公应用，将文档、表格、演示文稿、PDF、Markdown 和 HTML 编辑放进同一个工作空间，并提供大狗 Agent 与编辑器内嵌 AI 助手。

[下载安装包](https://github.com/whdite/bdOffice/releases/latest) · [国内下载入口](https://bdoffice.deepthink-x.com/updates) · [更新日志](CHANGELOG.md) · [下载与校验说明](DOWNLOADS.md)

## 当前版本：0.9.2

2026-09-23 更新：合入经审查的上游功能与修复，改进 Word 长文加载、编辑与分页，新增 Word／表格独立编辑器窗口，保留大狗 Agent、草稿、保存和退出保护。0.9.0 用户可通过新版本提示下载后手动安装。**本次轮换 Minisign 发布公钥**，下载节点需按[公钥轮换公告](SIGNING-KEYS.md)更新信任配置。

| 平台                | 下载                                                                                                                                                                                                                                                                                                                      |
| ------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Windows x64         | [EXE 安装器](https://github.com/whdite/bdOffice/releases/download/v0.9.2/BigDogOffice-0.9.2-windows-x64.exe)                                                                                                                                                                                                              |
| Linux x64           | [AppImage](https://github.com/whdite/bdOffice/releases/download/v0.9.2/BigDogOffice-0.9.2-linux-x64.AppImage) · [DEB](https://github.com/whdite/bdOffice/releases/download/v0.9.2/BigDogOffice-0.9.2-linux-x64.deb) · [RPM](https://github.com/whdite/bdOffice/releases/download/v0.9.2/BigDogOffice-0.9.2-linux-x64.rpm) |
| macOS Apple Silicon | [DMG](https://github.com/whdite/bdOffice/releases/download/v0.9.2/BigDogOffice-0.9.2-macos-arm64.dmg) · [ZIP](https://github.com/whdite/bdOffice/releases/download/v0.9.2/BigDogOffice-0.9.2-macos-arm64.zip)                                                                                                             |

Windows 安装器目前没有 Authenticode 发布者证书，可能显示“未知发布者”。macOS 包为 Apple Silicon（ARM64），采用 ad-hoc 签名，未使用 Apple Developer ID 或公证；Intel 包尚未提供。Minisign 用于检查文件来源及完整性，不能替代系统发布者认证。

0.8.6 起，「关于」和顶部新版本提示统一使用本仓库的 Releases。更新需由你主动下载并安装，不静默安装或重启。0.8.5 及更早版本须先手动安装此版本，完成更新通道迁移。

## 在一个工作空间里，从想法到文件

- **起草与整理**：描述主题、提供要点或参考材料，与大狗 AI 一起完成初稿和修改。
- **多种文档**：编辑 DOCX、XLSX、PPTX、PDF、Markdown 和 HTML；不同格式的支持边界以实际编辑器为准。
- **看得见的执行过程**：大狗 Agent 支持工具调用与执行记录，可选择自动执行或逐次审批。
- **按需选择模型**：通过模型设置连接支持的厂商，配置文本与图像生成模型。联网 AI 服务按所选服务商的规则使用。
- **新增 HTML 编辑器**：源代码编辑、沙箱预览、保存，以及 PDF／DOCX 导出。
- **本地办公命令行**：应用自带 `bigdogoffice`，不依赖全局 Node；在「设置 → 集成」中启用命令行入口或安装操作 Skill。安装的是操作指引，不是其他 Agent 本体。
- **更连贯的生成**：长文按流式生成，提前停止可选择采用或丢弃草稿；未采用的草稿不会触发自动保存。

本次发布不包含尚在开发的企业版组织管理、云存储和协同审批功能。HTML 编辑器已支持内嵌 AI 审批；全局 Agent 对 HTML 的支持不等同于完整语义编辑。

## 关于这个仓库

这里是大狗Office 的**公开介绍与安装包发布仓库**，不提供应用源码或构建流程。

安装包、Minisign 签名、SHA-256 校验清单与机器可读版本清单位于 [Releases](https://github.com/whdite/bdOffice/releases)。GitHub 自动显示的 “Source code” 压缩包仅包含本仓库公开文档，不是应用源码。

大狗Office 基于 [GenOffice](https://github.com/genspark-ai/genoffice) 构建，感谢上游项目和开源社区。第三方组件保留各自许可证，详见[致谢](ACKNOWLEDGEMENTS.md)及安装包内的许可声明。
