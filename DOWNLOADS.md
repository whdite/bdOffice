# 下载与完整性校验

## 用户下载

从 [最新 Release](https://github.com/whdite/bdOffice/releases/latest) 选择与系统、CPU 架构匹配的安装包。国内用户也可使用[下载门户](https://bdoffice.deepthink-x.com/updates)。

每个安装包附带同名 `.minisig` 文件；`SHA256SUMS.txt` 和 `release-manifest.json` 也有各自签名。请先验证文件再安装。

固定 Minisign 公钥：

```text
RWQzJu6+Jww+ZBxyLybA+HIpfIIp+DBGe0/dA29WMOaZW2F/w6UPTxYC
```

使用 [Minisign 官方工具](https://jedisct1.github.io/minisign/) 验证，例如：

```sh
minisign -Vm BigDogOffice-0.8.6-windows-x64.exe -P RWQzJu6+Jww+ZBxyLybA+HIpfIIp+DBGe0/dA29WMOaZW2F/w6UPTxYC
```

签名验证只证明文件与签名相符；不要仅信任同一下载目录中可被一同替换的公钥。Minisign 不替代 Windows Authenticode、Apple Developer ID 或 Apple 公证。

## 下载服务器入口

- 最新版本元数据：`https://api.github.com/repos/whdite/bdOffice/releases/latest`
- 最新安装包清单：`https://github.com/whdite/bdOffice/releases/latest/download/release-manifest.json`
- 清单签名：在以上清单地址后追加 `.minisig`。
- 固定版本资源：`https://github.com/whdite/bdOffice/releases/download/v0.8.6/<资产文件名>`。

服务器应先验证清单签名，然后读取 `version` 和 `artifacts`。按 `platform`、`arch`、`format` 选择资源，使用清单中的 `url` 下载，并检查 `bytes`、`sha256` 和对应的 `signatureUrl`。清单不列出的平台表示该版本尚无可用安装包，不应自动回退成其他系统或旧版本。

公开下载不需要 SSH 密钥或 GitHub Token；匿名 API 访问有速率限制。建议缓存已验证清单，并在版本号发生变化时拉取新包。不要使用 Git clone、仓库自动生成的 Source code 压缩包或私有 Actions 构建产物作为用户安装入口。

0.8.6 起桌面应用在启动后和运行期间检查本仓库，使用顶部提示与「关于」检查更新。
更新需要用户主动下载并安装，不静默下载、安装或重启。0.8.5 及更早版本须先手动安装 0.8.6，才能使用新通道。

macOS 安装包目前只提供 Apple Silicon（ARM64），使用 ad-hoc 签名，未使用 Apple Developer ID 或 Apple 公证；可能被 Gatekeeper 拦截。不要将此包用于 Intel Mac。Windows 尚无 Authenticode 证书；文件签名验证不消除系统的未知发布者提示。
