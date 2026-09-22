# 0.9.2 发布签名公钥轮换

2026-09-23 起，0.9.2 使用新的 Minisign 发布签名身份。原私钥无法恢复，项目维护者已批准创建新身份；没有证据表明原私钥泄露。旧公钥保留用于校验 0.9.0 及此前的已发布文件，历史安装包与签名不变。

## 当前公钥（0.9.2 起）

```text
RWS8xW2k8Tr9YO19U6DabeC7fL4xNDFDN8jN29tTP4etDVffqhOqeoxd
```

- Key ID：`60FD3AF1A46DC5BC`
- 公钥解码后 42 字节的 SHA-256：`3e3ef3d0a85c3314aef68f64e1f5c59ca0a5550a93e0f01d9310523875c4c799`

## 历史公钥（截至 0.9.0）

```text
RWQzJu6+Jww+ZBxyLybA+HIpfIIp+DBGe0/dA29WMOaZW2F/w6UPTxYC
```

- Key ID：`643E0C27BEEE2633`

## 下载节点迁移

原私钥不可用，因此本次轮换没有旧身份的交叉签名。下载节点管理员须通过已信任的项目发布渠道审查本公告，再主动更新固定公钥。不要直接采用未验证清单或同一下载目录里提供的公钥。

0.9.2 的安装包、`release-manifest.json` 和 `SHA256SUMS.txt` 均使用新身份签名。先验证清单签名，再核对平台、架构、大小、SHA-256 与安装包签名。新公钥校验失败时不可自动退回旧公钥接受新版文件；历史公钥仅适用于历史版本。

门户代码中的信任配置更新不代表线上下载节点已部署。维护者应在更新节点配置后重新同步并验证文件。Minisign 文件签名不替代 Windows Authenticode、Apple Developer ID 或 Apple 公证。
