# 地算通 · 版本更新源

这个仓库用来给桌面软件「地算通」提供**版本更新通知**，用 GitHub Pages 免费托管，不需要服务器。

| 文件 | 作用 |
|---|---|
| `version.json` | 软件启动时读取的版本信息（版本号、日期、更新日志、下载地址） |
| `index.html` | 用户访问站点时看到的下载页 |
| `.nojekyll` | 让 Pages 原样发布文件，跳过 Jekyll 构建 |
| `deploy.py` | 一键部署脚本（传文件、开 Pages、建 Release、传 exe、验证） |

- 版本信息：https://yzq0906.github.io/DST/version.json
- 下载页：https://yzq0906.github.io/DST/
- 历史版本：https://github.com/yzq0906/DST/releases

## 以后发新版本，只需要改 `version.json`

1. 把新 exe 打包好，上传成对应 tag 的 Release 资产（例如 `V2.9.0` → `地算通V2.9.0.exe`）。
2. 编辑本仓库根目录的 `version.json`：
   - `version` 改成新版本号（必须比用户端的大，三段式如 `V2.9.0`）
   - `date` 改成发布日期
   - `notes` 改成这一版的更新条目
   - `url` 改成新版 exe 的下载地址
3. 顺手更新 `index.html` 里的版本号、下载链接和更新列表。
4. 提交后约 1 分钟生效，用户启动软件（或点「帮助 → 检查更新」）就会收到提示。

> `version.json` 里的版本号**不要低于**已发布安装的软件版本，否则用户端不会弹更新提示。

## 一键部署脚本

**最省事：直接双击 / 运行，不用创建任何 token。**

```powershell
cd "C:\Users\82235\WorkBuddy\2026-11-04-11-10-29"
python deploy\deploy.py
```

脚本会自动打开 GitHub 设备授权页，并打印一个 8 位验证码。在网页里输入验证码、点一下 **Authorize**，剩下的（传文件、开 Pages、建 Release、传 exe、验证）全部自动完成。token 只在内存里，不落盘、不显示。

如果已经有 PAT，也可以 `$env:GITHUB_TOKEN="ghp_xxx"` 或写进 `deploy/.token`，这样跳过授权页直接部署。**`.token` 千万不要提交到仓库。**

脚本会依次完成：上传站点文件 → 开启 Pages → 创建 Release 并上传 `../dist/地算通V2.8.0.exe` → 等待 Pages 生效并实测 `/DST/`、`/dst/` 哪个路径可用。可重复运行。

Token 权限：classic PAT 勾 `repo`；或 fine-grained PAT 给本仓库的 **Contents / Pages / Releases** 读写权限。

## 软件端配置

源码顶部：

```python
VERSION_URL = 'https://yzq0906.github.io/DST/version.json'
VERSION_URL_FALLBACK = [
    'https://yzq0906.github.io/dst/version.json',
    'https://raw.githubusercontent.com/yzq0906/DST/main/version.json',
]
```

改完 `VERSION_URL` 需要重新打包 exe，用户端才会指向新的更新源。
