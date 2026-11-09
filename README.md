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

1. 把新 exe 打包好，上传成对应 tag 的 Release 资产（例如 `V2.9.0` → **`V2.9.0.exe`**）。
2. 编辑本仓库根目录的 `version.json`：
   - `version` 改成新版本号（必须比用户端的大，三段式如 `V2.9.0`）
   - `date` 改成发布日期
   - `notes` 改成这一版的更新条目
   - `url` 改成新版 exe 的下载地址
3. 顺手更新 `index.html` 里的版本号、下载链接和更新列表。
4. 提交后约 1 分钟生效，用户启动软件（或点「帮助 → 检查更新」）就会收到提示。

> `version.json` 里的版本号**不要低于**已发布安装的软件版本，否则用户端不会弹更新提示。

> ⚠️ **Release 资产名务必用纯 ASCII（如 `V2.9.0.exe`），不要用中文。**
> GitHub 上传接口会丢掉中文字符：实测传 `地算通V2.8.0.exe`，线上资产变成 `V2.8.0.exe`，
> 按中文名拼出来的下载直链就是 404。`deploy.py` 里的 `ASSET_NAME` 已经处理了这一点。
> 另外软件端会用公开的 GitHub API 反查该 tag 的真实资产地址，即使 json 里的 `url` 写错也能自动校正，
> 并在更新提示里额外提供「发布页」按钮作为兜底。

## 一键部署脚本

**最省事：直接双击 / 运行，不用创建任何 token。**

```powershell
cd "C:\Users\82235\WorkBuddy\2026-11-04-11-10-29"
python deploy\deploy.py
```

脚本会自动打开**已预填验证码**的 GitHub 授权页，你只需点一下 **Authorize**，剩下的（传文件、开 Pages、建 Release、传 exe、验证）全部自动完成。首次成功后会把授权记到本机 `deploy/.token`，**以后部署不用再点**；删除该文件即可撤销本机授权，也可在 GitHub → Settings → Applications 里撤销。token 不会打印，也不会上传。

如果已经有 PAT，也可以 `$env:GITHUB_TOKEN="ghp_xxx"` 或写进 `deploy/.token`，这样跳过授权页直接部署。**`.token` 千万不要提交到仓库。**

脚本会依次完成：上传站点文件 → 开启 Pages → 创建 Release 并上传 `../dist/地算通V2.8.0.exe`（Release 资产名用 `V2.8.0.exe`）→ 等待 Pages 生效并实测 `/DST/`、`/dst/` 哪个路径可用。可重复运行。发新版时记得同步修改脚本顶部的 `TAG` 与 `ASSET_NAME`。

Token 权限：classic PAT 勾 `repo`；或 fine-grained PAT 给本仓库的 **Contents / Pages / Releases** 读写权限。

### 补传历史版本：`deploy_history.py`

```powershell
python deploy\deploy_history.py
```

同样走设备授权（或 `GITHUB_TOKEN` / `.token`），一次授权会把这些事全部做完：

1. 为 `V2.7.0` / `V2.6.0` / `V2.5.0` 各建一个 Release 并上传对应 exe（资产名 `V2.7.0.exe` 等）；
2. 三个历史版本标记为 **Pre-release**，保证 `/releases/latest` 仍指向正式版 `V2.8.0`；
3. 用本地最新的 `dist/地算通V2.8.0.exe` 覆盖 `V2.8.0` Release 里的旧资产；
4. 同步修正线上 `version.json` 的下载直链、`index.html` 的历史版本列表、`README.md`；
5. 逐个 HEAD 校验下载直链，并确认 `/releases/latest` 没跑偏。

要补更多历史版本（V2.4 / V2.3 / V2.2 / V2.1 的 exe 也都还在 `dist/`），改脚本顶部的 `HISTORY` 列表再跑一次即可，已存在的 Release 会自动复用。

## 软件端配置

源码顶部：

```python
VERSION_URL = 'https://yzq0906.github.io/DST/version.json'
VERSION_URL_FALLBACK = [
    'https://yzq0906.github.io/dst/version.json',
    'https://raw.githubusercontent.com/yzq0906/DST/main/version.json',
]
APP_RELEASES_URL = 'https://github.com/yzq0906/DST/releases/latest'
```

改完 `VERSION_URL` 需要重新打包 exe，用户端才会指向新的更新源。`APP_RELEASES_URL` 是更新提示里「发布页」按钮的目标，一般不用动。

## 当前部署状态（2026-11-09）

| 项 | 结果 |
|---|---|
| Pages | 已开启，main / (root)，`https://yzq0906.github.io/DST/version.json` 返回 200 |
| Release | `V2.8.0` 已创建，资产 `V2.8.0.exe`（29.7 MB） |
| 下载页 | `https://yzq0906.github.io/DST/` 返回 200 |
| 历史版本 | 在线下载页 `https://dicalc-history.app.workbuddy.host/?sharecode=WMLSROEL2A`，含 V2.7.0 / V2.6.0 / V2.5.0 三个 exe，逐条实测 200 |

### 历史版本为什么放在独立下载页

把历史 exe 补进 GitHub Release 需要一次 GitHub 写入授权（设备码 14 分钟内有效，两次都过期未确认），所以先用内置的站点发布能力把三个版本做成了独立下载页，**零授权、已上线可用**。

- 站点源文件：`deploy/history/`（`index.html` + 三个 `V<版本>.exe`）。
- 托管站点要求静态资源也带 `sharecode`，裸路径返回 403；`index.html` 里用一小段脚本把当前地址的 `sharecode` 自动补到下载链接上，重新发布换码也不会失效。
- `index.html`（GitHub Pages 那份）的「历史版本」区块已改为指向这个在线下载页，避免死链；等哪天完成一次授权、把新版页面推上去即可生效。
- 想把三个历史版本也补进 GitHub Release 时，运行 `python deploy\deploy_history.py`（点一次 Authorize，之后会记住授权），脚本会自动建 Release、传 exe、覆盖 V2.8.0 资产并同步线上文件。
