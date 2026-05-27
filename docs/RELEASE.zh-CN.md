# 发布说明

本文档只记录仓库发布卫生和 GitHub Release 的基本流程，不包含任何 API Key 或本机私有配置。

## 生成本地 release

在项目根目录运行：

```powershell
.\build_exe.ps1
```

如需指定 Python 环境：

```powershell
.\build_exe.ps1 -Python "D:\ProgramFiles\Anaconda3\envs\paper2excel\python.exe"
```

脚本会先运行测试，再重新生成 `build/`、`dist/`、`release/Paper2Excel/` 和 `release/Paper2Excel-v0.1.1-windows.zip`。这些目录和 zip 是本机构建产物，不应提交到源码仓库；zip 应作为 GitHub Release 附件上传。

发布前请确认：

- `release/Paper2Excel/Paper2Excel.exe` 能正常启动。
- `release/Paper2Excel-v0.1.1-windows.zip` 解压后能直接运行。
- release 内应包含中文 `README.md` 和英文 `README.en.md`，不再包含 `README.zh-CN.md`。
- `release/Paper2Excel/config.example.json` 只包含示例配置。
- 不上传 `user_config.json`、`.env`、`outputs/`、`logs/` 或任何 API Key。
- `LICENSE` 为 MIT 协议，并包含在源码和 release 中。

## 上传 GitHub 源码

只提交源码、测试、模板、README、`config.example.json`、构建脚本和发布说明等可公开文件。

```powershell
git status
git add .
git commit -m "feat: prepare Paper2Excel release"
git tag v0.1.1
git push
git push origin v0.1.1
```

如果仓库中已经误跟踪了构建产物或本机配置，先用 `git rm --cached` 从索引移除，再提交；不要删除本机仍需保留的文件。

## 创建 GitHub Release

1. 在 GitHub 仓库页面进入 `Releases`，选择 `Draft a new release`。
2. 创建版本标签，例如 `v0.1.1`。
3. 上传脚本生成的 `release/Paper2Excel-v0.1.1-windows.zip` 作为 Release 附件。
4. 不要把 `release/` 目录本身提交到源码仓库。
5. 在说明中写明主要变更、运行方式和配置提醒，不粘贴任何 API Key。
