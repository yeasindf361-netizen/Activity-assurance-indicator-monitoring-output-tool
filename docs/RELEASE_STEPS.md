# Release Steps

## 命名规范
- 版本号：`vMAJOR.MINOR.PATCH`
- Release 资产：`Event-Support-Tool-<tag>-win64.zip`

## 本地命令（手动）
```powershell
git init
git add .
git commit -m "feat: initialize standard project structure"
git branch -M main
git remote add origin <YOUR_GITHUB_REPO_URL>
git push -u origin main
git tag v1.0.0
git push origin v1.0.0
```

## 自动发布（推荐）
- 推送 tag 后，GitHub Actions 自动构建并上传 Release 资产。

