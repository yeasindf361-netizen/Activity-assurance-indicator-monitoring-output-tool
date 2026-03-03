# Maintenance Guide

## 1. 分支与提交建议
- 分支命名：`feature/REQ-xxx`、`fix/REQ-xxx`、`chore/...`
- 提交规范：`feat/fix/chore/docs: 简述`

## 2. 提交前检查
```powershell
git status --short
git diff --name-only --cached
```

## 3. 发布流程
1. 更新 `CHANGELOG.md`
2. 提交并打 tag（如 `v1.0.1`）
3. `git push origin <branch> --tags`
4. 等待 GitHub Actions 自动生成 Release

## 4. 回滚
- 按 tag 回滚：`git checkout <old-tag>`
- 或按 commit 回滚：`git revert <commit-id>`

