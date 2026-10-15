# gitcode

项目说明（待补充）。

## 开始使用

```powershell
git clone https://github.com/K934/gitcode.git
cd gitcode
```

## 团队协作约定

- `main` 为稳定分支，**禁止直接提交/推送**，一切改动走 Pull Request。
- 功能分支命名：`feat/xxx`、`fix/xxx`、`chore/xxx`、`docs/xxx`。
- 提交信息遵循 Conventional Commits：`feat: ...`、`fix: ...`、`docs: ...`。
- 合并方式统一使用 **Squash and merge**。
- PR 需 CI 通过 + 至少 1 人 Approve 才能合并。

## 日常流程

```powershell
git switch main
git pull --rebase
git switch -c feat/your-task

# ... 写代码 ...

git add -A
git commit -m "feat: your change"
git push -u origin feat/your-task
# 然后在 GitHub 上开 PR
```

合并后清理：

```powershell
git switch main
git pull --rebase
git branch -d feat/your-task
```
