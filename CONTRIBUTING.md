# 团队 Git / GitHub 协作操作手册

> 适用对象：**从没接触过 Git 和 GitHub 的团队成员**
> 本仓库：https://github.com/K934/gitcode
> 负责人：K934
>
> 📄 **需要打印或离线阅读？** 仓库里提供了 Word 版本：[`GIT-协作手册.docx`](./GIT-协作手册.docx)（可直接下载/打印）。
> 该 Word 文档由本手册自动生成，**请勿直接编辑它**——要改内容请改 Markdown 后重新生成，方法见文末附录。

---

## 目录

- [一、先搞懂三件事](#一先搞懂三件事)
- [二、名词速查表](#二名词速查表)
- [三、第一次配置（只做一次，约 20 分钟）](#三第一次配置只做一次约-20-分钟)
- [四、最常用的一套命令（背下来就够用了）](#四最常用的一套命令背下来就够用了)
- [五、完整走一遍：提交你的第一个文件](#五完整走一遍提交你的第一个文件)
- [六、团队规矩（本仓库的硬性规则）](#六团队规矩本仓库的硬性规则)
- [七、常见报错翻译表](#七常见报错翻译表)
- [八、出错了怎么办](#八出错了怎么办)
- [九、速查卡](#九速查卡)

---

## 一、先搞懂三件事

**1. Git 是什么？**
Git 是装在你电脑上的**版本管理工具**。它像一个无限次的「存档」按钮：你随时可以把代码存一个档，以后能翻看任何一次存档、能对比改动、能回退到以前的样子。

**2. GitHub 是什么？**
GitHub 是一个**网站**，用来存放 Git 的存档。Git 管你本机的存档，GitHub 让全组人的存档汇到一个地方，别人能看到你的改动、给你提意见。

**3. 它们怎么配合？**

```
你的电脑                           GitHub 网站
┌──────────────┐                  ┌──────────────┐
│  工作区       │                  │              │
│  （你编辑的    │   git push →     │   远程仓库    │
│    文件）      │   ───────────→   │   (origin)   │
│              │                  │              │
│    ↓ git add │   ← git pull     │              │
│    ↓ commit  │   ────────────   │              │
│  本地仓库     │                  │              │
└──────────────┘                  └──────────────┘
```

- **`git push`（推）**：把你的存档上传到 GitHub
- **`git pull`（拉）**：把 GitHub 上的新内容下载到你电脑

> ⚠️ 关键认知：**你的改动在 `push` 之前，别人是看不到的。**

---

## 二、名词速查表

| 名词 | 大白话解释 |
|---|---|
| **仓库 / repo** | 一个被 Git 管理的项目文件夹 |
| **远程仓库 / origin** | GitHub 上那份项目；`origin` 是它默认的代号 |
| **本地仓库** | 你电脑上那份 |
| **工作区** | 你实际在编辑的文件 |
| **提交 / commit** | 一次「存档」，带一条说明 |
| **暂存 / stage** | 挑选哪些改动要放进这次存档（`git add`） |
| **分支 / branch** | 一条独立的开发线，互不干扰 |
| **`main`** | 主分支，最稳定的那条；**本仓库禁止直接改动** |
| **推送 / push** | 把本地提交上传到远程 |
| **拉取 / pull** | 把远程的新提交下载到本地 |
| **克隆 / clone** | 第一次把整个远程仓库下载到本地 |
| **PR / Pull Request** | 「合并请求」：请别人审阅并把你分支的改动并入 `main` |
| **Review / 评审** | 别人看你的 PR 并给意见 |
| **Merge / 合并** | 把 PR 正式并入 `main` |
| **冲突 / conflict** | 两个人改了同一个文件的同一处，Git 无法自动决定听谁的 |
| **CI** | 自动检查：你推送后机器人自动跑一遍检查，通过才允许合并 |

---

## 三、第一次配置（只做一次，约 20 分钟）

### 步骤 1：注册 GitHub 账号

1. 打开 https://github.com/signup
2. 用**工作邮箱**注册
3. 用户名建议用姓名拼音，例如 `zhangsan-dev`

### 步骤 2：开启双重验证（2FA）

GitHub 要求必须开启，否则很多操作会被限制。

1. 打开 https://github.com/settings/security
2. 点 **Enable two-factor authentication**
3. 用手机装 **Authenticator**（Google Authenticator / Microsoft Authenticator 都行）扫码
4. **把恢复码抄下来存好**（手机丢了就靠它）

### 步骤 3：安装工具

打开 **PowerShell**（开始菜单搜索 `PowerShell`），依次执行：

```powershell
winget install Git.Git
winget install GitHub.cli
```

**装完后必须关掉 PowerShell 重新打开一次**，否则命令找不到。

验证安装成功：

```powershell
git --version      # 应显示 git version 2.x.x
gh --version       # 应显示 gh version 2.x.x
```

### 步骤 4：配置你的身份

> ⚠️ **这一步绝不能跳过。** 没配的话，你提交的代码不会关联到你的账号，别人不知道是谁改的。

把下面两行的内容换成你自己的（邮箱用注册 GitHub 的那个）：

```powershell
git config --global user.name "张三"
git config --global user.email "zhangsan@example.com"
```

检查是否配置成功：

```powershell
git config --global user.name
git config --global user.email
```

### 步骤 5：登录 GitHub

```powershell
gh auth login
```

按提示选择（用方向键选择，回车确认）：

1. 选 **GitHub.com**
2. 选 **HTTPS**
3. 选 **Login with a web browser**
4. 屏幕会显示一个 **一次性代码**（形如 `A1B2-C3D4`），**先抄下来**
5. 按回车会自动打开浏览器，粘贴那个代码，点授权
6. 回到 PowerShell，看到 `✓ Logged in as 你的用户名` 就成功了

### 步骤 6：克隆仓库到本地

```powershell
# 先进入你想放项目的目录，例如 D 盘
cd D:\

git clone https://github.com/K934/gitcode.git

# 进入项目文件夹
cd gitcode
```

> 本仓库是公开的，这一步**不需要密码或令牌**。

### 步骤 7：告诉负责人你的用户名

在 GitHub 上点右上角头像，你的用户名就在那里。把**用户名**（不是邮箱）发给负责人 `K934`，让他把你加为协作者。

**收到邀请后**：打开 https://github.com/notifications ，点 **Accept invitation**。不点的话你没有推送权限。

### ✅ 首次配置检查清单

- [ ] GitHub 账号已注册
- [ ] 2FA 已开启，恢复码已保存
- [ ] `git --version` 和 `gh --version` 都有输出
- [ ] `user.name` 和 `user.email` 已配置
- [ ] `gh auth login` 显示已登录
- [ ] 仓库已克隆到本地
- [ ] 用户名已告知负责人，邀请已接受

---

## 四、最常用的一套命令（背下来就够用了）

**日常工作，99% 的时间只需要这 8 条：**

```powershell
git switch main             # 1. 切到主分支
git pull --rebase           # 2. 拉取最新代码（开工前必做！）
git switch -c feat/我的任务  # 3. 建自己的分支
                            # 4. ... 写你的代码 ...
git status                  # 5. 看看改了哪些文件
git add 文件名               # 6. 把改动加入本次提交
git commit -m "feat: 说明"   # 7. 提交（存档）
git push -u origin feat/我的任务  # 8. 推送到 GitHub
```

然后去网页上开 PR。

> 💡 **为什么每次开工前都要 `git pull --rebase`？**
> 因为你的同事可能刚往 `main` 推了新代码。你不拉，就是基于旧代码开工，越晚同步越容易撞冲突。

---

## 五、完整走一遍：提交你的第一个文件

下面用真实例子把整个流程走完。假设你要新增一个 `hello.txt`。

### 第 1 步：同步主干

```powershell
git switch main
git pull --rebase
```

期望输出类似：

```
Already up to date.
```
或
```
Fast-forward
 ... 
```

### 第 2 步：创建自己的分支

分支名用 `类型/简短描述`：

```powershell
git switch -c feat/add-hello
```

**类型前缀约定：**

| 前缀 | 用途 |
|---|---|
| `feat/` | 新功能 |
| `fix/` | 修 bug |
| `docs/` | 文档 |
| `chore/` | 杂项（配置、构建） |
| `refactor/` | 重构 |

### 第 3 步：写你的文件

用记事本、VS Code 随便什么工具，在项目文件夹里新建 `hello.txt`，写点内容。

### 第 4 步：查看改了什么

```powershell
git status
```

期望输出：

```
On branch feat/add-hello
Untracked files:
  (use "git add <file>..." to include in what will be committed)
	hello.txt

nothing added to commit but untracked files present
```

`Untracked` = Git 发现了这个新文件，但还没纳入管理。

### 第 5 步：加入本次提交

```powershell
git add hello.txt
```

再查一次状态：

```powershell
git status
```

现在显示的是 `Changes to be committed:` 和绿色的 `new file: hello.txt`。

> ⚠️ **不要用 `git add -A` 或 `git add .`**
> 它们会把工作区里**所有**文件都加进来，包括你不小心放进去的密码文件、临时文件。**养成指定文件名的习惯。**

### 第 6 步：提交（存档）

```powershell
git commit -m "feat: add hello.txt"
```

**提交信息规范**（本仓库要求）：

```
类型: 简短说明
```

- 类型用上面的前缀（`feat` / `fix` / `docs` / `chore` / `refactor`）
- 说明用一句话讲清做了什么
- 好的例子：`feat: 增加用户登录页面`、`fix: 修正金额计算错误`
- 不好的例子：`update`、`改了一下`、`111`

### 第 7 步：推送到 GitHub

```powershell
git push -u origin feat/add-hello
```

期望输出：

```
remote: Create a pull request for 'feat/add-hello' on GitHub by visiting:
remote:      https://github.com/K934/gitcode/pull/new/feat/add-hello
```

> `-u` 的作用：建立本地分支和远程分支的关联。**下一次这条分支上直接写 `git push` 就行。**

### 第 8 步：在网页上创建 PR

1. 打开 https://github.com/K934/gitcode
2. 页面顶部会出现黄色提示条，点 **Compare & pull request**
   （如果没有，点 **Pull requests** 标签 → **New pull request**，选择你的分支）
3. **正文会自动填充 PR 模板**，按模板把每一项都填上：
   - 改动内容：一句话说明
   - 关联 Issue：没有就写「无」
   - 改动类型：勾选对应项
   - 影响范围：没有就写「无」
   - 自测情况：写清楚你怎么验证的，**不要只写「测过了」**
   - 自查清单：逐条勾选
4. **标题**用和提交信息一样的规范，例如 `feat: add hello.txt`
5. 点 **Create pull request**

### 第 9 步：等 CI 检查

创建 PR 后，页面下方会出现检查状态：

- 🟡 黄色圆点 = 正在跑
- ✅ 绿色对勾 = 通过
- ❌ 红色叉 = 失败

**本仓库规定：CI 不通过就不能合并。** 如果红了，点 **Details** 看日志，改好代码后重新 `git add` → `commit` → `push`，CI 会自动重跑。

### 第 10 步：请人评审并合并

1. 在 PR 页面右侧 **Reviewers** 里选一位同事
2. 对方看过后会点 **Approve**
3. Approve 后点 **Squash and merge** → **Confirm**

> **合并方式必须是 Squash and merge**，这是本仓库的统一约定。

### 第 11 步：回来清理

```powershell
git switch main
git pull --rebase

# 删除已经合并的本地分支
git branch -D feat/add-hello
```

> ⚠️ 这里用 **`-D`（大写）**，不是 `-d`。
> 因为 Squash 合并会生成一个全新的提交，Git 认不出亲缘关系，用 `-d` 会报错说「分支未合并」——**这是正常的，不是出错**。

### ✅ 一次完整流程检查清单

- [ ] 开工前 `git pull --rebase` 了
- [ ] 在自己的分支上干活，没碰 `main`
- [ ] `git add` 指定了文件名
- [ ] 提交信息符合 `类型: 说明` 规范
- [ ] `git push` 成功
- [ ] PR 模板每一项都填了
- [ ] CI 变绿
- [ ] 有人 Approve 后 Squash and merge
- [ ] 回到 main 同步并删除了本地分支

---

## 六、团队规矩（本仓库的硬性规则）

这些规则由 GitHub 在服务器端强制执行，**不是靠自觉**。

### 🚫 绝对不要做的事

| 禁止 | 原因 |
|---|---|
| 直接往 `main` 推代码 | 服务器会拒绝，报 `GH006: Protected branch update failed` |
| 对 `main` 使用 `git push --force` | 会抹掉别人的提交历史 |
| 在 `main` 分支上直接改代码 | 同上，会被拒绝 |
| 提交密码、密钥、`.env` 文件 | 本仓库是**公开**的，一旦提交等于对全世界公开，删了也没用，只能重新生成密钥 |
| 用 `git add -A` | 容易把无关文件和密钥一起提交 |

### ✅ 必须做的事

| 要求 | 说明 |
|---|---|
| 开工前先 `git pull --rebase` | 避免基于旧代码开发 |
| 一切改动走 PR | 服务器强制 |
| 一个分支只做一件事 | PR 越小越好评审 |
| 提交信息用 `类型: 说明` | 便于追踪 |
| 合并用 Squash and merge | 统一约定 |
| CI 必须通过 | 服务器强制 |

### 为什么有这些规则？

不是为了限制你，是为了防止**事故**：

- 半夜改 bug 手滑把半成品推上主干 → 全组拉到坏代码
- 有人 `--force` 覆盖历史 → 别人的工作瞬间消失
- 密钥误提交 → 被爬虫秒抓，只能轮换密钥

---

## 七、常见报错翻译表

| 英文报错 | 中文含义 | 怎么办 |
|---|---|---|
| `fatal: not a git repository` | 当前目录不是 Git 仓库 | 用 `cd` 进入项目文件夹 |
| `error: GH006: Protected branch update failed` | 你试图直接推 `main` | 新建分支，走 PR |
| `! [rejected] ... (fetch first)` | 远程有你没有的提交 | 先 `git pull --rebase` 再 push |
| `! [rejected] ... (non-fast-forward)` | 同上，本地落后于远程 | 同上 |
| `CONFLICT (content): Merge conflict in 文件` | 冲突：你和别人改了同一处 | 见第八章 |
| `error: The branch 'xxx' is not fully merged` | 分支未被识别为已合并 | 确认 PR 已合并后，用 `git branch -D xxx` |
| `nothing to commit, working tree clean` | 没有新改动可提交 | 你还没改文件，或已经提交过了 |
| `fatal: could not read Username for 'https://github.com'` | 没登录 / 凭据失效 | 重新 `gh auth login` |
| `Permission denied (publickey)` | SSH 密钥问题 | 本手册统一用 HTTPS，重新 `gh auth login` 选 HTTPS |
| `Please tell me who you are` | 没配身份 | 执行第 3 章步骤 4 |
| `warning: LF will be replaced by CRLF` | Windows 换行符提示 | **可以忽略**，不是错误 |
| `Everything up-to-date` | 没有新内容需要推送 | 正常，可能你还没提交 |

---

## 八、出错了怎么办

### 场景 1：改乱了想撤销（还没提交）

```powershell
# 看看具体改了什么
git diff

# 放弃某个文件的所有改动，恢复原样
git restore 文件名

# 放弃所有未提交的改动（谨慎！不可恢复）
git restore .
```

### 场景 2：提交信息写错了（还没推送）

```powershell
git commit --amend -m "正确的新信息"
```

### 场景 3：已经推送了，想撤销最后一次提交

```powershell
# 撤销提交但保留改动（改动还在工作区，可以重新改后提交）
git reset --soft HEAD~1
```

### 场景 4：遇到冲突（CONFLICT）

这通常发生在 `git pull` 或合并时。别慌，按下面走：

**第 1 步**：`git status` 会列出冲突的文件：

```
Unmerged paths:
  (use "git add <file>..." to mark resolution)
	both modified:   README.md
```

**第 2 步**：用编辑器打开那个文件，会看到这样的标记：

```
<<<<<<< HEAD
这是你本地的内容
=======
这是远程/别人的内容
>>>>>>> origin/main
```

**第 3 步**：手动决定最终要保留什么，然后**把 `<<<<<<<`、`=======`、`>>>>>>>` 这三行标记全部删掉**，只留下你想要的内容。

**第 4 步**：标记为已解决并完成：

```powershell
git add README.md
git rebase --continue     # 如果是 rebase 过程中
# 或
git commit                # 如果是 merge 过程中
```

**第 5 步**：如果搞砸了想重来：

```powershell
git rebase --abort        # 放弃 rebase，回到操作前
# 或
git merge --abort         # 放弃 merge
```

> 💡 **实在搞不定就求助。** 不要在不懂的情况下乱试 `git reset --hard`，它会永久丢弃你的改动。

### 场景 5：想看看别人最近做了什么

```powershell
git log --oneline -10        # 最近 10 条提交
git log --oneline --graph    # 带分支图
git show 提交编号             # 看某次提交的具体内容
```

也可以直接在网页上看：https://github.com/K934/gitcode/commits/main

### 场景 6：本地乱成一团，想重新开始

最稳妥的办法：**重新克隆一份**（确认本地没有未推送的重要改动之后）

```powershell
cd ..
git clone https://github.com/K934/gitcode.git gitcode-新
```

---

## 九、速查卡

**把这一页打印出来贴在显示器旁边。**

```powershell
# ===== 开工前 =====
git switch main
git pull --rebase
git switch -c feat/我的任务

# ===== 干活中 =====
git status                    # 看改了哪些文件
git diff                      # 看具体改了什么
git add 文件名                 # 加入本次提交（指定文件名！）
git commit -m "feat: 说明"     # 提交

# ===== 收工 =====
git push -u origin feat/我的任务
# → 去网页开 PR → 等 CI 绿 → 找人 Approve → Squash and merge

# ===== 合并后 =====
git switch main
git pull --rebase
git branch -D feat/我的任务

# ===== 救命命令 =====
git status                    # 迷茫时先看这个
git diff                      # 看改了什么
git log --oneline -10         # 看历史
git restore 文件名              # 放弃单个文件的改动
git rebase --abort            # 放弃出错的 rebase
```

---

## 附：负责人需要做的事

### 添加协作者

```powershell
gh api -X PUT repos/K934/gitcode/collaborators/对方用户名 -f permission=push
```

对方必须在 https://github.com/notifications 接受邀请才生效。

### 队友到位后，开启强制评审

当前设置为 Approve 人数 = 0（单人开发阶段的临时设置）。**有队友后应改回 1**：

```powershell
gh api -X PATCH repos/K934/gitcode/branches/main/protection/required_pull_request_reviews -F required_approving_review_count=1
```

### 待办：把 CI 换成真实检查

当前 `.github/workflows/ci.yml` 只做了占位检查（验证 README 和 .gitignore 存在）。技术栈确定后，把里面注释掉的 `npm ci` / `npm test` / `npm run lint` 换成真实命令。

---

## 需要帮助时

1. **先查第七章的报错翻译表**
2. **再用 `git status` 看当前状态**，把输出发给负责人
3. **报错信息请完整复制**，不要只描述「报错了」
4. 实在搞不定，**不要乱试 `--hard` 类命令**，直接求助

---

## 附录：Word 版本是怎么生成的

仓库里的 `GIT-协作手册.docx` **不是手工维护的**，它由本 Markdown 文件自动转换而来。转换脚本是 `tools/md2docx.ps1`，**不依赖 pandoc、Python 或 Microsoft Word**，只用 Windows 自带的 .NET 组件构造 OOXML 包。

**修改文档内容的正确姿势：**

```powershell
# 1. 改 Markdown（这是唯一的内容来源）
# 2. 重新生成 Word
& .\tools\md2docx.ps1 -InputPath .\CONTRIBUTING.md -OutputPath '.\GIT-协作手册.docx' -Title '团队 Git / GitHub 协作操作手册' -Author 'K934'

# 3. 两个文件一起提交
git add CONTRIBUTING.md 'GIT-协作手册.docx'
git commit -m "docs: update collaboration manual"
```

> ⚠️ **不要直接编辑 .docx 文件。** 它是生成物，下次重新生成时你的手改会被覆盖。
>
> ⚠️ 脚本参数名是 `-InputPath` / `-OutputPath`，**不是** `-Input` / `-Output`——因为 `$Input` 是 PowerShell 的保留自动变量，会与参数冲突导致绑定失败。

---

*最后更新：2026-10-15*
