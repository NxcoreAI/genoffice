# NxcoreAI GenOffice 上游同步与维护指南

本 Fork 为 EverRoom 提供 Apache-2.0 范围内的 Office 编辑 runtime。它应尽量贴近
`genspark-ai/genoffice`，仅保留一层稳定、无产品副作用的 Electron Embed 适配。

## 1. 远端约定

```text
origin    https://github.com/NxcoreAI/genoffice.git
upstream  https://github.com/genspark-ai/genoffice.git
```

首次设置：

```bash
cd /Users/rlacat/projects/Everroom/apps/desktop/vendor/genoffice
git remote add upstream https://github.com/genspark-ai/genoffice.git
git fetch --unshallow origin
git fetch upstream main --tags
```

如果仓库已经不是浅克隆，跳过 `git fetch --unshallow origin`。公司网络需要代理时，
只给当次命令传入 `https_proxy`、`http_proxy` 和 `all_proxy`，不要把个人代理地址写进仓库配置。

## 2. EverRoom 下游补丁

当前补丁栈保持为三个独立意图：

| 补丁 | 目的 | 主要文件 |
| --- | --- | --- |
| Side-effect-free embed entry | 导出 Docs 嵌入 API，不启动独立窗口、单实例或应用生命周期 | `apps/docs/src/main/embed.ts`、`electron.vite.config.ts` |
| Standalone updater isolation | updater 只属于 GenOffice 独立应用，不进入 Embed 主进程闭包 | `apps/docs/src/main/index.ts` |
| EverRoom embed mode | 支持 `hostMode=everroom`，构建时移除 GenOffice AI/Genspark UI 与主进程实现 | `docs-main.ts`、`App.tsx`、renderer bootstrap/styles |

同步时优先保留“补丁意图”，不要机械保留旧代码。如果上游已经原生提供等价的无副作用
Embed API，应删除相应下游补丁，并让 EverRoom 适配上游正式接口。

## 3. 稳定 Embed 契约

EverRoom 目前只依赖 `apps/docs/src/main/embed.ts` 导出的下列能力：

- `configureDocsRuntime`
- `createDocsView(openPath, { hostMode: 'everroom' })`
- `registerDocsIpc`
- `setActiveDocsResolver`
- `setDocsShellWindow`
- `teardownDocsRenderer`
- dirty/close 相关导出

约束：

1. import `embed.js` 不得启动 BrowserWindow、申请单实例锁、注册 updater 或修改 app 生命周期。
2. `GENOFFICE_EMBED_ONLY=1` 必须只生成 Embed 主进程入口。
3. EverRoom 模式不得注册或分发 GenOffice AI Provider、Genspark 登录和 updater 实现。
4. preload/renderer 保持 `sandbox: true`、`contextIsolation: true`、`nodeIntegration: false` 可用。
5. 保存、dirty、关闭保护及 DOCX 字节保真不能因嵌入模式退化。

## 4. 推荐同步流程

不要在 Fork `main` 上直接试合并，也不要重写或强推已经被 EverRoom 引用的历史。

```bash
cd /Users/rlacat/projects/Everroom/apps/desktop/vendor/genoffice

git status --short
git fetch origin
git fetch upstream main --tags
git switch main
git pull --ff-only origin main
git switch -c sync/upstream-YYYYMMDD
git merge --no-ff upstream/main
```

解决冲突后，先检查范围：

```bash
git diff --check
git diff --stat origin/main...HEAD
git log --oneline --decorate origin/main..HEAD
```

高概率冲突文件：

- `apps/docs/electron.vite.config.ts`
- `apps/docs/src/main/docs-main.ts`
- `apps/docs/src/main/index.ts`
- `apps/docs/src/renderer/App.tsx`
- `apps/docs/src/renderer/main.tsx`
- AI Panel 和 Ribbon 入口相关文件

如果上游跨越多个大版本，按版本或日期分段合并，比一次吞下全部变更更容易定位回归。

## 5. Fork 验证门禁

在 submodule 中运行：

```bash
npm ci --ignore-scripts --no-audit --no-fund
npm run typecheck -w @genoffice/docs
npm run test -w @genoffice/docs
GENOFFICE_EMBED_ONLY=1 npm run build -w @genoffice/docs
```

至少手工验证以下 fixture：

- `fixtures/generated/simple.docx`
- `fixtures/generated/kitchen-sink.docx`
- `apps/docs/tests/pagination-corpus/docx/20-cn-short-report.docx`

然后回到 EverRoom 根目录验证最终裁剪产物和 Electron 兼容性：

```bash
cd /Users/rlacat/projects/Everroom
node apps/desktop/scripts/prepare-genoffice-runtime.mjs
pnpm --filter @nxcore/desktop typecheck
pnpm --filter @nxcore/desktop exec electron-vite build
git diff --check
```

启动 `pnpm dev`，从左侧“Office 测试”打开内置 DOCX，检查：

- 文档加载、滚动、编辑、保存与撤销；
- 左侧导航和右侧 Agent Panel 展开/收起时 View bounds；
- 离开/返回 Office 测试页面时隐藏和恢复；
- 窗口缩放、关闭保护、外链和弹窗行为；
- runtime 不含 `ee/`、standalone main、updater、AiPanel 或 AI Provider 网络实现。

## 6. 合并和发布 Fork

验证通过后将同步分支推到 Fork，并通过 PR 合入 `NxcoreAI/genoffice:main`：

```bash
git push -u origin sync/upstream-YYYYMMDD
```

PR 应记录：

- 同步前后的 upstream commit/tag；
- 冲突及解决方式；
- 下游补丁是否删除、替换或新增；
- Docs 测试、Embed 构建、EverRoom 构建及手工 fixture 结果；
- 已知兼容性变化和回滚 commit。

禁止为了同步方便对 Fork `main` 使用 `push --force`。

## 7. 在 EverRoom 更新 submodule

只有 Fork commit 已存在于 `NxcoreAI/genoffice` 远端后，才能更新父仓 gitlink：

```bash
cd /Users/rlacat/projects/Everroom
git -C apps/desktop/vendor/genoffice fetch origin
git -C apps/desktop/vendor/genoffice switch main
git -C apps/desktop/vendor/genoffice pull --ff-only origin main

node apps/desktop/scripts/prepare-genoffice-runtime.mjs
pnpm --filter @nxcore/desktop typecheck
pnpm --filter @nxcore/desktop exec electron-vite build

git add apps/desktop/vendor/genoffice
git commit -m "chore(desktop): update GenOffice runtime"
```

父仓提交说明应写明旧/新 Fork commit、对应 upstream tag/commit 和验证结果。不要提交一个仅存在
于某台开发机的 submodule commit，否则其他开发者和 CI 无法 checkout。

## 8. 回滚

GenOffice runtime 由父仓 gitlink 固定。出现回归时，在 EverRoom 新建回滚提交，把 submodule
指针恢复到上一个已验证 commit，再重新运行准备脚本和桌面构建。不要删除或改写已经发布的
Fork commit；保留它们用于问题复现和二分定位。

## 9. 建议节奏

- 安全修复：确认影响后尽快同步。
- 普通上游更新：每 2–4 周集中同步一次，避免长期积累冲突。
- DOCX 主架构、IPC、Electron 或构建系统大改：单独同步，不与功能开发混在同一 PR。
- XLSX/PPTX/PDF 接入后，为每个格式增加相同的固定 fixture 和产物门禁。

