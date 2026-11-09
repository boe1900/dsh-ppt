# dsh-ppt

从 [dataelement/dsh-desktop](https://github.com/dataelement/dsh-desktop) 提取的独立 **DSH / DeepSeek Harness 插件**。一个包包含 PPT 按钮、模板选择、PPTD 编写与校验、预览和可编辑 PPTX 导出，无需 Electron 或原 Desktop 仓库。

- 16 套模板、192 个版式，保留英文/中文 PPTD 示例和本地预览图。
- 提供 DeepSeek Harness `0.2.1-rc.5` 原生 **PPT 预设**；PPT 工具、Skill 和基础工作工具都随 preset 挂载。
- 支持上传 PPTX 新建自定义模板、逐页预览、保存、编辑和删除，保存后可在后续会话复用。
- 保留原有文件授权、路径限制、写入哈希检查、版本记录和导出校验。
- 提供 `dsh-pptd` 命令行和 `@cola1900/dsh-ppt/pptd` JavaScript API。

## 安装

需要 Node.js `^22.19.0 || >=24.0.0`，以及 DeepSeek Harness Web / Desktop。DSH peer 依赖声明为 `^0.2.1-rc.5`，开发依赖固定为 `0.2.1-rc.5`。本次核心升级使用宿主的工作目录和沙箱策略服务，不再声明兼容旧的 `0.1.7` / `0.2.0` 接口。

此版本只支持新版 Agent Preset 架构，不兼容旧版 `.agent-presets` 目录加载方式。

验证基线为 `0.2.1-rc.5`。执行 `npm ci && npm run check` 可复现预设接入、模板上下文、可编辑 PPTX 导出和自定义模板生命周期测试。

从 [公共 npm 包](https://www.npmjs.com/package/@cola1900/dsh-ppt) 直接安装到目标 DSH profile：

```sh
dsh plugin --profile web add @cola1900/dsh-ppt
```

把 `web` 替换为实际 profile，然后重启该 profile。新版原生预设选择器会出现 **PPT**；选择后，输入框下方显示模板面板，在首次发送前选择模板，再描述需求。首次发送后预设和模板固定，新会话可以重新选择。预设由 bundle patch 注册，不会复制或修改 `$DSH_HOME/.agent-presets` 文件。

界面只使用官方 `conversation.input.dock` 插槽，只设置插件自身样式，不修改原 DSH。模板名称继续在卡片内显示省略号，输入框底部不再显示 `PPT · 模板名` 控件。

新版 DSH 中，主插件只负责路由、预览和持久化状态；PPT 工具、提示词和 Skill 由 `ppt` preset 子插件启用。不要再安装旧的 `dsh-ppt-composer`。

## 自定义模板

选择 **PPT** 预设，在现有模板面板中点击 **新建自定义模板**，上传 `.pptx` 文件。预览转换后的页面，填写名称并保存。保存后自动选中，可在 **自定义** 分类中重新选择、编辑名称与描述或删除。

自定义模板保存在当前插件数据目录的 `personal-templates/`，同一 profile 下跨会话和重启复用。最多保存 100 套模板，单次导入最多 200 页。生成任务通过 `ppt_template_create_project` 创建独立工程副本，修改副本不会覆盖模板原稿，删除模板也会保留已生成文件。

模板预览使用随依赖安装的 LibreOffice Kit。复杂 Office 对象和缺失字体可能影响转换效果，应在保存前检查预览。不支持的资源会保留原位置的占位图。

## 命令行

使用独立命令行时，从 npm 全局安装：

```sh
npm install -g @cola1900/dsh-ppt
dsh-pptd --help
dsh-pptd check /path/to/deck --json
dsh-pptd render /path/to/deck -o /path/to/output.pptx --json
```

把 `/path/to/deck` 替换为包含 `deck.pptd` 的 PPTD 项目目录。

英文示例目录为 `source/`，中文为 `source-zh/`；预览语言不限制最终文稿语言。导出的文字、图形、表格与图表保留原生可编辑对象。机器上的字体会影响实际显示效果，插件不分发字体文件。

## 开发

`lib/` 是上游维护的 JavaScript 运行时代码；`src/client.js` 是本项目输入框客户端，`src/personal-template-manager.js` 提供自定义模板交互。`presets/ppt.patch.yml` 是 Harness `0.2.1` 系列的原生 Agent Preset 声明，安装 bundle 后由 profile 直接加载。完整原始 TypeScript 源码不在上游提取基线中。

`skills/dsh-ppt/references/` 保存模板元数据、设计说明、双语源文件和预览。修改元数据或预览后执行 `npm run build`，生成模板目录、预览白名单和客户端预览映射。构建使用已随包保存的预览，不依赖网络或原桌面工程。修改版式后可用 `dsh-pptd screenshot` 重新生成图片，再更新对应的 JPG 预览。

默认数据目录使用 `dshHomePath('kimi-ppt')`。可以在 Cordis 插件配置中设置绝对路径 `root`；`maxSlides`、`maxDecksPerSession`、`maxActivities` 保留原有上限配置。`pptSkillRoot` 或 `DSH_PPT_SKILL_ROOT` 可覆盖技能目录。

与上游源码对照可运行 `node scripts/check-upstream.mjs /path/to/dsh-desktop`。该命令逐文件比对核心运行模块和 Skill 文档；预设接入、客户端及公开 SDK 预览适配由集成测试验证。

## 自动构建与发布

推送到 `main` 或提交 Pull Request 时，GitHub Actions 自动安装锁定依赖、构建、运行测试，并上传 npm 安装包。构建还会检查已提交的生成文件与源码是否一致。

正式发布使用 npm Trusted Publishing，无需在 GitHub 保存 npm token。在 npm 包设置中添加 GitHub Actions 信任发布者，仓库填写 `boe1900/dsh-ppt`，工作流文件填写 `publish.yml`，环境名留空，并允许直接发布。

发布步骤如下。

1. 执行 `npm version minor --no-git-tag-version`，或按改动选择 `patch`，提交并推送版本改动。
2. 为该提交创建匹配版本的标签，例如 `v0.4.0`，再发布对应的 GitHub Release。
3. `Publish to npm` 工作流重新构建和测试，将同一份安装包通过 OIDC 发布到 npm，并附带来源证明。工作流保留安装包供下载。

标签必须与 `package.json` 版本一致。预发布 Release 不会发布到 npm `latest`。若发布失败，在解决原因后重跑失败任务，或手动运行工作流并选择原版本标签。

提取版本和具体改动见 [UPSTREAM.md](UPSTREAM.md)。许可证与模板来源见 [LICENSE](LICENSE)、[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) 及 `licenses/`。
