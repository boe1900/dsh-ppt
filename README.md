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

界面使用官方 `conversation.input.dock` 插槽，沿用上游透明模板面板、分类栏、三列卡片和选中边框。面板在输入框下方绝对定位，不占空白会话居中布局的高度；宽度跟随输入卡片和会话视口，滚动区域随窗口高度调整。插件只设置自身样式，不修改原 DSH。模板名称继续在卡片内显示省略号，输入框底部不再显示 `PPT · 模板名` 控件。

新版 DSH 中，主插件只负责路由、预览和持久化状态；PPT 工具、提示词和 Skill 由 `ppt` preset 子插件启用。不要再安装旧的 `dsh-ppt-composer`。

## 自定义模板

选择 **PPT** 预设，在模板面板点击 **新建自定义模板**，在对话里说明需求，无需自己准备或选择模板文件。

1. 在对话里说明用途、风格和品牌要求。例如：“帮我做一套蓝色商务周报模板，中文、16:9，包含封面、章节、图文和数据页，数字标为示例数据。”信息不全时，AI 会简短询问。
2. AI 先确定配色、字体、间距和版式用途，制作代表页，并用长标题或增加内容验证复用。生成后在输入框旁的 **模板预览** 卡片显示结果。点击 **逐页预览** 检查所有页面。
3. 需要调整时，直接在对话里说“配色换成墨绿，增加一页对比版式”。修改后卡片会更新预览；旧版本不能继续保存。
4. 满意后填写名称，点击 **保存到自定义模板**。新建 PPT 对话，在 **自定义** 分类选择使用。保存只加入模板库，当前对话已固定的模板不会被切换。

用户输入可以是自然语言，也可以在对话里附上普通 `.pptx` 作为参考，不要求 `.potx`、特殊占位符或编写代码。AI 根据用途选择有实际作用的版式，不固定页数。内部使用 PPTD 工程保存可编辑页面和图片，并通过工程根目录的 `template-design.json` 保存视觉规范、每页用途、容量、固定元素、替换内容和超量时的处理规则。AI 负责维护这些文件，用户无需手工编写。

制作期间通过对话沟通，不额外显示操作按钮。模板引导尚未发送时，显示原因说明和重试入口。生成后才显示预览和保存卡片，**结束制作** 操作也在这张卡片中，点击后退出并清理未保存的预览。切换会话或重启后可以继续未完成的制作。保存后卡片收起，显示可关闭的保存提示；再次进入会话时不重复提示。自定义模板保存在当前插件数据目录的 `personal-templates/`，同一 profile 下跨会话复用，最多保存 100 套，每套最多 200 页。制作流程不会占普通 PPT 的导出名额。

生成任务先读取模板设计规范和版式说明，再通过 `ppt_template_create_project` 创建独立工程副本。设计说明随工程复制，修改副本不会覆盖模板原稿。可在模板库预览、编辑名称与描述或删除，删除模板也会保留已生成文件。AI 制作的模板直接保存审阅过的原生工程和设计说明，不经导出再导入，示例内容保持原样；新任务按版式容量替换示例内容和数据。旧模板没有设计说明时，仍可使用实际页面和已有结构索引。

预览使用随依赖安装的 LibreOffice Kit。AI 可通过 `ppt_template_preview` 读取当前预览图片，检查换行、对齐和跨页一致性；结构校验通过不代表已经目检。参考 PPT 中的复杂 Office 对象和缺失字体可能影响转换效果，应在保存前检查预览。

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
