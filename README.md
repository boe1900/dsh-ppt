# dsh-ppt

从 [dataelement/dsh-desktop](https://github.com/dataelement/dsh-desktop) 提取的独立 **DSH / DeepSeek Harness 插件**。一个包包含 PPT 按钮、模板选择、PPTD 编写与校验、预览和可编辑 PPTX 导出，无需 Electron 或原 Desktop 仓库。

- 16 套模板、192 个版式，保留英文/中文 PPTD 示例和本地预览图。
- 提供原生 DSH **PPT 预设**；只有选择该预设时才注入 PPT 创作指引。
- 保留原有文件授权、路径限制、写入哈希检查、版本记录和导出校验。
- 提供 `dsh-pptd` 命令行和 `@cola1900/dsh-ppt/pptd` JavaScript API。

## 安装

需要 Node.js `^22.19.0 || >=24.0.0`，以及提供标准会话输入框插槽的 DSH Web / DSH Desktop。提取基线使用 DSH `0.1.5-rc.2`。

从 [公共 npm 包](https://www.npmjs.com/package/@cola1900/dsh-ppt) 直接安装到目标 DSH profile：

```sh
dsh plugin --profile web add @cola1900/dsh-ppt
```

把 `web` 替换为实际 profile，然后重启该 profile。新会话顶部的原生预设选择器会出现 **PPT**；选择后，输入框下方显示模板面板，在首次发送前选择模板，再描述需求。首次发送后预设和模板固定，新会话可以重新选择。插件首次启动会把内置 preset 写入 `$DSH_HOME/.agent-presets/ppt`，无需手动复制或配置。

界面只使用官方 `conversation.input.dock` 插槽，只设置插件自身样式，不修改原 DSH。模板名称继续在卡片内显示省略号，输入框底部不再显示 `PPT · 模板名` 控件。

如果目标 Desktop 已内置 `dsh-ppt-composer`，先在目标 profile 的插件配置中停用原 PPT bundle，再加载本包。不要同时启用两套 PPT bundle，否则会重复注册工具、路由和输入框按钮。本包已经合并 composer，无需再安装 `dsh-ppt-composer`。

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

`lib/` 是上游维护的 JavaScript 运行时代码；`src/client.js` 是输入框客户端的构建输入。`presets/ppt/` 是随插件安装的原生 Agent Preset 定义，首次启动会同步到 DSH 的用户 preset 根目录。完整原始 TypeScript 源码不在上游提取基线中，本项目没有把声明文件冒充实现源码。

`skills/dsh-ppt/references/` 保存模板元数据、设计说明、双语源文件和预览。修改元数据或预览后执行 `npm run build`，生成模板目录、预览白名单和客户端预览映射。构建使用已随包保存的预览，不依赖网络或原桌面工程。修改版式后可用 `dsh-pptd screenshot` 重新生成图片，再更新对应的 JPG 预览。

默认数据目录沿用 `dshHomePath('kimi-ppt')`，以便保留旧会话、版本和输出。可以在 Cordis 插件配置中设置绝对路径 `root`；`maxSlides`、`maxDecksPerSession`、`maxActivities` 保留原有上限配置。`pptSkillRoot` 或 `DSH_PPT_SKILL_ROOT` 可覆盖技能目录。

提取版本和具体改动见 [UPSTREAM.md](UPSTREAM.md)。许可证与模板来源见 [LICENSE](LICENSE)、[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) 及 `licenses/`。
