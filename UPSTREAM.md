# 提取来源

- 仓库：https://github.com/dataelement/dsh-desktop
- 固定提交：`46666d34277571a7b48ada6060a9428ee503b136`
- 同步日期：2026-11-09
- 上次基线：`97e02f8bf773a3b28bd4c6c05f874156be8ab5ab`
- 核心：`packages/ppt-runtime/core/lib/`
- Skill：`packages/ppt-runtime/core/skills/dsh-ppt/`
- 自定义模板交互：`packages/ppt-runtime/client/personal-template-manager.js`
- 模板许可证与来源资料：`packages/ppt-runtime/upstream/`

本次同步核心编译器、CLI、PPTX 转换、原生对象导出、字体与文本布局、素材处理、校验、模板参考工具和个人模板库。个人模板包括分片上传、隔离导入、逐页预览、保存与复用、编辑、删除、取消、重复文件识别及独立工程复制。

独立插件继续使用本项目的接入方式。主插件保存共享 PPT 服务、HTTP 路由与状态，`ppt` 原生 Agent Preset 子插件注册工具、提示词和 Skill。客户端通过官方 `conversation.input.dock` 插槽显示现有模板面板，只在 PPT 预设的空白会话中显示，首次发送后固定预设和模板。没有引入 Desktop 的工具栏按钮、Office 模式管理或桌面启动器。

公共包名、浏览器模块标识与 invariant 归属继续使用 `@cola1900/dsh-ppt`。数据目录保留 `kimi-ppt`，延续已有会话、修订与输出。内置 16 套模板及双语示例保留在包内，构建只使用本地预览生成目录与白名单，不依赖上游工程或网络。

上游的 Office 预览依赖其对 `@deepseek-ai/libreoffice-kit@0.1.5` 添加的逐图回调补丁。本项目使用公开 SDK 的 `renderImages` 返回结果按页发布预览，保留批次、页数、取消及资源预算检查，不修改宿主依赖。

Harness 接口以 `0.2.1-rc.5` 为兼容和开发基线。PPTD 工具通过宿主的 `workingDirectory` 与 `sandboxPolicy` 服务取得实际目录和权限。`scripts/check-upstream.mjs` 可核对未改动的上游核心模块与三份 Skill 文档；本地接入层和 SDK 适配单独测试。

上游自身说明其运行时代码来自已分发 JavaScript，完整原始 TypeScript 实现不可得。保留实现中的来源声明，并同步对应声明文件。PPT 核心原始 MIT 声明位于根 LICENSE，桌面项目 MIT 声明位于 `licenses/dsh-desktop-LICENSE`，其他归属见 THIRD_PARTY_NOTICES.md。
