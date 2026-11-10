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

独立插件继续使用本项目的接入方式。主插件保存共享 PPT 服务、HTTP 路由与状态，`ppt` 原生 Agent Preset 子插件注册工具、提示词和 Skill。客户端通过官方 `conversation.input.dock` 插槽显示模板面板，只在 PPT 预设的空白会话中显示，首次发送后固定预设和模板。模板选择器恢复上游 `46666d34277571a7b48ada6060a9428ee503b136` 的基础 CSS、零高度定位容器、透明面板和宽高测量；兼容插槽只补充排序和间距抵消，避免模板区域把输入框向上推。保留本项目的 AI 模板制作入口与预览保存流程，没有引入 Desktop 的工具栏按钮、Office 模式管理或桌面启动器。

公共包名、浏览器模块标识与 invariant 归属继续使用 `@cola1900/dsh-ppt`。数据目录保留 `kimi-ppt`，延续已有会话、修订与输出。内置 16 套模板及双语示例保留在包内，构建只使用本地预览生成目录与白名单，不依赖上游工程或网络。

上游的 Office 预览依赖其对 `@deepseek-ai/libreoffice-kit@0.1.5` 添加的逐图回调补丁。本项目使用公开 SDK 的 `renderImages` 返回结果按页发布预览，保留批次、页数、取消及资源预算检查，不修改宿主依赖。

Harness 接口以 `0.2.1-rc.5` 为兼容和开发基线。PPTD 工具通过宿主的 `workingDirectory` 与 `sandboxPolicy` 服务取得实际目录和权限。CLI 与插件共用 `pptd.js` 的解析、校验和导出实现，嵌套字段规则集中在 `pptd-schema.js`。CLI 的快速 SVG/PNG 预览仍是独立的草图渲染，正式模板预览使用 Office 渲染结果。

`scripts/upstream-manifest.json` 固定上游提交，并逐项记录本地补丁、生成文件和仅属于本项目的模块。`scripts/check-upstream.mjs <upstream-checkout>` 检查完整模块清单、未改动模块、三份 Skill 文档，以及每个修改模块的上游哈希、本地哈希和可重放补丁。`index.js` 不再整文件豁免；存储与审计移至 `ppt-store.js`，交付恢复记录移至 `ppt-exports.js`。客户端由 `src` 构建并通过交互测试验证。

修改已列明的上游模块后，先审查差异，再运行 `node scripts/update-upstream-patches.mjs <upstream-checkout>` 更新补丁和哈希，并检查生成的 `scripts/upstream-patches/`。该命令拒绝自动放行清单外的新差异。CI 检出固定上游提交并重放这些补丁。

本地可靠性补丁包括每个模型步骤更新模板上下文、隔离损坏模板记录、制作预览的失败与取消状态，以及导出的恢复记录。交付目录发布后、会话写入失败时，重试根据工程内容及调用标识恢复原交付，并保留较新的会话状态。会话状态原子写入是业务提交点；审计通过待投递记录恢复，失败不撤销已提交业务。审计在崩溃恢复后可能重复投递，活动 ID 用于识别同一事件。

个人模板结构索引按实际文本解析中英文字体配置，保持索引中的字体名称为字符串，原生工程中的字体配置不变。新索引在生成和保存时校验。旧版已保存或待保存的原生模板，如果派生索引无效，会从保留的受控工程重新提取，并核对页面索引；读取只返回恢复后的结果，显式保存或修改时才写入元数据。无法恢复的单条记录不阻断内置模板及其他个人模板加载。

Harness rc.5 的 profile 包解析器无法解析 JSdom 间接依赖中的 `punycode/`。发布构建将 JSdom 及纯 JavaScript 依赖打包到 `lib/vendor`，同时保留默认样式表、同步 XHR worker 和依赖许可证。核心模块只将 JSdom 导入改为本地适配器，核对上游时仅归一化这一导入路径。DOM 实现仍为上游固定的 JSdom 29.1.1，无需修改宿主或执行消费者安装脚本。

上游自身说明其运行时代码来自已分发 JavaScript，完整原始 TypeScript 实现不可得。保留实现中的来源声明，并同步对应声明文件。PPT 核心原始 MIT 声明位于根 LICENSE，桌面项目 MIT 声明位于 `licenses/dsh-desktop-LICENSE`，其他归属见 THIRD_PARTY_NOTICES.md。
