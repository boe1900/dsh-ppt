# 提取来源

- 仓库：https://github.com/dataelement/dsh-desktop
- 固定提交：`97e02f8bf773a3b28bd4c6c05f874156be8ab5ab`
- 提取日期：2026-09-17
- 核心：`packages/ppt-runtime/core/`
- 输入框界面：`packages/ppt-runtime/adapter/lib/client.js`
- 模板：`packages/ppt-runtime/templates/`
- 模板许可证与来源资料：`packages/ppt-runtime/upstream/`

独立化改动：合并核心和 composer 为单个 `dsh-ppt` 包；客户端模块标识统一为 `dsh-ppt`；将 Desktop 专有输入框插槽替换为官方 `conversation.input.dock`，按钮、选中缩略图和模板列表在输入框上方显示；移除指向缺失文件的 `client-standard` 导出，并使用匹配实际客户端的类型声明；将构建改为读取包内模板、生成目录和预览白名单；移除未使用的 TypeScript、Zod 运行依赖；增加独立安装、打包说明及运行验证。

公共 npm 包名为 `@cola1900/dsh-ppt`，插件加载入口、浏览器模块标识和 invariant 归属同步使用该包名。

核心工具、PPTD 编译器、文件访问和授权逻辑继续使用上游维护的 JavaScript。保留 `kimi-ppt` 历史存储目录、旧 RPC 别名和迁移逻辑。没有复制 Electron、桌面启动器、插件市场或原工程的补丁/安装脚本。

上游自身说明其运行时代码来自已分发 JavaScript，完整原始 TypeScript 实现不可得。声明文件和历史来源注释原样保留。PPT 核心原始 MIT 声明位于根 LICENSE，桌面项目 MIT 声明位于 `licenses/dsh-desktop-LICENSE`，其他归属见 THIRD_PARTY_NOTICES.md。
