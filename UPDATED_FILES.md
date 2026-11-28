# 2.2.2 更新文件清单

基于用户提供的 2.2.1 完整项目，调整序列手动坐标轴分配的配置归属。

| 文件 | 变更 |
| --- | --- |
| src/core.js | seriesAxes 归入序列配置；分类校验、默认值及旧配置迁移 |
| src/app.js | 序列配置保存和恢复全部手动分配；坐标轴配置应用不干扰序列归属；已修改状态和数量提示 |
| src/config-store.js | IndexedDB 一次性迁移；旧云端与导入文件兼容；分类导出使用新归属 |
| src/template.html | 分类改名为“序列”，更新手动分配提示和使用说明 |
| configs/axes.json | 移除 defaults.seriesAxes |
| configs/selection.json | 增加 defaults.seriesAxes，交付默认为空数组 |
| configs/README.md | 更新字段归属与旧配置迁移说明 |
| tests/config-types.test.cjs | 分类字段归属、旧配置拆分和冲突保留测试 |
| tests/config-store.test.cjs | IndexedDB 持久迁移、旧云端默认值和分类导出测试 |
| tests/paste-ui.test.cjs | 四种序列状态、未勾选分配、刷新恢复和轴配置独立性测试 |
| package.json、package-lock.json | 版本更新为 2.2.2 |
| README.md、RELEASE_NOTES.md、UPDATED_FILES.md | 使用方法、JSON 示例、迁移行为及验证结果 |
| index.html | 从源码重新构建 |
| curve-viewer.zip、curve-viewer-static.zip | 完整项目包与静态部署包重新生成并逐文件校验 |
| 最终结果/ | 同步本版交付文件及 ZIP，保留无关历史内容 |

146 项自动化测试通过。vendor、configs/regex.json、数据解析及绘图模块未修改。原始输入压缩包未改写。

部署时更新页面及相应配置结构；已有云端配置请保留并迁移，不要用空配置列表覆盖。每条序列的手动归属只写入 selection.json，axes.json 不再包含 seriesAxes。
