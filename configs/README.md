# 默认值与云端配置目录

本目录与网站的 index.html 同级。打开网页、刷新网页或点击云端“刷新”，会读取 ./configs/ 下的三个 JSON，无需运行命令或选择文件夹。

每个文件采用 curve-viewer-bundle v1 文档结构：

- schema：curve-viewer-bundle。
- version：1。
- type：与文件对应的 axes、regex 或 selection。
- defaults：该类型的默认值，是程序默认配置的唯一来源。
- configs：同类型的已保存配置数组；当前交付为空数组，不包含示例曲线数据。

## 默认字段

| 文件 | 字段 | 含义 |
| --- | --- | --- |
| axes.json | defaults.timeMode | 默认时间格式 |
| axes.json | defaults.axisGroups | 坐标组的范围、名称、颜色及顺序；固定标识为 small、thermal、medium、large、auto |
| axes.json | defaults.axisRegions | 各组起始、终止比例，从顶部按 0～100 计算；必须包含五个组 |
| axes.json | defaults.autoBounds | paddingPercent 为向两端拓宽百分比，minimumConstantPadding 为常量最小拓宽量，roundOutward 控制向外取整 |
| axes.json | defaults.axisLimits | 默认手动上下限；空对象表示按数据自动计算 |
| axes.json | defaults.seriesAxes | 默认的单条曲线坐标分配 |
| regex.json | defaults.headerRule | enabled、pattern、flags、replacement |
| selection.json | defaults.selectedSeries | 默认勾选序列；当前为空，首次加载不全选 |
| selection.json | defaults.palette | 曲线默认配色列表 |
| selection.json | defaults.fallbackColor | 配色列表用完后的 HSL 色彩规则：hueStep、saturation、lightness |

直接修改上述 JSON 并更新网站，刷新网页后生效，无需修改业务 JS 或重建 HTML。最后选择的个人配置优先；“恢复默认规则”“恢复默认区域”从当前已读取的 defaults 取值。

页面的“保存”仍只更新本地 IndexedDB。“导出”会将当前已读取的 defaults 和当前类型的全部本地配置 configs 一起写入同一个 JSON，可手动替换网站文件。没有本地配置时 configs 为 []，仍保留 defaults。导出不自动合并云端旧列表，需保留时先导入本地或自行合并 configs。

导入到本地时只导入 configs，不会改变网站默认值。旧版单项配置和配置数组继续支持本地导入；升级旧云端文件时，可将旧数组放入新文档的 configs，保留完整 defaults。

三个文件都必须存在且默认值有效。首次载入失败时提示检查文件，不使用 JS 中的备用默认值；刷新失败保留之前成功载入的配置。成功刷新不会自动覆盖当前曲线，新的默认值用于后续加载、恢复默认与导出。
