# 2.2.1 更新文件清单

本次针对直接打开本地 HTML 时出现“Failed to fetch”补充明确诊断，静态网站的相对路径读取方式保持不变。

| 文件 | 变更 |
| --- | --- |
| src/app.js | 检测 file:// 打开方式，提示静态网站部署；传入网页基地址供错误诊断，区分首次失败与保留已有配置 |
| src/config-store.js | 网络失败显示文件名、请求地址、原始原因，区分超时或取消 |
| tests/config-store.test.cjs | 新增请求错误地址及原因验证 |
| tests/paste-ui.test.cjs | 新增 file:// 下不请求 JSON、明确提示且不引入默认值副本的测试 |
| package.json、package-lock.json | 版本更新为 2.2.1 |
| README.md、RELEASE_NOTES.md、UPDATED_FILES.md | 补充打开方式说明、验证与更新清单 |
| index.html | 重新构建页面 |
| curve-viewer.zip、curve-viewer-static.zip | 重打包完整项目和静态网站 |
| 最终结果/ | 同步以上交付文件和 ZIP |

三个 configs JSON 及 vendor 本次未修改。部署时可使用新的静态包整体更新；已有云端配置需保留自己的 JSON。
