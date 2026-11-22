# Zotero–Word

联动 Word 与 Zotero，按目标期刊样式整理动态引文和参考文献。提供 macOS 与 Windows 两个版本。

## 功能

- 按 DOI、论文链接或作者与年份匹配文献。
- 整理稿件对应的 Zotero 文献集合。
- 在 Word 工作副本中处理动态引文和 References，并由 Zotero 刷新。
- 保留原稿；遇到匹配歧义或权限问题，保留现场并支持继续处理。

## 选择版本

下载下方适合系统的 ZIP，解压后阅读包内说明。

| 系统 | 下载文件 | 安装与使用 |
| --- | --- | --- |
| macOS | [zotero-word-macos-v2.4.4.zip](downloads/zotero-word-macos-v2.4.4.zip?raw=1) | [Mac 使用说明](macos/zotero-word/使用说明.md) |
| Windows | [zotero-word-windows-v2.4.4.zip](downloads/zotero-word-windows-v2.4.4.zip?raw=1) | [Windows 安装说明](windows/zotero-word/INSTALL-WINDOWS.md) |

两个包的技能名均为 `zotero-word`，按运行系统选择其中一个安装。

## 运行条件

本机安装 Word、Zotero 及对应的 Word 插件；使用能调用本机应用的 Codex 或 DeepSeek Harness。具体依赖和首次连接步骤见各平台说明。此处的 DeepSeek Harness 不等于 DeepSeek 网页聊天。

## 使用示例

> 使用 zotero-word skill，按目标期刊要求整理这份 Word 稿件的动态引文和参考文献。

请提供稿件和目标期刊；匹配有歧义时确认实际文献条目。

## 验证范围

当前版本为 2.4.4。本次发行未新增两平台原生回归测试。Windows 适配尚未完成实机验收，建议先用测试副本验证；各平台的支持范围以包内说明为准。

## 源码目录

- `macos/zotero-word/`：macOS 版本。
- `windows/zotero-word/`：Windows 版本。

提交问题时，请注明系统、Word/Zotero 版本、执行宿主及错误信息；分享日志前移除稿件正文和个人信息。
