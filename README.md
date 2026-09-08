# 万相台周报助手

一个可分享的 Codex 插件：首次为每个品牌创建统一飞书周报并确认计划映射；以后只需说日期范围，就能抓取阿里妈妈万相台“人群推广”的计划汇总与全部素材，并安全追加到历史周报。

## 能力

- 不同品牌、不同计划数量分别初始化，不携带作者的品牌数据。
- 抓取展现、CR、CTR、直接成交金额、花费和总 ROI。
- 记录全部图片和视频素材，素材超过模板容量时自动扩列。
- 填写素材日期、图片、视频封面、视频地址、类型、方向和直接 ROI 公式。
- 图片和视频分别选择最佳素材并标黄。
- 写入前检查空白块，写后回读；已存在的周期拒绝重复写入。
- 每个计划独立保存原始 JSON 和进度，可断点继续。

## 安装

私有仓库要求安装者的 Git 已有该仓库访问权限。

```powershell
codex plugin marketplace add puyuchen86-cpu/wanxiangtai-weekly-reporter --ref main
codex plugin add wanxiangtai-weekly-reporter@puyuchen86-wanxiangtai
```

安装或升级后，请新建一个 Codex 任务，使插件说明被重新加载。

## 首次使用

先在 Chrome 登录自己的阿里妈妈万相台账号，并完成本机 `lark-cli` 的飞书用户授权。然后对 Codex 说：

> 初始化我的万相台周报，创建飞书表格并匹配当前账号中的所有人群推广计划。

Codex 会展示一次计划与工作表映射，确认后创建周报并把配置保存在使用者本机。仓库不保存 Cookie、密码、飞书访问令牌、工作簿 token 或计划 ID。

## 每周使用

初始化完成后，只需说：

> 抓取 9.7-9.13。

默认行为是追加到下一空白周期块、不覆盖历史、忽略隐藏工作表和月度数据表，并记录页面中实际出现的全部素材。

## 环境要求

- Windows 10/11
- Codex 桌面版及可用的浏览器控制能力
- Chrome 中已登录阿里妈妈万相台
- Python 3.10+
- PowerShell 7+
- 已配置并以用户身份授权的 `lark-cli`

## 隐私与安全

- 登录凭据只保留在用户自己的浏览器或 `lark-cli` 凭据存储中。
- 本地品牌配置位于用户目录，不应提交到 Git。
- 所有飞书写入均先检查目标块；异常时停止，不自动清空或覆盖。
- 本仓库为私有分发，未经仓库所有者许可不得再分发。

## 仓库结构

```text
.agents/plugins/marketplace.json
plugins/wanxiangtai-weekly-reporter/
  .codex-plugin/plugin.json
  skills/wanxiangtai-weekly-reporter/
    SKILL.md
    agents/openai.yaml
    scripts/
    references/
```
