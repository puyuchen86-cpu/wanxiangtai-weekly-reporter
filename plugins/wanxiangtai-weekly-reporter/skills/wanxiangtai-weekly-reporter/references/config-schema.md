# 品牌配置结构

配置文件存放在使用者本机，不进入 Git。

```json
{
  "schema_version": 1,
  "profile_id": "brand-a",
  "brand_name": "品牌A",
  "spreadsheet": {
    "token": "spreadsheet token",
    "url": "https://.../sheets/..."
  },
  "excluded_sheets": ["计划月数据"],
  "template": {
    "first_block_start": 3,
    "block_height": 27,
    "block_stride": 28,
    "material_start_column": "G",
    "default_material_capacity": 24
  },
  "plans": [
    {
      "campaign_id": "123456789",
      "plan_name": "人群推广_示例计划_内容智能出价",
      "sheet_id": "abc123",
      "sheet_name": "示例计划_内容智能出价"
    }
  ],
  "data_root": "用户本机的绝对目录"
}
```

约束：

- `profile_id` 只允许小写字母、数字和连字符。
- campaign ID、sheet ID 和 sheet name 在一个 profile 内必须唯一。
- 不允许 Cookie、密码、access token、refresh token、app secret 等字段。
- 工作簿 token 是资源标识，不是登录凭据，但仍只保存在本机配置。
