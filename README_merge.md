# Excel 表格合并工具

这个工具用于合并两个 Excel 表格，自动跳过表头和表尾各 4 行非列表内容。

## 功能特点

- 自动跳过表头 4 行和表尾 4 行非列表内容
- 合并两个 Excel 文件的数据部分
- 保持原始表头和表尾格式
- 支持自定义输出路径
- 自动创建输出目录

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 命令行使用

```bash
# 基本用法（自动生成输出文件名）
python src/merge.py resource/file1.xlsx resource/file2.xlsx

# 指定输出文件路径
python src/merge.py resource/file1.xlsx resource/file2.xlsx output/merged_result.xlsx
```

### 参数说明

- `文件1路径`: 第一个要合并的 Excel 文件路径
- `文件2路径`: 第二个要合并的 Excel 文件路径
- `输出文件路径`: (可选) 合并后的输出文件路径，如果不指定会自动生成

## 输出文件

合并后的文件将保存在 `output/` 目录下，文件名格式为：
`merged_文件1名_文件2名.xlsx`

## 注意事项

1. 确保两个 Excel 文件都有相同的列结构
2. 表头和表尾各 4 行会被保留，但不会参与数据合并
3. 如果文件行数少于 8 行，会跳过表头 4 行，不处理表尾
4. 合并后的文件会使用第一个文件的表头和表尾格式

## 示例

假设您有两个文件：

- `resource/销售数据1.xlsx`
- `resource/销售数据2.xlsx`

执行合并：

```bash
python src/merge.py resource/销售数据1.xlsx resource/销售数据2.xlsx
```

输出文件将是：
`output/merged_销售数据1_销售数据2.xlsx`
