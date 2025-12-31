# Excel表格合并工具

一个简单易用的Python工具，用于合并多个具有相同表头的Excel文件，并自动重新生成第一列的递增序号。

## 功能特点

✨ **自动合并** - 合并多个Excel文件到一个文件中
✨ **表头验证** - 自动检查所有文件的表头是否一致
✨ **序号重建** - 自动重新生成第一列的递增序号（从1开始）
✨ **友好提示** - 清晰的进度显示和错误提示

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 基本用法

```bash
python excel_merger.py <输出文件名> <输入文件1> <输入文件2> [输入文件3] ...
```

### 示例

```bash
# 合并两个文件
python excel_merger.py merged.xlsx file1.xlsx file2.xlsx

# 合并多个文件
python excel_merger.py output.xlsx data1.xlsx data2.xlsx data3.xlsx data4.xlsx
```

### 使用示例文件测试

项目中包含了三个示例Excel文件用于测试：

```bash
python excel_merger.py result.xlsx example1.xlsx example2.xlsx example3.xlsx
```

## 工作原理

1. **加载文件** - 读取所有指定的Excel文件
2. **验证表头** - 确保所有文件的表头完全一致
3. **合并数据** - 将所有文件的数据行合并到一起
4. **重建序号** - 自动为第一列生成从1开始的递增序号
5. **保存结果** - 将合并后的数据保存到新的Excel文件

## 要求

- 所有输入的Excel文件必须具有**完全相同的表头**
- 第一列通常是序号列，合并后会自动重新编号
- 支持 `.xlsx` 和 `.xls` 格式的Excel文件

## 示例场景

假设你有以下三个Excel文件，都包含学生信息：

**file1.xlsx:**
| 序号 | 姓名 | 年龄 | 成绩 |
|------|------|------|------|
| 1    | 张三 | 20   | 85   |
| 2    | 李四 | 21   | 90   |

**file2.xlsx:**
| 序号 | 姓名 | 年龄 | 成绩 |
|------|------|------|------|
| 1    | 王五 | 19   | 88   |
| 2    | 赵六 | 22   | 92   |

**file3.xlsx:**
| 序号 | 姓名 | 年龄 | 成绩 |
|------|------|------|------|
| 1    | 孙七 | 20   | 87   |

运行命令：
```bash
python excel_merger.py merged.xlsx file1.xlsx file2.xlsx file3.xlsx
```

**结果 merged.xlsx:**
| 序号 | 姓名 | 年龄 | 成绩 |
|------|------|------|------|
| 1    | 张三 | 20   | 85   |
| 2    | 李四 | 21   | 90   |
| 3    | 王五 | 19   | 88   |
| 4    | 赵六 | 22   | 92   |
| 5    | 孙七 | 20   | 87   |

## 错误处理

工具会自动检测并提示以下错误：

- ❌ 文件不存在
- ❌ 文件为空
- ❌ 表头不一致
- ❌ 文件读取错误
- ❌ 文件保存错误

## 技术栈

- Python 3.6+
- pandas - 数据处理
- openpyxl - Excel文件读写

## 许可证

MIT License
