# openpyxl完全学习笔记

## 得到一个工作簿

```python
from openpyxl import Workbook
wb = Workbook()
```
### 从文件读入

```python
wb = load_workbook('test.xlsx')
```

## 选择一个活动的sheet

```python
ws = wb.active

# 通过名字取得一个sheet

ws = wb["sheet_name"]
```

# 遍历所有的sheet的名字

(list[str])wb.sheetnames

# 取得副本

ws_copy = wb.copy_worksheet(sheet_name)

注意，属性之类都会拷贝上，但是图表不会。不能在wb之间拷贝sheet

### 创建

```python
Workbook.create_sheet("Sheet Name", [SheetPos])

# 缺省-append，0-开始位置，-1-penultimate 倒数第二
```

### attribute

× title - 标题

× sheet_properities.tabcolor - 标签下面的颜色

## 访问表格

所有的操作都是在sheet中的！

```python
c = ws['A4']

ws['A4'] = 4

# 提供函数方法

d = ws.cell(row = 4, column = 2, walue = 10)

# 新创建的sheet中并不存在cell，在第一次访问时创建，所以在遍历时要小心。

```

### 访问大量表格

```python
cell_range = ws['A1':'C2']

# 选中方式取得区间的内容

row_range = ws[5:10]
col_range = ws['C:D']
row10 = ws[10]
row_range = ws[5:10]

```

### 迭代器遍历

```python
# 遍历每一行中的内容

for row in ws.iter_rows(min_row = 1, max_col = 3, max_row = 2):
	for cell in row:
		print(cell)

'''
输出结果如下：
A1
B1
C1
A2
B2
C2
'''

# 遍历列内容
for col in ws.iter_cols(min_row=1, max_row=3, max_col=3):
	for cell in col:
		print(cell)

'''
输出结果如下：
A1
A2
B1
B2
C1
C2
'''
# 目前题目中的例子是左边顶到头的
```

注意只读模式下列遍历是不可用的

#### 遍历整个sheet

```python
tuple(ws.rows)
tuple(ws.columns)
```
同理，在只读模式下列遍历不可用

#### 只访问值

不访问cell对象

```python
for row in ws.rows:
	for value in row:
		print(value)

# 在更高级的循环中使用是只需要打上标记

for row in ws.iter_rows(min_row=1, max_row=2, max_col=3, values_only=True):
	print(row)

```
## 存信息

主要操作Cell

访问Cell.value

## 存盘

```python
wb.save('filename.xlsx')
```
### 以文件流存储信息

待补
