# pexcel
example

```python
from pexcel import Workbook
from pexcel.font import Font


wb = Workbook()
data = [['hello', 'world', 'excel'], ['hello', 5.33, 6.78]]
ws = wb.new_sheet('H5', data)
ws['A1'].font = Font(bold=True)
ws['A2'].font = Font(italic=True)
ws['B1'].font = Font(italic=True)

wb.save('hello.xlsx')
```
