# excel
A simple PHP utility for working with Excel files using PhpSpreadsheet.

```php
use Natilosir\Excel\Excel;
```


### Example 1: Create Excel file from array
```php
$data = [
    ["Name", "Age", "Salary"],
    ["John Doe", 28, 5000],
    ["Jane Smith", 34, 7000],
    ["Mark Taylor", 45, 9000],
];
Excel::ToExcel($data, "example.xlsx");
```
---
### Example 2: Convert Excel file to array
```php
    $arrayFromExcel = Excel::ToArray("example.xlsx");
    print_r($arrayFromExcel);
```
