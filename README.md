# ArrayToExcel [![NuGet version](https://badge.fury.io/nu/ArrayToExcel.svg)](http://badge.fury.io/nu/ArrayToExcel)
Create Excel from Array

### Example 1: Create with default settings
```C#
using ArrayToExcel;

var items = Enumerable.Range(1, 10).Select(x => new
{
    Prop1 = $"Text #{x}",
    Prop2 = x * 1000,
    Prop3 = DateTime.Now.AddDays(-x),
});

var excel = items.ToExcel();
```
Result:
[example1.xlsx](https://github.com/mustaddon/ArrayToExcel/raw/master/Examples/example1.xlsx)

![](https://raw.githubusercontent.com/mustaddon/ArrayToExcel/master/Examples/example1.png)


### Example 2: Rename sheet and columns
```C#
var excel = items.ToExcel(schema => schema
    .SheetName("Example name")
    .ColumnName(m => m.Name.Replace("Prop", "Column #")));
```
Result:
[example2.xlsx](https://github.com/mustaddon/ArrayToExcel/raw/master/Examples/example2.xlsx)

![](https://raw.githubusercontent.com/mustaddon/ArrayToExcel/master/Examples/example2.png)


### Example 3: Sort columns
```C#
var excel = items.ToExcel(schema => schema
    .ColumnSort(m => m.Name, desc: true));
```
Result:
[example3.xlsx](https://github.com/mustaddon/ArrayToExcel/raw/master/Examples/example3.xlsx)

![](https://raw.githubusercontent.com/mustaddon/ArrayToExcel/master/Examples/example3.png)


### Example 4: Custom column's mapping
```C#
var excel = items.ToExcel(schema => schema
    .AddColumn("MyColumnName#1", x => x.Prop1)
    .AddColumn("MyColumnName#2", x => $"test:{x.Prop2}")
    .AddColumn("MyColumnName#3", x => x.Prop3));
```
Result:
[example4.xlsx](https://github.com/mustaddon/ArrayToExcel/raw/master/Examples/example4.xlsx)

![](https://raw.githubusercontent.com/mustaddon/ArrayToExcel/master/Examples/example4.png)


### Example 5: Additional sheets
```C#
var excel = items.ToExcel(schema => schema
    .SheetName("Main")
    .AddSheet(extraItems));
```
Result:
[example5.xlsx](https://github.com/mustaddon/ArrayToExcel/raw/master/Examples/example5.xlsx)

![](https://raw.githubusercontent.com/mustaddon/ArrayToExcel/master/Examples/example5.png)


### Example 6: Create from DataSet
```C#
var dataSet = new DataSet();

foreach (var i in Enumerable.Range(1, 3))
{
    var table = new DataTable($"Table{i}");
    dataSet.Tables.Add(table);

    table.Columns.Add($"Column {i}-1", typeof(string));
    table.Columns.Add($"Column {i}-2", typeof(int));
    table.Columns.Add($"Column {i}-3", typeof(DateTime));

    foreach (var x in Enumerable.Range(1, 10 * i))
        table.Rows.Add($"Text #{x}", x * 1000, DateTime.Now.AddDays(-x));
}

var excel = dataSet.ToExcel();
```
Result:
[example6.xlsx](https://github.com/mustaddon/ArrayToExcel/raw/master/Examples/example6.xlsx)

![](https://raw.githubusercontent.com/mustaddon/ArrayToExcel/master/Examples/example6.png)


[Example.ConsoleApp](https://github.com/mustaddon/ArrayToExcel/tree/master/Examples/Example.ConsoleApp/)
