using ArrayToExcel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;

namespace ConsoleApp;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
        Example3();
        Example4();
        Example5();
        Example6();

        TestTypes();
        TestDictionary();
        TestExpandoObject();
        TestHashtable();
        TestDataTable();
        TestObject();
    }

    static readonly IEnumerable<SomeItem> SomeItems = Enumerable.Range(1, 10).Select(x => new SomeItem
    {
        Prop1 = $"Text #{x}",
        Prop2 = x * 1000,
        Prop3 = DateTime.Now.AddDays(-x),
    });

    // default settings
    static void Example1()
    {
        var excel = SomeItems.ToExcel();

        File.WriteAllBytes($@"..\..\..\..\{nameof(Example1)}.xlsx".ToLower(), excel);
    }

    // rename sheet and columns
    static void Example2()
    {
        var excel = SomeItems.ToExcel(schema => schema
            .SheetName("Example name")
            .ColumnName(m => m.Name.Replace("Prop", "Column #")));

        File.WriteAllBytes($@"..\..\..\..\{nameof(Example2)}.xlsx".ToLower(), excel);
    }

    // sort columns
    static void Example3()
    {
        var excel = SomeItems.ToExcel(schema => schema
            .ColumnSort(m => m.Name, desc: true));

        File.WriteAllBytes($@"..\..\..\..\{nameof(Example3)}.xlsx".ToLower(), excel);
    }

    // custom column's mapping
    static void Example4()
    {
        var excel = SomeItems.ToExcel(schema => schema
            .AddColumn("MyColumnName#1", x => new CellHyperlink($"https://www.google.com/search?q={x.Prop1}", x.Prop1))
            .AddColumn("MyColumnName#2", x => $"test:{x.Prop2}")
            .AddColumn("MyColumnName#3", x => x.Prop3));

        File.WriteAllBytes($@"..\..\..\..\{nameof(Example4)}.xlsx".ToLower(), excel);
    }

    // additional sheets
    static void Example5()
    {
        var extraItems = Enumerable.Range(1, 10).Select(x => new
        {
            ExtraListProp1 = x,
            ExtraListProp2 = x * 10,
            ExtraListProp3 = x * 100,
            ExtraListProp4 = x * 1000,
        });

        var excel = SomeItems.ToExcel(schema => schema
            .SheetName("Main")
            .AddSheet(extraItems));

        File.WriteAllBytes($@"..\..\..\..\{nameof(Example5)}.xlsx".ToLower(), excel);
    }

    // DataSet
    static void Example6()
    {
        var dataSet = new DataSet();

        for (var i = 1; i <= 3; i++)
        {
            var table = new DataTable($"Table{i}");
            dataSet.Tables.Add(table);

            table.Columns.Add($"Column {i}-1", typeof(string));
            table.Columns.Add($"Column {i}-2", typeof(int));
            table.Columns.Add($"Column {i}-3", typeof(DateTime));

            for (var x = 1; x <= 10 * i; x++)
                table.Rows.Add($"Text #{x}", x * 1000, DateTime.Now.AddDays(-x));
        }

        var excel = dataSet.ToExcel();

        File.WriteAllBytes($@"..\..\..\..\{nameof(Example6)}.xlsx".ToLower(), excel);
    }


    // different types + stream
    static void TestTypes()
    {
        var items = Enumerable.Range(1, 100).Select(x => new
        {
            String = "  1) text text text; \n2) text text text !!!",
            WrapText = new CellText("1) text text text; \n2) text text text", true),
            Bool = x % 2 == 0,
            NullableBool = x % 2 == 0 ? (bool?)true : null,
            Int = -x * 100,
            Uint = (uint)x * 100,
            Long = (long)x * 100,
            Double = 1.1d + x,
            Float = 1.1f + x,
            Decimal = 1.1m + x,
            DateOnly = DateOnly.FromDateTime(DateTime.Now.AddDays(-x)),
            DateTime = DateTime.Now.AddDays(-x),
            DateTimeOffset = DateTimeOffset.Now.AddDays(-x),
            Uri = new Uri($"https://www.google.com/search?q={x}"),
            Hyperlink = new CellHyperlink($"https://www.google.com/search?q={x}", $"link_{x}"),
            Formula = new CellFormula(row => $"G{row}+H{row}"),
            Percent = new CellPercent(1d / x),
        });


        using var file = File.Create($@"..\{nameof(TestTypes)}.xlsx");
        items.ToExcel(file);
    }

    // list of dictionaries 
    static void TestDictionary()
    {
        var items = Enumerable.Range(1, 100).Select(x => new Dictionary<object, object>
        {
            { "Column #1", $"Text #{x}" },
            { "Column #2", x * 1000 },
            { "Column #3", DateTime.Now.AddDays(-x) },
        });

        var excel = items.ToExcel(s => s
            .AddSheet(items.Skip(10))); // extra sheet

        File.WriteAllBytes($@"..\{nameof(TestDictionary)}.xlsx".ToLower(), excel);
    }

    // list of expandos 
    static void TestExpandoObject()
    {
        var items = Enumerable.Range(1, 100).Select(x =>
        {
            var item = new ExpandoObject();
            var itemDict = item as IDictionary<string, object>;
            itemDict.Add("Column #1", $"Text #{x}");
            itemDict.Add("Column #2", x * 1000);
            itemDict.Add("Column #3", DateTime.Now.AddDays(-x));
            return item;
        });

        var excel = items.ToExcel(s => s
            .AddSheet(items.Skip(10))); // extra sheet

        File.WriteAllBytes($@"..\{nameof(TestExpandoObject)}.xlsx", excel);
    }

    // list of hashtables
    static void TestHashtable()
    {
        var items = Enumerable.Range(1, 100).Select(x =>
        {
            var item = new Hashtable
            {
                { "Column #1", $"Text #{x}" },
                { "Column #2", x * 1000 },
                { "Column #3", DateTime.Now.AddDays(-x) }
            };
            return item;
        });

        var excel = items.ToExcel(s => s
            .AddSheet(items.Skip(10))); // extra sheet

        File.WriteAllBytes($@"..\{nameof(TestHashtable)}.xlsx", excel);
    }

    // DataTable
    static void TestDataTable()
    {
        var table = new DataTable("Table1");

        table.Columns.Add("Column #1", typeof(string));
        table.Columns.Add("Column #2", typeof(int));
        table.Columns.Add("Column #3", typeof(DateTime));

        for (var x = 1; x <= 100; x++)
            table.Rows.Add($"Text #{x}", x * 1000, DateTime.Now.AddDays(-x));

        var excel = table.ToExcel(s => s
            .AddSheet(table, ss => ss.SheetName("Table2"))
            .AddSheet(SomeItems));

        File.WriteAllBytes($@"..\{nameof(TestDataTable)}.xlsx", excel);
    }

    // list of objects
    static void TestObject()
    {
        var excel = SomeItems.AsEnumerable<object>().ToExcel(s => s
            .AddSheet(SomeItems.AsEnumerable<object>().Skip(3))); // extra sheet

        File.WriteAllBytes($@"..\{nameof(TestObject)}.xlsx", excel);
    }
}
