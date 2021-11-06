using ArrayToExcel;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Example1();
            Example2();
            Example3();
            Example4();
            Example5();
            Example6();
            Example7();
            TestTypes();
        }

        static IEnumerable<SomeItem> SomeItems = Enumerable.Range(1, 10).Select(x => new SomeItem
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
                .AddColumn("MyColumnName#1", x => x.Prop1)
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

        // list of dictionaries 
        static void Example6()
        {
            var items = Enumerable.Range(1, 100).Select(x => new Dictionary<object, object>
            {
                { "Column #1", $"Text #{x}" },
                { "Column #2", x * 1000 },
                { "Column #3", DateTime.Now.AddDays(-x) },
            });

            var excel = items.ToExcel();

            File.WriteAllBytes($@"..\..\..\..\{nameof(Example6)}.xlsx".ToLower(), excel);
        }

        // list of expandos 
        static void Example7()
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

            var excel = items.ToExcel();

            File.WriteAllBytes($@"..\..\..\..\{nameof(Example7)}.xlsx".ToLower(), excel);
        }

        static void TestTypes()
        {
            var items = Enumerable.Range(1, 100).Select(x => new
            {
                Bool = x % 2 == 0,
                NullableBool = x % 2 == 0 ? true : (bool?)null,
                Int = -x * 100,
                Uint = (uint)x * 100,
                Long = (long)x * 100,
                Double = 1.1d + x,
                Float = 1.1f + x,
                Decimal = 1.1m + x,
                DateTime = DateTime.Now.AddDays(-x),
                DateTimeOffset = DateTimeOffset.Now.AddDays(-x),
                String = $"text text text #{x} !!!",
            });

            var data = items.ToExcel();

            File.WriteAllBytes($@"..\{nameof(TestTypes)}.xlsx", data);
        }




    }
}
