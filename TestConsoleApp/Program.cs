using RandomSolutions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace TestConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Example1();
            Example2();
            TestTypes();
        }

        static void Example1()
        {
            var items = Enumerable.Range(1, 10).Select(x => new
            {
                Prop1 = $"Text #{x}",
                Prop2 = x * 1000,
                Prop3 = DateTime.Now.AddDays(-x),
            });

            var excel = items.ToExcel();

            File.WriteAllBytes(@"..\..\..\..\Examples\example1.xlsx", excel);
        }

        static void Example2()
        {
            var items = Enumerable.Range(1, 10).Select(x => new
            {
                Prop1 = $"Text #{x}",
                Prop2 = x * 1000,
                Prop3 = DateTime.Now.AddDays(-x),
            });

            var excel = items.ToExcel(scheme => scheme
                .AddColumn("MyColumnName#1", x => x.Prop1)
                .AddColumn("MyColumnName#2", x => $"test:{x.Prop2}")
                .AddColumn("MyColumnName#3", x => x.Prop3));

            File.WriteAllBytes(@"..\..\..\..\Examples\example2.xlsx", excel);
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
                String = $"text text text #{x}",
            });

            //var data = items.ToExcel("Test!");

            var data = items.ToExcel(x =>
            {
                x.SheetName = "Test!";
                //x.AddColumn("MyColumn#1", x => x.Int);
                //x.AddColumn("MyColumn#2", x => x.Bool, 40);
                //x.AddColumn("MyColumn#3", x => x.String);
            });

            File.WriteAllBytes(@"..\test.xlsx", data);
        }
    }
}
