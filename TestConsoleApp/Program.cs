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

            var excel = items.ToExcel(x => x
                .AddColumn("MyColumnName#1", x => x.Prop1)
                .AddColumn("MyColumnName#2", x => $"test:{x.Prop2}")
                .AddColumn("MyColumnName#3", x => x.Prop3));

            File.WriteAllBytes(@"..\..\..\..\Examples\example2.xlsx", excel);
        }

        static void TestTypes()
        {
            var items = Enumerable.Range(1, 100).Select(x => new
            {
                Bool = true,
                Int = -1,
                Uint = 1u,
                Long = 1L,
                Double = 1.1d,
                Float = 1.1f,
                Decimal = 1.1m,
                DateTime = DateTime.Now,
                DateTimeOffset = DateTimeOffset.Now,
                String = $"Text#{x}",
            });

            var data = items.ToExcel();

            //var data = items.ToExcel(x => x
            //    .AddColumn("MyColumn#1", x => x.Int)
            //    .AddColumn("MyColumn#2", x => x.Bool,40)
            //    .AddColumn("MyColumn#3", x => x.String));

            File.WriteAllBytes(@".\test.xlsx", data);
        }
    }
}
