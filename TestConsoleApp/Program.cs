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
            //    .AddColumn("MyColumn#2", x => x.Bool)
            //    .AddColumn("MyColumn#3", x => x.String));

            File.WriteAllBytes(@".\test.xlsx", data);
        }
    }
}
