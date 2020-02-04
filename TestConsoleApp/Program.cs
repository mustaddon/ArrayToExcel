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
                Id = x,
                Name01 = $"Name01#{x}",
                Name02 = $"Name02#{x}",
                Name03 = $"Name03#{x}",
                Name04 = $"Name04#{x}",
                Name05 = $"Name05#{x}",
                Name06 = $"Name06#{x}",
                Name07 = $"Name07#{x}",
                Name08 = $"Name08#{x}",
                Name09 = $"Name09#{x}",
                Name10 = $"Name10#{x}",
                Name11 = $"Name11#{x}",
                Name12 = $"Name12#{x}",
                Name13 = $"Name13#{x}",
                Name14 = $"Name14#{x}",
                Name15 = $"Name15#{x}",
                Name16 = $"Name16#{x}",
                Name17 = $"Name17#{x}",
                Name18 = $"Name18#{x}",
                Name19 = $"Name19#{x}",
                Name20 = $"Name20#{x}",
                Name21 = $"Name21#{x}",
                Name22 = $"Name22#{x}",
                Name23 = $"Name23#{x}",
                Name24 = $"Name24#{x}",
                Name25 = $"Name25#{x}",
                Name26 = $"Name26#{x}",
                Name27 = $"Name27#{x}",
                Name28 = $"Name28#{x}",
                Name29 = $"Name29#{x}",
                Name30 = $"Name30#{x}",
                Name31 = $"Name31#{x}",
                Name32 = $"Name32#{x}",
                Name33 = $"Name33#{x}",
                Name34 = $"Name34#{x}",
                Name35 = $"Name35#{x}",
                Name36 = $"Name36#{x}",
                Name37 = $"Name37#{x}",
                Name38 = $"Name38#{x}",
                Name39 = $"Name39#{x}",
            });

            //var data = items.ToExcel(x => x.AddColumn("кол 2", x => x.Name01).AddColumn("кол 1", x => x.Id));
            var data = items.ToExcel();

            File.WriteAllBytes(@".\test.xlsx", data);
        }
    }

    class Test
    {
        public int Id { get; set; }
        public string Name;
        public static string Stat;
        public void Meth() { }
    }
}
