using System;
using System.Collections.Generic;

namespace ExcelPOC
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                List<Package> packages =
                    new List<Package>
                        { new Package { Company = "Coho Vineyard", Weight = 25.2,
                              TrackingNumber = 89453312L,
                              DateOrder = DateTime.Today, HasCompleted = false },
                          new Package { Company = "Lucerne Publishing", Weight = 18.7,
                              TrackingNumber = 89112755L,
                              DateOrder = DateTime.Today, HasCompleted = false },
                          new Package { Company = "Wingtip Toys", Weight = 6.0,
                              TrackingNumber = 299456122L,
                              DateOrder = DateTime.Today, HasCompleted = false },
                          new Package { Company = "Adventure Works", Weight = 33.8,
                              TrackingNumber = 4665518773L,
                              DateOrder =  DateTime.Today.AddDays(-4),
                              HasCompleted = true },
                          new Package { Company = "Test Works", Weight = 35.8,
                              TrackingNumber = 4665518774L,
                              DateOrder =  DateTime.Today.AddDays(-2),
                              HasCompleted = true },
                          new Package { Company = "Good Works", Weight = 48.8,
                              TrackingNumber = 4665518775L,
                              DateOrder =  DateTime.Today.AddDays(-1), HasCompleted = true },

                        };

                var tuples = new[]
                {
                    new Tuple<string, bool, int>("alpha", false, 5),
                    new Tuple<string, bool, int>("beta", true, 11),
                    new Tuple<string, bool, int>("epsilon", false, 10),
                    new Tuple<string, bool, int>("gamma", true, 1)
                };

                var excel = new Excel();

                excel.AddSheet<Package>("SheetName")
                    .Map("A", "Alphabet", p => p.DateOrder, "yyyy-MM-dd")
                    .Map("B", "Bim", p => p.Company)
                    .Map("C", "", p => p.Weight, "0.00")
                    .Map("D", "Damned", p => p.TrackingNumber)
                    .Map("E", "Epsilon", p => (int)p.Weight % 2 == 0)
                    .SetData(packages);

                excel.AddSheet<Tuple<string, bool, int>>("TupleSheet")
                    .Map("A", t => t.Item1)
                    .Map("B", t => t.Item3)
                    .Map("C", t => t.Item2)
                    .Map("D", t => t.Item3 * (t.Item2 ? 2 : 3))
                    .SetData(tuples);


                excel.SaveTo("grab.xlsx");
                

                Console.WriteLine("Completed");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            Console.Read();
        }
    }

    

    public class Package
    {
        public string Company { get; set; }
        public double Weight { get; set; }
        public long TrackingNumber { get; set; }
        public DateTime DateOrder { get; set; }
        public bool HasCompleted { get; set; }
    }
}
