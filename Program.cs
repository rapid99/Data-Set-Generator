using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace DataSetGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Data Set Generator");
            Console.WriteLine("--------------------");
            Console.WriteLine("What is this data set called?");
            var title = Console.ReadLine();
            Console.WriteLine("--------------------");
            var colAmount = GetColumnAmount();         
            Console.WriteLine("--------------------");
            var cols = new List<KeyValuePair<string, string>>();
            for (var i = 0; i < colAmount; i++)
            {
                Console.WriteLine($"Column {i + 1} name:");
                var name = Console.ReadLine();
                Console.WriteLine($"Column {i + 1} data type:");
                //TODO: validate data types
                var dataType = Console.ReadLine();
                cols.Add(new KeyValuePair<string, string>(name, dataType));
                Console.WriteLine("--------------------");
            }
            Console.WriteLine($"Collected {cols.Count} column names");
            Console.WriteLine("--------------------");
            Console.WriteLine("How many rows?");
            var rowsCount = Console.ReadLine();
            Console.WriteLine("---Summary---");
            Console.WriteLine($"Data set name: {title}");
            Console.WriteLine($"Rows to create: {rowsCount}");
            Console.WriteLine("Columns:");
            Console.WriteLine("--------------------");
            Console.WriteLine($"Total: {colAmount}");
            Console.WriteLine("--------------------");
            for (var i = 0; i < cols.Count; i++)
            {
                Console.WriteLine($"Column {i + 1}");
                Console.WriteLine($"Name: {cols[i].Key}");
                Console.WriteLine($"Date Type: {cols[i].Value}");
                Console.WriteLine("--------------------");
            }
            ConfirmSettings();
            Console.WriteLine("Generating data set.......");
            GeneratorFile(title, cols, Convert.ToInt32(rowsCount));

            Console.WriteLine(@$"Data has been saved to C:\temp\{title}.xlsx");
            Console.WriteLine("Start over? (Y)");
            if (Console.ReadKey().Key == ConsoleKey.Y)
                Main(null);
        }

        private static int GetColumnAmount()
        {
            Console.WriteLine("How many columns?");
            var colsInput = Console.ReadLine();
            int.TryParse(colsInput, out int cols);
            if (cols == 0)
            {
                Console.WriteLine("Invalid column number, try again");
                GetColumnAmount();
            }

            return cols;
        }

        private static void ConfirmSettings()
        {
            Console.WriteLine("Continue with above settings (Y) or start over (S)");
            var settingsOk = Console.ReadKey();
            switch (settingsOk.Key)
            {
                case ConsoleKey.Y:
                    break;
                case ConsoleKey.S:
                    Main(null);
                    break;
                default:
                    Console.WriteLine("Invalid choice, try again");
                    Console.WriteLine("--------------------");
                    ConfirmSettings();
                    break;
            }

        }

        private static void GeneratorFile(string name, List<KeyValuePair<string, string>> cols, int totalRows)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(@$"c:\temp\{name}.xlsx");
            using (var package = new ExcelPackage(file))
            {
                var sheet = package.Workbook.Worksheets.Add($"{name}");
                ExcelWorksheet workSheet = package.Workbook.Worksheets[0];

                for (int i = 0; i < cols.Count; i++)
                {
                    sheet.Cells[1, i + 1].Value = cols[i].Key;
                    sheet.Cells[1, i + 1].Style.Font.Bold = true;
                }
                              
                for (int i = 2; i <= totalRows; i++)
                {
                    for (int c = 0; c < cols.Count; c++)
                    {
                        switch (cols[c].Value)
                        {
                            case "string":
                                sheet.Cells[i, c + 1].Value = Faker.Name.First();
                                break;
                            case "int":
                                sheet.Cells[i, c + 1].Value = Faker.RandomNumber.Next(100);
                                break;
                            case "bool":
                                sheet.Cells[i, c + 1].Value = Faker.Boolean.Random().ToString();
                                break;
                            case "ssn":
                                sheet.Cells[i, c + 1].Value = Faker.Identification.SocialSecurityNumber();
                                break;
                            case "company":
                                sheet.Cells[i, c + 1].Value = Faker.Company.Name();
                                break;
                            default:
                                sheet.Cells[i, c + 1].Value = Faker.Country.Name();
                                break;
                        }
                    }
                }

                package.Save();
            }
        }
    }
}