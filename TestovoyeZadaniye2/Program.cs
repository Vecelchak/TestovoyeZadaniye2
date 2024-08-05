using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data;
using System.Data.SqlClient;


namespace TestovoyeZadaniye2
{
    internal class Program
    {
        static void Main(string[] args)
        
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                Console.Write("Введите количество элементов массива: ");
                int n = int.Parse(Console.ReadLine());
                string[] array = GenerateArray(n);
                string[] sortedAsc = array.OrderBy(x => x).ToArray();
                string[] sortedDesc = array.OrderByDescending(x => x).ToArray();
                SaveToExcel(array, sortedAsc, sortedDesc);

            }

            static string[] GenerateArray(int n)
            {
                string[] array = new string[n];
                Random random = new Random();
                for (int i = 0; i < n; i++)
                {
                    int length = random.Next(10, 20);
                    string str = "";
                    for (int j = 0; j < length; j++)
                    {
                        char c = (char)random.Next('0', '9' + 1);
                        str += c;
                        c = (char)random.Next('A', 'Z' + 1);
                        str += c;
                    }
                    array[i] = str;
                }
                return array;
            }
            static void SaveToExcel(string[] original, string[] sortedAsc, string[] sortedDesc)
            {
                Console.Write("Введите путь для сохранения файла (например C:\\Users\\User\\Desktop): ");
                string path = Console.ReadLine();
                string fileName = "SortedData-" + DateTime.Now.ToString("yyyy-MMMM-dd HH-mm-ss") + ".xlsx";
                string filePath = path + "\\" + fileName;
                using (ExcelPackage excel = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("SortedData");
                    worksheet.Cells[1, 1].Value = "Начальные сгенерированные данные";
                    worksheet.Cells[1, 2].Value = "Отсортированные по возрастанию";
                    worksheet.Cells[1, 3].Value = "Отсортированные по убыванию";

                    for (int i = 0; i < original.Length; i++)
                    {
                        worksheet.Cells[i + 2, 1].Value = original[i];
                        worksheet.Cells[i + 2, 2].Value = sortedAsc[i];
                        worksheet.Cells[i + 2, 3].Value = sortedDesc[i];
                    }
                    using (var range = worksheet.Cells[1, 1, original.Length + 1, 3])
                    {
                        range.AutoFitColumns();
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                    FileInfo excelFile = new FileInfo(filePath);
                    excel.SaveAs(excelFile);
                    Console.WriteLine("Файл " + fileName + " успешно сохранен по адресу: " + filePath);
                }
            }

        }
    }


