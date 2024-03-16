using Aspose.Cells;
using CsvHelper;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Collections.Generic;
using System.Xml.Linq;
using CsvHelper.Configuration.Attributes;
using CsvHelper.Configuration;
using System.Drawing;

namespace ExcelReading
{
    internal class Program
    {
        private static void Main()
        {
            CsvElementReader csvElementReader = new CsvElementReader("DataBase.csv");

            Console.WriteLine("Press any key...");
            Console.ReadKey();

            FileReader fileReader = new FileReader();
            fileReader.Run("test.xlsx");

            ExcelExport excelExport = new ExcelExport();
            excelExport.Export("exportTest.xlsx", fileReader.ElementsCounter);
        }

    }

    internal class FileReader
    {
        public ElementsCounter ElementsCounter;

        public FileReader()
        {
            ElementsCounter = new ElementsCounter();
        }
        public void Run(string filename)
        {
            DownloadFile(filename);
        }

        private void DownloadFile(string filename)
        {
            // Загрузить файл Excel
            Workbook workBook = new Workbook(filename);

            // Получить все рабочие листы
            WorksheetCollection collection = workBook.Worksheets;

            // Перебрать все рабочие листы
            for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
            {

                // Получить рабочий лист, используя его индекс
                Worksheet worksheet = collection[worksheetIndex];

                // Печать имени рабочего листа
                Console.WriteLine("Worksheet: " + worksheet.Name);

                // Получить количество строк и столбцов
                int rows = worksheet.Cells.MaxDataRow;
                int cols = worksheet.Cells.MaxDataColumn;

                // Цикл по строкам
                for (int i = 0; i < rows; i++)
                {
                    string[] rowData = new string[cols];
                    float count = 0;

                    // Перебрать каждый столбец в выбранной строке
                    for (int j = 0; j < cols; j++)
                    {
                        // Значение ячейки Pring
                        Console.Write(worksheet.Cells[i, j].Value + " | ");

                        if (worksheet.Cells[i, j].Value is string)
                        {
                            rowData[j] = (string)worksheet.Cells[i, j].Value;
                        }
                        else if (worksheet.Cells[i, j].Value is Int32)
                        {
                            rowData[j] = worksheet.Cells[i, j].Value.ToString();
                            count = Convert.ToSingle(worksheet.Cells[i, j].Value);
                        }
                    }

                    Console.WriteLine(" ");
                    //Console.WriteLine("Press any key");
                    //Console.ReadKey();

                    if (ElementsCounter.CountElement(rowData, count) == true)
                    {
                        worksheet.Cells[i, cols + 5].Value = "+";
                    }
                    else
                    {
                        worksheet.Cells[i, cols + 5].Value = "-";
                    }
                }
            }

            ElementsCounter.PrintInfo();

            
            workBook.Save("testEditted.xlsx");
        }
    }
}