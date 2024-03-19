using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;

namespace ExcelReading
{
    internal class ExcelExport
    {
        public void Export(string fileName, ElementsCounter elementsCounter)
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            int rowNumber = 1;

            worksheet.Cells[0, 0].PutValue("Категория");
            worksheet.Cells[0, 1].PutValue("Название");
            worksheet.Cells[0, 2].PutValue("Тип");
            worksheet.Cells[0, 3].PutValue("Код");
            worksheet.Cells[0, 4].PutValue("Единица измерения");
            worksheet.Cells[0, 5].PutValue("Площадь, м2");
            worksheet.Cells[0, 6].PutValue("Количество");
            worksheet.Cells[0, 7].PutValue("Площадь всего");

            foreach(KeyValuePair<Element, float> keyValuePair in elementsCounter.ElementsCounts)
            {
                worksheet.Cells[rowNumber, 0].PutValue(keyValuePair.Key.Category);
                worksheet.Cells[rowNumber, 1].PutValue(keyValuePair.Key.Name);
                worksheet.Cells[rowNumber, 2].PutValue(keyValuePair.Key.Type);
                worksheet.Cells[rowNumber, 3].PutValue(keyValuePair.Key.Code);
                worksheet.Cells[rowNumber, 4].PutValue(keyValuePair.Key.Unit);
                worksheet.Cells[rowNumber, 5].PutValue(keyValuePair.Key.Square);
                worksheet.Cells[rowNumber, 6].PutValue(keyValuePair.Value);
                worksheet.Cells[rowNumber, 7].PutValue(keyValuePair.Value * keyValuePair.Key.Square);

                rowNumber++;
            }

            worksheet.AutoFitRows();
            worksheet.AutoFitColumns();

            workbook.Save(fileName);
        }
    }
}
