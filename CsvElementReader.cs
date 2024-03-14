using CsvHelper.Configuration;
using CsvHelper;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReading
{
    internal class CsvElementReader
    {
        public CsvConfiguration csvConfiguration;

        public CsvElementReader(string filePath)
        {
            csvConfiguration = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ";",
            };

            using (StreamReader streamReader = new StreamReader(filePath))
            using (CsvReader csv = new CsvReader(streamReader, csvConfiguration))
            {
                var elements = csv.GetRecords<ElementCsvWrapper>().ToList();
                foreach (ElementCsvWrapper element in elements)
                {
                    Console.WriteLine(element.Name + element.Type + element.TypeVariants + element.Code + element.CodeVariants + element.Unit + element.UnitVariants);
                }

                foreach (ElementCsvWrapper element in elements)
                {
                    Element newElement = element.ConvertToElement();
                    ElementsVariants.s_Elements.Add(newElement);
                }
            }

        }
    }
}
