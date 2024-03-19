using CsvHelper.Configuration;
using CsvHelper;
using System.Globalization;

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

                Console.WriteLine("Распознаны следующие элементы:");

                foreach (ElementCsvWrapper element in elements)
                {
                    Element newElement = element.ConvertToElement();
                    ElementsVariants.s_Elements.Add(newElement);

                    Console.WriteLine(newElement);
                }
            }

        }
    }
}
