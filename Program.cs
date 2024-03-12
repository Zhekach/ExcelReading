using Aspose.Cells;
using CsvHelper;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Collections.Generic;
using System.Xml.Linq;
using CsvHelper.Configuration.Attributes;
using CsvHelper.Configuration;

internal class Program
{
    private static void Main()
    {
        CsvElementReader csvElementReader = new CsvElementReader("DataBase.csv");
        FileReader fileReader = new FileReader();
        fileReader.Run("test.xlsx");
    }

}

internal class FileReader
{
    private ElementsCounter elementsCounter;

    public FileReader()
    {
        elementsCounter = new ElementsCounter();
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
                // Распечатать разрыв строки
                Console.WriteLine(" ");

                elementsCounter.CountElement(rowData, count);
            }
        }

        elementsCounter.PrintInfo();
    }
}

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
            foreach(ElementCsvWrapper element in elements)
            {
                Console.WriteLine(element.Name + element.Type + element.Code + element.Unit);
            }
        }
        
    }
}

internal class ElementCsvWrapper
{ 
    [Name("Name")]
    public string Name { get; set; }
    [Name("Type")]
    public string Type { get; set; }
    [Name("Code")]
    public string Code { get; set; }
    [Name("Unit")]
    public string Unit { get; set; }

}

internal class ElementsCounter
{
    public Dictionary<Element, float> ElementsCounts = new Dictionary<Element, float>();

    public ElementsCounter()
    {
        foreach (Element element in ElementsVariants.s_elements)
        {
            ElementsCounts[element] = 0;
        }
    }

    public Element FindElementInRow(string[] rowData)
    {
        Element? newElement = null;

        bool isNameOk = false;
        bool isTypeOk = false;
        bool isCodeOk = false;
        bool isUnitOk = false;

        for (int i = 0; i < rowData.Length; i++)
        {
            if (isNameOk == false && rowData[i] != "")
            {
                foreach (Element element in ElementsCounts.Keys)
                {
                    if (element.Name.Equals(rowData[i]))
                    //if (string.Equals(rowData[i], element.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        newElement = element;
                        isNameOk = true;
                    }
                }
            }

            if (isNameOk)
            {
                if (newElement.Type == null)
                {
                    isTypeOk = true;
                }
                else if (newElement.Type.Equals(rowData[i]))
                //if (string.Equals(rowData[i], newElement.Type, StringComparison.OrdinalIgnoreCase))
                {
                    isTypeOk = true;
                }
            }

            if (isNameOk && isTypeOk)
            {
                if (newElement.Code == null)
                {
                    isCodeOk = true;
                }
                else if (newElement.Code.Equals(rowData[i]))
                //if (string.Equals(rowData[i], newElement.Code, StringComparison.OrdinalIgnoreCase))
                {
                    isCodeOk = true;
                }
            }

            if (isNameOk && isTypeOk && isCodeOk)
            {
                if (newElement.Unit.Equals(rowData[i]))
                //if (string.Equals(rowData[i], newElement.Unit, StringComparison.OrdinalIgnoreCase))
                {
                    isUnitOk = true;
                    Console.WriteLine("=============\n" +
                                      "Элемент найден!!\n" +
                                      "=============");
                    Console.WriteLine("Press any key");
                    Console.ReadKey();

                    break;
                }
            }
        }

        if (isUnitOk == false || isCodeOk == false || isTypeOk == false || isNameOk == false)
        {
            newElement = null;
        }

        return newElement;
    }

    public void CountElement(string[] rowData, float count)
    {
        Element newElement = FindElementInRow(rowData);

        if (newElement != null && ElementsCounts.ContainsKey(newElement))
        {
            ElementsCounts[newElement] += count;
        }
    }

    public void PrintInfo()
    {
        foreach (KeyValuePair<Element, float> keyValuePair in ElementsCounts)
        {
            Console.WriteLine($"Элемент {keyValuePair.Key} в количестве:{keyValuePair.Value}");
        }
    }
}

internal class Element
{

    public ElementField Name { get; private set; }
    public ElementField Type { get; private set; }
    public ElementField Code { get; private set; }
    public ElementField Unit { get; private set; }

    public Element(ElementField name, ElementField type, ElementField code, ElementField unit)
    {
        Name = name;
        Type = type;
        Code = code;
        Unit = unit;
    }

    public override string ToString()
    {
        return $"Название {Name}, тип {Type} код {Code} единица {Unit}.";
    }

    public Element Clone()
    {
        return new Element(this.Name, this.Type, this.Code, this.Unit);
    }
}

internal class ElementField
{
    public string Name { get; private set; }
    public List<string> Variants { get; private set; }

    public ElementField(string name)
    {
        Name = name;
        Variants = new List<string>();
    }

    public ElementField(string name, List<string> variants)
    {
        Name = name;
        Variants = variants;
    }

    public override string ToString()
    {
        return Name;
    }

    public override bool Equals(object? obj)
    {
        if (obj is string == false || obj == null && Name != null)
        {
            return false;
        }

        if (obj == null && Name == null)
        {
            return true;
        }
        else if (string.Equals(obj.ToString(), Name, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        foreach (string variant in Variants)
        {
            if (string.Equals(obj.ToString(), variant, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }


        return false;
    }

    public void AddVariant(string newVariant)
    {
        Variants.Add(newVariant);
    }

    public void RemoveVariant(string newVariant)
    {
        Variants.Remove(newVariant);
    }
}

internal class ElementsVariants
{
    public readonly static List<Element> s_elements = new List<Element>()
    {
        {new Element(new ElementField("Переход прямоугольного сечения", new List<string>{"Переход прямоугольногА сечения", "Переходной"}),
            null,
            new ElementField("П-500х350-350х500", new List<string>{"П-500х350-350х500"}),
            new ElementField("шт.")) }
    };
}