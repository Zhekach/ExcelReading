using Aspose.Cells;
using System.Collections.Generic;

internal class Program
{
    private static void Main()
    {
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
                string[] rowData = new string[rows];
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

internal class ElementsCounter
{
    public Dictionary<Element, float> ElementsCounts = new Dictionary<Element, float>();

    public Element? TestElement;
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
                    if (string.Equals(rowData[i], element.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        newElement = element;
                        isNameOk = true;
                    }
                }
            }

            if (isNameOk)
            {
                if (string.Equals(rowData[i], newElement.Type, StringComparison.OrdinalIgnoreCase))
                {
                    isTypeOk = true;
                }
            }

            if (isNameOk && isTypeOk)
            {
                if (string.Equals(rowData[i], newElement.Code, StringComparison.OrdinalIgnoreCase))
                {
                    isCodeOk = true;
                }
            }

            if (isNameOk && isTypeOk && isCodeOk)
            {
                if (string.Equals(rowData[i], newElement.Unit, StringComparison.OrdinalIgnoreCase))
                {
                    isUnitOk = true;
                    Console.WriteLine("=============\n" +
                                      "Элемент найден!!\n" +
                                      "=============");
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

    public string Name { get; private set; }
    public string Type { get; private set; }
    public string Code { get; private set; }
    public string Unit { get; private set; }

    public Element(string name, string type, string code, string unit)
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

internal class ElementsVariants
{
    public readonly static List<Element> s_elements = new List<Element>()
    {
        {new Element("Переход прямоугольного сечения", null, "П-630х550-500х350", "шт.") }
    };
}