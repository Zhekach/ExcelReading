using System.Collections.Generic;

namespace ExcelReading
{
    internal class ElementsCategoriesCounter
    {
        public Dictionary<string, CategoryCounterValues> CategoriesCount;
        public ElementsCounter ElementsCounter { get; private set; }

        public ElementsCategoriesCounter(ElementsCounter elementsCounter)
        {
            CategoriesCount = new Dictionary<string, CategoryCounterValues>();
            ElementsCounter = elementsCounter;
        }

        public void CountElements()
        {
            foreach (KeyValuePair<Element, float> kvp in ElementsCounter.ElementsCounts)
            {
                string newCategory;
                newCategory = kvp.Key.Category;

                if (CategoriesCount.ContainsKey(newCategory))
                {
                    CategoriesCount[newCategory].Value += kvp.Value;
                    CategoriesCount[newCategory].Square += kvp.Key.Square;
                }
                else
                {
                    CategoriesCount.Add(newCategory, new CategoryCounterValues(kvp.Value, kvp.Key.Square));
                }
            }
        }

        public void PrintInfo()
        {
            foreach (KeyValuePair<string, CategoryCounterValues> keyValuePair in CategoriesCount)
            {
                Console.Write($"Категория {keyValuePair.Key} в количестве:{keyValuePair.Value.Value}, " +
                    $"общей площадью: {keyValuePair.Value.Square} м2");

                if (keyValuePair.Value.Value == 0 || keyValuePair.Value.Square == 0)
                    Console.Write("<==========> NULL VALUE <============>");

                Console.WriteLine();
            }
        }
    }
}
