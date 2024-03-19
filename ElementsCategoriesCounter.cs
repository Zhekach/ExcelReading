using System.Collections.Generic;

namespace ExcelReading
{
    internal class ElementsCategoriesCounter
    {
        public Dictionary<string, float> CategoriesCount;
        public ElementsCounter ElementsCounter { get; private set; }

        public ElementsCategoriesCounter(ElementsCounter elementsCounter)
        {
            CategoriesCount = new Dictionary<string, float>();
            ElementsCounter = elementsCounter;
        }

        public void CountElements()
        {
            foreach(KeyValuePair<Element, float> kvp in ElementsCounter.ElementsCounts) 
            {
                string newCategory;
                newCategory = kvp.Key.Category;

                if(CategoriesCount.ContainsKey(newCategory))
                {
                    CategoriesCount[newCategory] += kvp.Value;
                }
                else 
                { 
                    CategoriesCount.Add(newCategory, kvp.Value); 
                }
            }
        }

        public void PrintInfo()
        {
            foreach (KeyValuePair<string, float> keyValuePair in CategoriesCount)
            {
                Console.Write($"Категория {keyValuePair.Key} в количестве:{keyValuePair.Value}");

                if (keyValuePair.Value == 0)
                    Console.Write("<==========> NULL VALUE <============>");

                Console.WriteLine();
            }
        }
    }
}
