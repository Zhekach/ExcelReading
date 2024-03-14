namespace ExcelReading
{
    internal class ElementsCounter
    {
        public Dictionary<Element, float> ElementsCounts = new Dictionary<Element, float>();

        public ElementsCounter()
        {
            foreach (Element element in ElementsVariants.s_Elements)
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

            foreach (Element element in ElementsCounts.Keys)
            {
                for (int i = 0; i < rowData.Length; i++)
                {
                    if (isNameOk == false && rowData[i] != "")
                    {
                        if (element.Name.Equals(rowData[i]))
                        {
                            newElement = element;
                            isNameOk = true;
                        }
                    }

                    if (isNameOk)
                    {
                        if (newElement.Type == null)
                        {
                            isTypeOk = true;
                        }
                        else if (newElement.Type.Equals(rowData[i]))
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
                        {
                            isCodeOk = true;
                        }
                    }

                    if (isNameOk && isTypeOk && isCodeOk)
                    {
                        if (newElement.Unit.Equals(rowData[i]))
                        {
                            isUnitOk = true;
                            Console.WriteLine("=============\n" +
                                              "Элемент найден!!\n" +
                                              "=============");
                            Console.WriteLine("Press any key");
                            Console.ReadKey();

                            return newElement;
                        }
                    }
                }

                isNameOk = false;
                isTypeOk = false;
                isCodeOk = false;
                isUnitOk = false;
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
}