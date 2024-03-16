using CsvHelper.Configuration.Attributes;

namespace ExcelReading
{
    internal class ElementCsvWrapper
    {
        static readonly char s_splitter = '#';

        [Name("Name")]
        public string Name { get; set; }

        [Name("NameVariants")]
        public string NameVariants { get; set; }

        [Name("Type")]
        public string Type { get; set; }

        [Name("TypeVariants")]
        public string TypeVariants { get; set; }

        [Name("Code")]
        public string Code { get; set; }

        [Name("CodeVariants")]
        public string CodeVariants { get; set; }

        [Name("Unit")]
        public string Unit { get; set; }

        [Name("UnitVariants")]
        public string UnitVariants { get; set; }

        [Name("Square")]
        public string Square { get; set; }

        public Element ConvertToElement()
        {
            Element newElement;

            string name = CheckEmptyString(Name);
            List<string> nameVariants = ElementCsvWrapper.ConvertVariants(NameVariants);

            string type = CheckEmptyString(Type);
            List<string> typeVariants = ElementCsvWrapper.ConvertVariants(TypeVariants);

            string code = CheckEmptyString(Code);
            List<string> codeVariants = ElementCsvWrapper.ConvertVariants(CodeVariants);

            string unit = CheckEmptyString(Unit);
            List<string> unitVariants = ElementCsvWrapper.ConvertVariants(UnitVariants);

            float square = CheckSquareData(Square);

            ElementField nameField = new ElementField(name, nameVariants);

            ElementField typeField = null;

            if (type != null && typeVariants != null)
                typeField = new ElementField(type, typeVariants);

            ElementField codeField = null;

            if(code != null && codeVariants != null)
                codeField = new ElementField(code, codeVariants);

            ElementField unitField = new ElementField(unit, unitVariants);

            newElement = new Element(nameField, typeField, codeField, unitField, square);

            return newElement;
        }

        private static string CheckEmptyString(string dataString)
        {
            if (dataString == "" || dataString == null)
                return null;

            return dataString;
        }

        private static List<string> ConvertVariants(string variants)
        {
            List<string> list = new List<string>();

            if (variants == "" || variants == null)
                return list;

            string[] strings = variants.Split(s_splitter);

            foreach (string splittedString in strings)
            {
                list.Add(splittedString);
            }

            return list;
        }

        private static float CheckSquareData(string data)
        {
            float result = 0;

            bool success = float.TryParse(data, out result);
            if (success)
            {
                if(result == 0)
                {
                    Console.WriteLine("<========> Нулевое значение площади элемента <========>");
                    Console.WriteLine("Press any key");
                    Console.ReadKey();
                }
                return result;
            }
            else
            {
                Console.WriteLine("<========> Площадь элемента не распознана <========>");
                return result;
            }
        }    
    }
}
