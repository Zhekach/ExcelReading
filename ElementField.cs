using System.Text.RegularExpressions;

namespace ExcelReading
{
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

            if(variants == null)
                Variants = new List<string>();
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
            //else if(string.Equals(Normalize(obj.ToString()), Normalize(Name), StringComparison.OrdinalIgnoreCase))
            //{
            //    return true;
            //}

            foreach (string variant in Variants)
            {
                if (string.Equals(obj.ToString(), variant, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }


            return false;
        }

        private string Normalize (string value)
        {
            string pattern = ("-? | ‐? | \\??");
            //string pattern = "(Mr\\.? |Mrs\\.? |Miss |Ms\\.? )";
            //           string patternX = ("-? | ‐?");


            Regex.Replace (value, pattern, "-");

            return value;
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
}
