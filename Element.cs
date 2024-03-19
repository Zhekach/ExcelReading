namespace ExcelReading
{
    internal class Element
    {
        public string Category { get; private set; }
        public ElementField Name { get; private set; }
        public ElementField Type { get; private set; }
        public ElementField Code { get; private set; }
        public ElementField Unit { get; private set; }

        public float Square;

        public Element(string category, ElementField name, ElementField type, ElementField code, ElementField unit, float square)
        {
            Category = category;
            Name = name;
            Type = type;
            Code = code;
            Unit = unit;
            Square = square;
        }

        public override string ToString()
        {
            return $"Категория {Category}, Название {Name}, тип {Type} код {Code} единица {Unit}, площадь {Square} м2.";
        }

        public Element Clone()
        {
            return new Element(this.Category, this.Name, this.Type, this.Code, this.Unit, this.Square);
        }
    }
}
