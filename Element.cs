using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReading
{
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
}
