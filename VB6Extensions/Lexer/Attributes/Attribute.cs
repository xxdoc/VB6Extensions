using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VB6Extensions.Lexer.Attributes
{
    public interface IAttribute
    {
        string Name { get; }
        string Value { get; set; }
    }

    public class Attribute : IAttribute
    {
        public Attribute(string name, string value)
        {
            _name = name;
            _value = value;
        }

        private readonly string _name;
        public string Name { get { return _name; } }

        private string _value;
        public string Value { get { return _value; } set { _value = value; } }
    }

    public class MemberAttribute : Attribute
    {
        public MemberAttribute(string name, string value, string member)
            : base(name, value)
        {
            Member = member;
        }

        public string Member { get; private set; }
    }
}
