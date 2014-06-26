using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VB6Extensions.Lexer.Attributes;
using VB6Extensions.Lexer.Tokens;
using VB6Extensions.Parser;
using VB6Extensions.Properties;

namespace VB6Extensions
{
    public class CodeModule : ISyntaxTree
    {
        public IEnumerable<IAttribute> Attributes { get; private set; }
        public IList<ISyntaxTree> Nodes { get; private set; }

        public CodeModule(string fileName, IEnumerable<IAttribute> attributes)
        {
            IAttribute nameAttribute;
            try
            {
                nameAttribute = attributes.First();
                if (string.IsNullOrEmpty(nameAttribute.Value))
                    throw new FormatException("[VB_Name] attribute value cannot be empty.");
            }
            catch (NullReferenceException exception)
            {
                throw new ArgumentException("CodeModule requires a [VB_Name] attribute, which must be specified first.", exception);
            }

            _fileName = fileName;

            Attributes = attributes;
            Nodes = new List<ISyntaxTree>();
        }

        private readonly string _fileName;
        public string FileName { get { return _fileName; } }

        public string Name 
        { 
            get 
            {
                return Attributes.First().Value.Replace("\"", string.Empty); 
            } 
            set
            {
                Attributes.First().Value = value;
            }
        }
    }

    public class ClassModule : CodeModule
    {
        public ClassModule(string fileName, IEnumerable<IAttribute> attributes)
            :base(fileName, attributes)
        {
        }

        public int MultiUse { get; set; }
        public int Persistable { get; set; }
        public int DataBindingBehavior { get; set; }
        public int DataSourceBehavior { get; set; }
        public int MTSTransactionMode { get; set; }
    }
}
