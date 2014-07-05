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
    public class CodeModule : SyntaxTreeNode
    {
        public CodeModule(string fileName)
            :base(SyntaxTreeNodeType.CodeFileTree, fileName)
        {
            _fileName = fileName;
        }

        private readonly string _fileName;
        public string FileName { get { return _fileName; } }

        private AttributeNode vbNameAttributeNode
        {
            get { return Nodes.OfType<AttributeNode>().FirstOrDefault(node => node.NodeName == "VB_Name"); }
        }

        public string NodeName
        { 
            get 
            {
                var node = vbNameAttributeNode;
                return node == null ? string.Empty : node.Value;                
            } 
            set
            {
                var node = vbNameAttributeNode;
                node.Value = value;
            }
        }
    }

    public class ClassModule : CodeModule
    {
        public ClassModule(string fileName)
            :base(fileName)
        {
        }

        public int MultiUse { get; set; }
        public int Persistable { get; set; }
        public int DataBindingBehavior { get; set; }
        public int DataSourceBehavior { get; set; }
        public int MTSTransactionMode { get; set; }
    }
}
