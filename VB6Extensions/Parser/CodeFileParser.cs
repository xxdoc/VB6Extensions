using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using VB6Extensions.Lexer.Attributes;
using VB6Extensions.Lexer.Tokens;
using VB6Extensions.Properties;

namespace VB6Extensions.Parser
{
    public interface ISyntaxTree
    {
        string Name { get; }
        IEnumerable<IAttribute> Attributes { get; }
        IList<ISyntaxTree> Nodes { get; }
    }

    public class CodeFileParser
    {
        public ISyntaxTree Parse (string fileName)
        {
            var content = File.ReadAllLines(fileName);
            var currentLine = 0;

            var header = ParseFileHeader(fileName, content, ref currentLine);
            var declarations = ParseDeclarations(content, ref currentLine);
            var members = ParseMembers(content, ref currentLine);

            var module = new ModuleNode(header, declarations, members);
            return module;
        }

        private ISyntaxTree ParseFileHeader(string fileName, string[] content, ref int currentLine)
        {
            var attributeParser = new AttributeParser();
            IList<IAttribute> attributes = new List<IAttribute>();
            ISyntaxTree result;

            var firstLine = content[0].Trim();
            if (firstLine == "VERSION 1.0 CLASS")
            {
                var multiUse = content[2].Trim();
                var persistable = content[3].Trim();
                var dataBindingBehavior = content[4].Trim();
                var dataSourceBehavior = content[5].Trim();
                var mtsTransactionMode = content[6].Trim();

                attributes.Add(attributeParser.Parse(content[8].Trim()));
                attributes.Add(attributeParser.Parse(content[9].Trim()));
                attributes.Add(attributeParser.Parse(content[10].Trim()));
                attributes.Add(attributeParser.Parse(content[11].Trim()));
                attributes.Add(attributeParser.Parse(content[12].Trim()));

                var regex = new Regex(@"\=\s\-?(?<IntValue>\d+)\s");
                result = new ClassModule(fileName, attributes)
                                {
                                    DataBindingBehavior = int.Parse(regex.Match(dataBindingBehavior).Groups["IntValue"].Value),
                                    DataSourceBehavior = int.Parse(regex.Match(dataSourceBehavior).Groups["IntValue"].Value),
                                    MTSTransactionMode = int.Parse(regex.Match(mtsTransactionMode).Groups["IntValue"].Value),
                                    MultiUse = int.Parse(regex.Match(multiUse).Groups["IntValue"].Value),
                                    Persistable = int.Parse(regex.Match(persistable).Groups["IntValue"].Value)
                                };

                currentLine = 13;
            }
            else
            {
                attributes.Add(attributeParser.Parse(content[0].Trim()));
                result = new CodeModule(fileName, attributes);

                currentLine = 1;
            }

            return result;
        }

        private IEnumerable<ISyntaxTree> ParseDeclarations(string[] content, ref int currentLine)
        {
            var result = new List<ISyntaxTree>();

            var pattern = @"((?<keyword>Dim|Static|Public|Private|Friend|Global)\s)?(?<keyword>Dim|Static|Public|Private|Friend|Global|Const|Declare|Type|Enum)\s+(?<identifier>\w+)(?<arraySize>\(.*\))?(\s+As\s+?(((?<initializer>New)\s+)?)(?<type>\w+(\.\w+)?))?$";
            var regex = new Regex(pattern);

            var isDeclarationSection = true;
            while (isDeclarationSection)
            {
                var line = content[currentLine];
                isDeclarationSection =  !(
                                               line.Contains(ReservedKeywords.Property)
                                            || line.Contains(ReservedKeywords.Sub)
                                            || line.Contains(ReservedKeywords.Function)
                                         );
                currentLine++;

                if (isDeclarationSection)
                {
                    var match = regex.Match(line);
                    if (match.Success)
                    {
                        if (match.Groups["keyword"].Captures.Count == 2)
                        {
                            result.Add(new DeclarationNode(match.Groups["keyword"].Captures[0].Value, 
                                                           match.Groups["keyword"].Captures[1].Value,
                                                           match.Groups["identifier"].Value,
                                                           match.Groups["arraySize"].Value,
                                                           match.Groups["initializer"].Value,
                                                           match.Groups["type"].Value));
                        }
                        else
                        {
                            result.Add(new DeclarationNode(null, match.Groups["keyword"].Value,
                                                                 match.Groups["identifier"].Value,
                                                                 match.Groups["arraySize"].Value,
                                                                 match.Groups["initializer"].Value,
                                                                 match.Groups["type"].Value));
                        }
                    }
                }
            }

            return result;
        }

        private IEnumerable<ISyntaxTree> ParseMembers(string[] content, ref int currentLine)
        {
            var result = new List<ISyntaxTree>();

            var pattern = @"((?<keyword>Public|Private|Friend)\s)?(?<keyword>Property|Function|Sub)\s+(?<keyword>Get|Let|Set)\s+(?<name>[a-zA-Z][a-zA-Z0-9_]*)(\((?<parameters>.*)\))?(\s+As\s+(?<type>.*))?$";
            var regex = new Regex(pattern);

            while (currentLine < content.Length)
            {
                var match = regex.Match(content[currentLine]);
                if (match.Success)
                {
                    var modifier = match.Groups["keyword"].Captures[0].Value;
                    if (!new[]{ ReservedKeywords.Sub, ReservedKeywords.Function, ReservedKeywords.Property }.Contains(modifier))
                    {
                        if (match.Groups["keyword"].Captures[1].Value == ReservedKeywords.Property)
                        {
                            var keyword = match.Groups["keyword"].Captures[2].Value;
                            var node = new PropertyNode(modifier, match.Groups["name"].Value,
                                                        keyword, match.Groups["parameters"].Value);
                            result.Add(node);
                        }
                    }
                    else
                    {
                        if (match.Groups["keyword"].Captures[0].Value == ReservedKeywords.Property)
                        {
                            var keyword = match.Groups["keyword"].Captures[1].Value;
                            var node = new PropertyNode(null, match.Groups["name"].Value,
                                                        keyword, match.Groups["parameters"].Value);
                            result.Add(node);
                        }
                    }
                }

                currentLine++;
            }

            return result;
        }
    }

    public class PropertyNode : ISyntaxTree
    {
        public PropertyNode(string modifier, string name, string keyword, string parameters)
        {
            Modifier = string.IsNullOrEmpty(modifier)
                ? (AccessModifier?)null
                : (AccessModifier)Enum.Parse(typeof(AccessModifier), modifier);

            Name = name;
            Accessor = keyword;

            Nodes = new List<ISyntaxTree>(); // todo: parse parameter nodes.
        }

        public AccessModifier? Modifier { get; private set; }

        public string Accessor { get; private set; }
        public string Name { get; private set; }

        public IEnumerable<IAttribute> Attributes { get; private set; }

        public IList<ISyntaxTree> Nodes { get; private set; }
    }

    public class ModuleNode : ISyntaxTree
    {
        public ModuleNode(ISyntaxTree header, IEnumerable<ISyntaxTree> declarations, IEnumerable<ISyntaxTree> members)
        {
            Name = header.Attributes.First().Value;
            Nodes = new List<ISyntaxTree>(declarations.Concat(members));
        }

        public string Name { get; private set; }

        public IEnumerable<IAttribute> Attributes { get; private set; }

        public IList<ISyntaxTree> Nodes { get; private set; }
    }

    public class DeclarationNode : ISyntaxTree
    {
        public DeclarationNode(string modifier, string keyword, string identifier, string arraySizeSpecifier, string initializer, string type)
        {
            Modifier = string.IsNullOrEmpty(modifier) 
                ? (AccessModifier?)null 
                : (AccessModifier)Enum.Parse(typeof(AccessModifier), modifier);

            Name = keyword;
            Nodes = new List<ISyntaxTree>
            {
                new IdentifierNode(identifier, arraySizeSpecifier, initializer, type)
            };
        }

        public AccessModifier? Modifier { get; private set; }
        public string Name { get; private set; }

        public IEnumerable<IAttribute> Attributes { get; private set; }

        public IList<ISyntaxTree> Nodes { get; private set; }
    }

    public class IdentifierNode : ISyntaxTree
    {
        public IdentifierNode(string identifier, string arraySizeSpecifier, string initializer, string type)
        {
            Name = identifier;
            Nodes = new List<ISyntaxTree>();
            if (!string.IsNullOrEmpty(arraySizeSpecifier))
            {
                Nodes.Add(new ArraySyntaxNode(arraySizeSpecifier));
            }

            if (!string.IsNullOrEmpty(initializer))
            {
                Nodes.Add(new InitializerNode(initializer, type));
            }
            else if (!string.IsNullOrEmpty(type))
            {
                Nodes.Add(new ReferenceNode(type));
            }
        }

        public bool IsArray { get { return Nodes.OfType<ArraySyntaxNode>().Any(); } }

        public string Name { get; private set; }

        public IEnumerable<IAttribute> Attributes { get; private set; }

        public IList<ISyntaxTree> Nodes { get; private set; }
    }

    public class ArraySyntaxNode : ISyntaxTree
    {
        public ArraySyntaxNode(string specifier)
        {
            Name = specifier;
        }

        public string Name { get; private set; }

        public IEnumerable<IAttribute> Attributes { get; private set; }

        public IList<ISyntaxTree> Nodes { get; private set; }
    }

    public class ReferenceNode : ISyntaxTree
    {
        public ReferenceNode(string type)
        {
            Name = type;
        }

        public string Name { get; private set; }

        public IEnumerable<IAttribute> Attributes { get; private set; }

        public IList<ISyntaxTree> Nodes { get; private set; }
    }

    public class InitializerNode : ISyntaxTree
    {
        public InitializerNode(string keyword, string type)
        {
            Name = keyword;
            Nodes = new List<ISyntaxTree>
            {
                new ReferenceNode(type)
            };
        }

        public string Name { get; private set; }

        public IEnumerable<IAttribute> Attributes { get; private set; }

        public IList<ISyntaxTree> Nodes { get; private set; }
    }
}
