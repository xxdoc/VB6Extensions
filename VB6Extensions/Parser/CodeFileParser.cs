﻿using System;
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

                var regex = new Regex(@"\=\s\-?(?<IntValue>\d+)\s");
                result = new ClassModule(fileName)
                                {
                                    DataBindingBehavior = int.Parse(regex.Match(dataBindingBehavior).Groups["IntValue"].Value),
                                    DataSourceBehavior = int.Parse(regex.Match(dataSourceBehavior).Groups["IntValue"].Value),
                                    MTSTransactionMode = int.Parse(regex.Match(mtsTransactionMode).Groups["IntValue"].Value),
                                    MultiUse = int.Parse(regex.Match(multiUse).Groups["IntValue"].Value),
                                    Persistable = int.Parse(regex.Match(persistable).Groups["IntValue"].Value)
                                };

                result.Nodes.Add(AttributeNode.Parse(content[8].Trim()));
                result.Nodes.Add(AttributeNode.Parse(content[9].Trim()));
                result.Nodes.Add(AttributeNode.Parse(content[10].Trim()));
                result.Nodes.Add(AttributeNode.Parse(content[11].Trim()));
                result.Nodes.Add(AttributeNode.Parse(content[12].Trim()));

                currentLine = 13;
            }
            else
            {
                result = new CodeModule(fileName);
                result.Nodes.Add(AttributeNode.Parse(content[0].Trim()));

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
                if (content[currentLine].Trim().StartsWith("'"))
                {
                    // comment node
                    currentLine++;
                    continue;
                }

                var line = content[currentLine];
                isDeclarationSection =  !(
                                               line.Contains(ReservedKeywords.Property)
                                            || line.Contains(ReservedKeywords.Sub)
                                            || line.Contains(ReservedKeywords.Function)
                                         );
                if (isDeclarationSection)
                {
                    currentLine++;
                    var match = regex.Match(line);
                    ISyntaxTree node = null;
                    if (match.Success && !line.StartsWith(ReservedKeywords.Implements))
                    {
                        if (match.Groups["keyword"].Captures.Count == 2)
                        {
                            node = new DeclarationNode(match.Groups["keyword"].Captures[0].Value, 
                                                           match.Groups["keyword"].Captures[1].Value,
                                                           match.Groups["identifier"].Value,
                                                           match.Groups["arraySize"].Value,
                                                           match.Groups["initializer"].Value,
                                                           match.Groups["type"].Value);
                        }
                        else
                        {
                            node = new DeclarationNode(null, match.Groups["keyword"].Value,
                                                                 match.Groups["identifier"].Value,
                                                                 match.Groups["arraySize"].Value,
                                                                 match.Groups["initializer"].Value,
                                                                 match.Groups["type"].Value);
                        }
                    }
                    else if(line.StartsWith(ReservedKeywords.Implements))
                    {
                        var implements = Regex.Match(line, ReservedKeywords.Implements + @"\s(?<type>[a-zA-Z][_a-zA-Z0-9]*)$");
                        if (implements.Success)
                        {
                            var reference = new TypeReferenceNode(implements.Groups["type"].Value);
                            node = new InterfaceNode(reference);
                        }
                    }

                    if (node != null && node.Name == ReservedKeywords.Type)
                    {
                        while(content[currentLine].Trim() != string.Format("{0} {1}", ReservedKeywords.End, ReservedKeywords.Type))
                        {
                            var member = Regex.Match(content[currentLine].Trim(), @"(?<identifier>\w+)(?<arraySize>\(.*\))?(\s+As\s+?(((?<initializer>New)\s+)?)(?<type>\w+(\.\w+)?))$");
                            if (member.Success)
                            {
                                node.Nodes.Add(new DeclarationNode(null, member.Groups["identifier"].Value, member.Groups["identifier"].Value, member.Groups["arraySize"].Value, member.Groups["initializer"].Value, member.Groups["type"].Value));
                            }
                            currentLine++;
                        }
                    }
                    else if (node != null && node.Name == ReservedKeywords.Enum)
                    {
                        while (content[currentLine].Trim() != string.Format("{0} {1}", ReservedKeywords.End, ReservedKeywords.Enum))
                        {
                            var member = Regex.Match(content[currentLine].Trim(), @"(?<identifier>\w+)(\s\=\s(?<value>.*))?$");
                            if (member.Success)
                            {
                                node.Nodes.Add(new EnumMemberNode(member.Groups["identifier"].Value));
                            }
                            currentLine++;
                        }
                    }

                    if (node != null)
                    {
                        result.Add(node);
                    }
                }
            }

            return result;
        }

        private IEnumerable<ISyntaxTree> ParseMembers(string[] content, ref int currentLine)
        {
            //todo: refactor / extract methods/classes, and recurse

            var result = new List<ISyntaxTree>();

            var pattern = @"((?<keyword>Public|Private|Friend)\s)?(?<keyword>Property|Function|Sub)\s+((?<keyword>Get|Let|Set)\s+)?(?<name>[a-zA-Z][a-zA-Z0-9_]*)(\((?<parameters>.*)\))?(\s+As\s+(?<type>.*))?$";
            var regex = new Regex(pattern);

            while (currentLine < content.Length)
            {
                if (content[currentLine].Trim().StartsWith("'"))
                {
                    // comment node
                    currentLine++;
                    continue;
                }

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
                                                        keyword, match.Groups["parameters"].Value,
                                                        match.Groups["type"].Value);
                            var body = new CodeBlockNode(match.Groups["name"].Value);
                            currentLine++;
                            while (content[currentLine].Trim() != string.Format("{0} {1}", ReservedKeywords.End, match.Groups["keyword"].Captures[1].Value))
                            {
                                var attribute = AttributeNode.Parse(content[currentLine]);
                                if (attribute != null)
                                {
                                    node.Nodes.Add(attribute);
                                }
                                else
                                {
                                    var trimmed = content[currentLine].Trim();
                                    if (Regex.IsMatch(trimmed, @"If\s.*Then$"))
                                    {
                                        var ifBlock = new CodeBlockNode(content[currentLine]);
                                        currentLine++;
                                        while (content[currentLine].Trim() != string.Format("{0} {1}", ReservedKeywords.End, ReservedKeywords.If))
                                        {
                                            ifBlock.Nodes.Add(new CodeBlockNode(content[currentLine]));
                                            currentLine++;
                                        }
                                        currentLine++;
                                        body.Nodes.Add(ifBlock);
                                    }
                                    else if (Regex.IsMatch(trimmed, @"For\s.*$"))
                                    {
                                        var forLoopBlock = new CodeBlockNode(content[currentLine]);
                                        currentLine++;
                                        while (content[currentLine].Trim() != ReservedKeywords.Next)
                                        {
                                            forLoopBlock.Nodes.Add(new CodeBlockNode(content[currentLine]));
                                            currentLine++;
                                        }
                                        currentLine++;
                                        body.Nodes.Add(forLoopBlock);
                                    }

                                    body.Nodes.Add(new CodeBlockNode(content[currentLine]));
                                }
                                currentLine++;
                            }

                            node.Nodes.Add(body);
                            result.Add(node);
                        }
                        else
                        {
                            var keyword = match.Groups["keyword"].Captures[1].Value;
                            var node = new MethodNode(modifier, match.Groups["name"].Value,
                                                        keyword, match.Groups["parameters"].Value,
                                                        match.Groups["type"].Value);
                            var body = new CodeBlockNode(match.Groups["name"].Value);
                            currentLine++;
                            while (content[currentLine].Trim() != string.Format("{0} {1}", ReservedKeywords.End, keyword))
                            {
                                var attribute = AttributeNode.Parse(content[currentLine]);
                                if (attribute != null)
                                {
                                    node.Nodes.Add(attribute);
                                }
                                else
                                {
                                    var trimmed = content[currentLine].Trim();
                                    if (Regex.IsMatch(trimmed, @"If\s.*Then$"))
                                    {
                                        var ifBlock = new CodeBlockNode(content[currentLine]);
                                        currentLine++;
                                        while (content[currentLine].Trim() != string.Format("{0} {1}", ReservedKeywords.End, ReservedKeywords.If))
                                        {
                                            ifBlock.Nodes.Add(new CodeBlockNode(content[currentLine]));
                                            currentLine++;
                                        }
                                        currentLine++;
                                        body.Nodes.Add(ifBlock);
                                    }
                                    else if (Regex.IsMatch(trimmed, @"For\s.*$"))
                                    {
                                        var forLoopBlock = new CodeBlockNode(content[currentLine]);
                                        currentLine++;
                                        while (content[currentLine].Trim() != ReservedKeywords.Next)
                                        {
                                            forLoopBlock.Nodes.Add(new CodeBlockNode(content[currentLine]));
                                            currentLine++;
                                        }
                                        currentLine++;
                                        body.Nodes.Add(forLoopBlock);
                                    }

                                    body.Nodes.Add(new CodeBlockNode(content[currentLine]));
                                }
                                currentLine++;
                            }

                            node.Nodes.Add(body);
                            result.Add(node);
                        }
                    }
                    else
                    {
                        if (match.Groups["keyword"].Captures[0].Value == ReservedKeywords.Property)
                        {
                            var keyword = match.Groups["keyword"].Captures[1].Value;
                            var node = new PropertyNode(null, match.Groups["name"].Value,
                                                        keyword, match.Groups["parameters"].Value,
                                                        match.Groups["type"].Value);
                            var body = new CodeBlockNode(match.Groups["name"].Value);
                            currentLine++;
                            while (content[currentLine].Trim() != string.Format("{0} {1}", ReservedKeywords.End, keyword))
                            {
                                var attribute = AttributeNode.Parse(content[currentLine]);
                                if (attribute != null)
                                {
                                    node.Nodes.Add(attribute);
                                }
                                else
                                {
                                    var trimmed = content[currentLine].Trim();
                                    if (Regex.IsMatch(trimmed, @"If\s.*Then$"))
                                    {
                                        var ifBlock = new CodeBlockNode(content[currentLine]);
                                        currentLine++;
                                        while (content[currentLine].Trim() != string.Format("{0} {1}", ReservedKeywords.End, ReservedKeywords.If))
                                        {
                                            ifBlock.Nodes.Add(new CodeBlockNode(content[currentLine]));
                                            currentLine++;
                                        }
                                        currentLine++;
                                        body.Nodes.Add(ifBlock);
                                    }
                                    else if (Regex.IsMatch(trimmed, @"For\s.*$"))
                                    {
                                        var forLoopBlock = new CodeBlockNode(content[currentLine]);
                                        currentLine++;
                                        while (content[currentLine].Trim() != ReservedKeywords.Next)
                                        {
                                            forLoopBlock.Nodes.Add(new CodeBlockNode(content[currentLine]));
                                            currentLine++;
                                        }
                                        currentLine++;
                                        body.Nodes.Add(forLoopBlock);
                                    }
                                    
                                    body.Nodes.Add(new CodeBlockNode(content[currentLine]));
                                }

                                currentLine++;
                            }

                            node.Nodes.Add(body);
                            result.Add(node);
                        }
                        else
                        {
                            var keyword = match.Groups["keyword"].Captures[1].Value;
                            var node = new MethodNode(null, match.Groups["name"].Value,
                                                        keyword, match.Groups["parameters"].Value,
                                                        match.Groups["type"].Value);
                            var body = new CodeBlockNode(match.Groups["name"].Value);
                            currentLine++;
                            while (content[currentLine].Trim() != string.Format("{0} {1}", ReservedKeywords.End, keyword))
                            {
                                var attribute = AttributeNode.Parse(content[currentLine]);
                                if (attribute != null)
                                {
                                    node.Nodes.Add(attribute);
                                }
                                else
                                {
                                    var trimmed = content[currentLine].Trim();
                                    if (Regex.IsMatch(trimmed, @"If\s.*Then$"))
                                    {
                                        var ifBlock = new CodeBlockNode(content[currentLine]);
                                        currentLine++;
                                        while (content[currentLine].Trim() != string.Format("{0} {1}", ReservedKeywords.End, ReservedKeywords.If))
                                        {
                                            ifBlock.Nodes.Add(new CodeBlockNode(content[currentLine]));
                                            currentLine++;
                                        }

                                        currentLine++;
                                        body.Nodes.Add(ifBlock);
                                    }
                                    else if (Regex.IsMatch(trimmed, @"For\s.*$"))
                                    {
                                        var forLoopBlock = new CodeBlockNode(content[currentLine]);
                                        currentLine++;
                                        while (content[currentLine].Trim() != ReservedKeywords.Next)
                                        {
                                            forLoopBlock.Nodes.Add(new CodeBlockNode(content[currentLine]));
                                            currentLine++;
                                        }

                                        currentLine++;
                                        body.Nodes.Add(forLoopBlock);
                                    }

                                    body.Nodes.Add(new CodeBlockNode(content[currentLine]));
                                }

                                currentLine++;
                            }

                            node.Nodes.Add(body);
                            result.Add(node);
                        }
                    }
                }

                currentLine++;
            }

            return result;
        }
    }

    public class CodeBlockNode : ISyntaxTree
    {
        public CodeBlockNode(string name)
        {
            Name = name;
            Nodes = new List<ISyntaxTree>();
        }

        public string Name { get; private set; }
        public IEnumerable<IAttribute> Attributes { get; private set; }
        public IList<ISyntaxTree> Nodes { get; private set; }
    }

    public class AttributeNode : ISyntaxTree, IAttribute
    {
        public AttributeNode(string name, string value)
            : this(name, value, null)
        { }

        public AttributeNode(string name, string value, MemberReferenceNode member)
        {
            Nodes = new List<ISyntaxTree>();
            if (member != null)
            {
                Nodes.Add(member);
            }

            Name = name;
            Value = value.StartsWith("\"")
                ? value.Substring(1, value.Length - 2)
                : value;
        }

        public string Name { get; private set; }
        public string Value { get; set; }

        public IEnumerable<IAttribute> Attributes { get; private set; }
        public IList<ISyntaxTree> Nodes { get; private set; }

        public static ISyntaxTree Parse(string instruction)
        {
            var syntax = @"^Attribute\s((?<Member>[a-zA-Z][a-zA-Z0-9_]*)\.)?(?<Name>VB_\w+)\s=\s(?<Value>.*)$";
            var regex = new Regex(syntax);

            if (!regex.IsMatch(instruction))
            {
                return null;
            }

            var match = regex.Match(instruction);
            var member = match.Groups["Member"].Value;
            var name = match.Groups["Name"].Value;
            var value = match.Groups["Value"].Value;

            MemberReferenceNode reference = null;
            if (!string.IsNullOrEmpty(member))
            {
                reference = new MemberReferenceNode(member);
            }

            return new AttributeNode(name, value, reference);
        }
    }

    public class InterfaceNode : ISyntaxTree
    {
        public InterfaceNode(TypeReferenceNode reference)
        {
            Name = reference.Name;
            Nodes = new List<ISyntaxTree> { reference };
        }

        public string Name { get; private set; }
        public IEnumerable<IAttribute> Attributes { get; private set; }
        public IList<ISyntaxTree> Nodes { get; private set; }
    }

    public class PropertyNode : ISyntaxTree
    {
        public PropertyNode(string modifier, string name, string keyword, string parameters, string type)
        {
            Modifier = string.IsNullOrEmpty(modifier)
                ? (AccessModifier?)null
                : (AccessModifier)Enum.Parse(typeof(AccessModifier), modifier);

            Name = name;
            Accessor = keyword;

            Nodes = new List<ISyntaxTree> { new IdentifierNode(name, type) };
            foreach (var node in ParameterNode.Parse(parameters))
            {
                Nodes.Add(node);
            }
        }

        public AccessModifier? Modifier { get; private set; }

        public string Accessor { get; private set; }
        public string Name { get; private set; }

        public ISyntaxTree Parent { get; private set; }
        public IList<ISyntaxTree> Nodes { get; private set; }
    }

    public class MethodNode : ISyntaxTree
    {
        public MethodNode(string modifier, string name, string keyword, string parameters, string type)
        {
            Modifier = string.IsNullOrEmpty(modifier)
                ? (AccessModifier?)null
                : (AccessModifier)Enum.Parse(typeof(AccessModifier), modifier);

            Name = name;
            Accessor = keyword;

            Nodes = new List<ISyntaxTree> { new IdentifierNode(name, type) };
            if (!string.IsNullOrWhiteSpace(parameters))
            {
                foreach (var node in ParameterNode.Parse(parameters))
                {
                    Nodes.Add(node);
                }
            }
        }

        public AccessModifier? Modifier { get; private set; }

        public string Accessor { get; private set; }
        public string Name { get; private set; }

        private readonly IList<IAttribute> _attributes = new List<IAttribute>();
        public IEnumerable<IAttribute> Attributes { get { return _attributes; } }
        public void AddAttribute(IAttribute attribute)
        {
            _attributes.Add(attribute);
        }

        public IList<ISyntaxTree> Nodes { get; private set; }
    }

    public enum ParameterType
    {
        Default,
        ByRef,
        ByVal
    }


    public class ParameterNode: ISyntaxTree
    {
        private static readonly string pattern = @"((?<optional>Optional)\s)?((?<by>ByRef|ByVal)\s)?((?<paramarray>ParamArray)\s)?((?<identifier>[a-zA-Z][_a-zA-Z0-9]*(\(\))?)\s?)(As\s(?<type>[a-zA-Z][_a-zA-Z0-9]*\.?[a-zA-Z][_a-zA-Z0-9]*)?)?(\s\=\s(?<default>[a-zA-Z][_a-zA-Z0-9]*))?";
        private static readonly Regex regex = new Regex(pattern);

        public ParameterNode(IdentifierNode identifier, ParameterType passedBy, bool isParamArray, bool isOptional, string defaultValue)
        {
            Name = identifier.Name;

            Nodes = new List<ISyntaxTree>
            {
                identifier
            };

            IsOptional = isOptional;
            DefaultValue = defaultValue;
            IsParamArray = isParamArray;
            PassedBy = passedBy;
        }

        public ParameterType PassedBy { get; private set; }

        public bool IsOptional { get; private set; }
        public string DefaultValue { get; private set; }

        public bool IsParamArray { get; private set; }
        public string Name { get; private set; }

        public IEnumerable<IAttribute> Attributes { get; private set; }
        public IList<ISyntaxTree> Nodes { get; private set; }

        public static IEnumerable<ParameterNode> Parse(string parameters)
        {
            var matches = regex.Matches(parameters);

            foreach (Match match in matches)
            {
                if (!match.Success)
                {
                    continue;
                }

                var optional = match.Groups["optional"].Value == ReservedKeywords.Optional;
                var defaultValue = match.Groups["default"].Value;

                var paramArray = match.Groups["paramarray"].Value == ReservedKeywords.ParamArray;

                var by = match.Groups["by"].Value;
                var passedBy = by == ReservedKeywords.ByRef ? ParameterType.ByRef
                                                            : by == ReservedKeywords.ByVal ? ParameterType.ByVal
                                                                                           : ParameterType.Default;

                var name = match.Groups["identifier"].Value;
                var type = match.Groups["type"].Value;
                var identifier = new IdentifierNode(name, paramArray ? "()" : null, null, type);

                yield return new ParameterNode(identifier, passedBy, paramArray, optional, defaultValue);
            }
        }
    }

    public class ModuleNode : ISyntaxTree
    {
        public ModuleNode(ISyntaxTree header, IEnumerable<ISyntaxTree> declarations, IEnumerable<ISyntaxTree> members)
        {
            Name = header.Nodes.OfType<AttributeNode>().First().Value;
            Nodes = new List<ISyntaxTree>(header.Nodes.Concat(declarations).Concat(members));
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

    public class EnumMemberNode : IdentifierNode
    {
        public EnumMemberNode(string identifier)
            : base(identifier, null)
        { }
    }

    public class IdentifierNode : ISyntaxTree
    {
        public IdentifierNode(string identifier, string type)
            : this(identifier, null, null, type)
        {
        }

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
                Nodes.Add(new InitializerNode(initializer, new TypeReferenceNode(type)));
            }
            else if (!string.IsNullOrEmpty(type))
            {
                Nodes.Add(new TypeReferenceNode(type));
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
            Nodes = new List<ISyntaxTree>();
        }

        public string Name { get; private set; }

        public IEnumerable<IAttribute> Attributes { get; private set; }

        public IList<ISyntaxTree> Nodes { get; private set; }
    }

    public class TypeReferenceNode : ISyntaxTree
    {
        public TypeReferenceNode(string type)
        {
            Name = type;
            Nodes = new List<ISyntaxTree>();
        }

        public string Name { get; private set; }

        public IEnumerable<IAttribute> Attributes { get; private set; }

        public IList<ISyntaxTree> Nodes { get; private set; }
    }

    public class MemberReferenceNode : ISyntaxTree
    {
        public MemberReferenceNode(string member)
        {
            Name = member;
            Nodes = new List<ISyntaxTree>();
        }

        public string Name { get; private set; }

        public IEnumerable<IAttribute> Attributes { get; private set; }

        public IList<ISyntaxTree> Nodes { get; private set; }
    }

    public class InitializerNode : ISyntaxTree
    {
        public InitializerNode(string keyword, TypeReferenceNode typeNode)
        {
            Name = keyword;
            Nodes = new List<ISyntaxTree> { typeNode };
        }

        public string Name { get; private set; }

        public IEnumerable<IAttribute> Attributes { get; private set; }

        public IList<ISyntaxTree> Nodes { get; private set; }
    }
}
