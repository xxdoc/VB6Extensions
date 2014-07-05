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

    public class CodeFileParser
    {
        public ISyntaxTreeNode Parse (string fileName)
        {
            var content = File.ReadAllLines(fileName);
            var currentLine = 0;

            var header = ParseFileHeader(fileName, content, ref currentLine);
            var declarations = ParseDeclarations(content, ref currentLine);
            var members = ParseMembers(content, ref currentLine);

            var module = new ModuleNode(header, declarations, members);
            return module;
        }

        private ISyntaxTreeNode ParseFileHeader(string fileName, string[] content, ref int currentLine)
        {
            ISyntaxTreeNode result;

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

        private IEnumerable<ISyntaxTreeNode> ParseDeclarations(string[] content, ref int currentLine)
        {
            var result = new List<ISyntaxTreeNode>();

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
                    ISyntaxTreeNode node = null;
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

                    if (node != null && node.NodeName == ReservedKeywords.Type)
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
                    else if (node != null && node.NodeName == ReservedKeywords.Enum)
                    {
                        int currentValue = 0;
                        while (content[currentLine].Trim() != string.Format("{0} {1}", ReservedKeywords.End, ReservedKeywords.Enum))
                        {
                            var member = Regex.Match(content[currentLine].Trim(), @"(?<identifier>\w+)(\s\=\s(?<value>.*))?$");
                            if (member.Success)
                            {
                                var value = member.Groups["value"].Value;
                                if (string.IsNullOrEmpty(value))
                                {
                                    if (!int.TryParse(value, out currentValue))
                                    {
                                        currentValue++;
                                    }
                                }
                                node.Nodes.Add(new EnumMemberNode(member.Groups["identifier"].Value, currentValue));
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

        private IEnumerable<ISyntaxTreeNode> ParseMembers(string[] content, ref int currentLine)
        {
            //todo: refactor / extract methods/classes, and recurse

            var result = new List<ISyntaxTreeNode>();

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

    public class CodeBlockNode : SyntaxTreeNode
    {
        public CodeBlockNode(string name)
            :base(SyntaxTreeNodeType.CodeBlockNode, name)
        {
        }
    }

    public class AttributeNode : SyntaxTreeNode, IAttribute
    {
        public AttributeNode(string name, string value)
            : this(name, value, null)
        { }

        public AttributeNode(string name, string value, MemberReferenceNode member)
            :base(SyntaxTreeNodeType.AttributeNode, name)
        {
            if (member != null)
            {
                Nodes.Add(member);
            }

            Value = value.StartsWith("\"")
                ? value.Substring(1, value.Length - 2)
                : value;
        }

        public string Value { get; set; }

        public static ISyntaxTreeNode Parse(string instruction)
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

        string IAttribute.Name
        {
            get { return NodeName; }
        }

        string IAttribute.Value
        {
            get
            {
                return Value;
            }
            set
            {
                Value = value;
            }
        }
    }

    public class InterfaceNode : SyntaxTreeNode
    {
        public InterfaceNode(TypeReferenceNode reference)
            :base(SyntaxTreeNodeType.InterfaceNode, reference.NodeName)
        {
            Nodes.Add(reference);
        }
    }

    public class PropertyNode : SyntaxTreeNode
    {
        public PropertyNode(string modifier, string name, string keyword, string parameters, string type)
            :base(SyntaxTreeNodeType.PropertyNode, name)
        {
            Modifier = string.IsNullOrEmpty(modifier)
                ? (AccessModifier?)null
                : (AccessModifier)Enum.Parse(typeof(AccessModifier), modifier);

            Accessor = keyword;

            Nodes.Add(new IdentifierNode(name, type));
            foreach (var node in ParameterNode.Parse(parameters))
            {
                Nodes.Add(node);
            }
        }

        public AccessModifier? Modifier { get; private set; }
        public string Accessor { get; private set; }
    }

    public class MethodNode : SyntaxTreeNode
    {
        public MethodNode(string modifier, string name, string keyword, string parameters, string type)
            :base(SyntaxTreeNodeType.MethodNode, name)
        {
            Modifier = string.IsNullOrEmpty(modifier)
                ? (AccessModifier?)null
                : (AccessModifier)Enum.Parse(typeof(AccessModifier), modifier);

            Accessor = keyword;

            Nodes.Add(new IdentifierNode(name, type));
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

    }

    public enum ParameterType
    {
        Default,
        ByRef,
        ByVal
    }

    public class ParameterNode: SyntaxTreeNode
    {
        private static readonly string syntax = @"((?<optional>Optional)\s)?((?<by>ByRef|ByVal)\s)?((?<paramarray>ParamArray)\s)?((?<identifier>[a-zA-Z][_a-zA-Z0-9]*(\(\))?)\s?)(As\s(?<type>[a-zA-Z][_a-zA-Z0-9]*\.?[a-zA-Z][_a-zA-Z0-9]*)?)?(\s\=\s(?<default>[a-zA-Z][_a-zA-Z0-9]*))?";
        private static readonly Regex regex = new Regex(syntax);

        public ParameterNode(IdentifierNode identifier, ParameterType passedBy, bool isParamArray, bool isOptional, string defaultValue)
            :base(SyntaxTreeNodeType.ParameterNode, identifier.NodeName)
        {
            Nodes.Add(identifier);

            IsOptional = isOptional;
            DefaultValue = defaultValue;
            IsParamArray = isParamArray;
            PassedBy = passedBy;
        }

        public ParameterType PassedBy { get; private set; }

        public bool IsOptional { get; private set; }
        public string DefaultValue { get; private set; }

        public bool IsParamArray { get; private set; }

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

    public class ModuleNode : SyntaxTreeNode
    {
        public ModuleNode(ISyntaxTreeNode header, IEnumerable<ISyntaxTreeNode> declarations, IEnumerable<ISyntaxTreeNode> members)
            :base(SyntaxTreeNodeType.ModuleNode, header.Nodes.OfType<AttributeNode>().FirstOrDefault(node => node.NodeName == "VB_Name").Value)
        {
            var nodes = header.Nodes.Concat(declarations).Concat(members);
            foreach (var node in nodes)
            {
                Nodes.Add(node);
            }
        }
    }

    public class DeclarationNode : SyntaxTreeNode
    {
        public DeclarationNode(string modifier, string keyword, string identifier, string arraySizeSpecifier, string initializer, string type)
            :base(SyntaxTreeNodeType.DeclarationNode, keyword)
        {
            Modifier = string.IsNullOrEmpty(modifier) 
                ? (AccessModifier?)null 
                : (AccessModifier)Enum.Parse(typeof(AccessModifier), modifier);

            Nodes.Add(new IdentifierNode(identifier, arraySizeSpecifier, initializer, type));
        }

        public AccessModifier? Modifier { get; private set; }
    }

    public class EnumMemberNode : IdentifierNode
    {
        public EnumMemberNode(string identifier)
            : this(identifier, null)
        { }

        public EnumMemberNode(string identifier, int? value)
            : base(identifier, null)
        {
            Value = value;
        }

        public int? Value { get; private set; }
    }

    public class IdentifierNode : SyntaxTreeNode
    {
        public IdentifierNode(string identifier, string type)
            : this(identifier, null, null, type)
        {
        }

        public IdentifierNode(string identifier, string arraySizeSpecifier, string initializer, string type)
            :base(SyntaxTreeNodeType.IdentifierNode, identifier)
        {
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
    }

    public class ArraySyntaxNode : SyntaxTreeNode
    {
        public ArraySyntaxNode(string specifier)
            :base(SyntaxTreeNodeType.ArraySyntaxNode, specifier)
        {
        }
    }

    public class TypeReferenceNode : SyntaxTreeNode
    {
        public TypeReferenceNode(string type)
            :base(SyntaxTreeNodeType.TypeReferenceNode, type)
        {
        }
    }

    public class MemberReferenceNode : SyntaxTreeNode
    {
        public MemberReferenceNode(string member)
            :base(SyntaxTreeNodeType.MemberReferenceNode, member)
        {
        }
    }

    public class CommentNode : SyntaxTreeNode
    {
        public CommentNode(string comment)
            :base(SyntaxTreeNodeType.CommentNode, comment)
        {
        }

        public static bool TryParse(string instruction, out CommentNode node)
        {
            node = null;
            int? commentStart = null;

            // comments parsing is context-sensitive, won't work with a regex.
            var isInsideQuotes = false;
            for (var i = 0; i < instruction.Length; i++)
            {
                if (instruction[i] == '"')
                {
                    isInsideQuotes = !isInsideQuotes;
                }

                if (!isInsideQuotes && instruction[i] == '\'')
                {
                    commentStart = i;
                    break;
                }
            }

            if (commentStart.HasValue)
            {
                node = new CommentNode(instruction.Substring(commentStart.Value));
            }

            return true;
        }
    }

    public class InitializerNode : SyntaxTreeNode
    {
        public InitializerNode(string keyword, TypeReferenceNode typeNode)
            :base(SyntaxTreeNodeType.InitializerSyntaxNode, keyword)
        {
        }

        private static string _syntax = @"\s\=\sNew\s(?<type>[a-zA-Z][a-zA-Z0-9_]*)";
        private static Regex _regex = new Regex(_syntax);

        public static bool TryParse(string instruction, out InitializerNode node)
        {
            node = null;

            var match = _regex.Match(instruction);
            if (!match.Success)
            {
                return false;
            }

            var refNode = new TypeReferenceNode(match.Groups["type"].Value);
            node = new InitializerNode(ReservedKeywords.New, refNode);

            CommentNode comment;
            if (CommentNode.TryParse(instruction, out comment))
            {
                node.Nodes.Add(comment);
            }

            return true;
        }
    }
}
