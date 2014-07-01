using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using VB6Extensions;
using VB6Extensions.Parser;

namespace VB6ExtensionsUI
{
    public class TreeNodeViewModel : ISyntaxTree
    {
        private readonly ISyntaxTree _node;
        private string _description;

        public TreeNodeViewModel(ISyntaxTree node)
        {
            _node = node;
            Icon = SetIcon();

            DescriptionVisibility = Visibility.Collapsed;
            if (node.Attributes != null)
            {
                var descriptionAttribute = node.Attributes.FirstOrDefault(a => a.Name.EndsWith("Description"));
                if (descriptionAttribute != null && !string.IsNullOrEmpty(descriptionAttribute.Name))
                {
                    Description = Regex.Match(descriptionAttribute.Value, @"\""(?<value>.*)\""").Groups["value"].Value;
                    DescriptionVisibility = Visibility.Visible;
                }
            }
        }

        public Visibility DescriptionVisibility { get; private set; }

        public string Description { get; private set; }

        public string Icon { get; private set; }

        private string SetIcon()
        {
            if (_node is MethodNode)
            {
                var node = _node as MethodNode;
                switch (node.Modifier)
                {
                    case AccessModifier.Private:
                        return "icons/method_private.png";
                        break;
                    case AccessModifier.Friend:
                        return "icons/method_friend.png";
                        break;
                    default:
                        return "icons/method.png";
                        break;
                }
            }
            else if (_node is PropertyNode)
            {
                var node = _node as PropertyNode;
                _description = string.Format(" ({0})", node.Accessor);
                switch (node.Modifier)
                {
                    case AccessModifier.Private:
                        return "icons/property_private.png";
                        break;
                    case AccessModifier.Friend:
                        return "icons/property_friend.png";
                        break;
                    default:
                        return "icons/property.png";
                        break;
                }
            }
            else if (_node is ParameterNode)
            {
                return "icons/variable.png";
            }
            else if (_node is DeclarationNode)
            {
                var node = _node as DeclarationNode;
                if (node.Modifier == null)
                {
                    if (node.Name == "Private")
                        return "icons/field_private.png";
                    if (node.Name == "Friend")
                        return "icons/field_friend.png";
                    return "icons/field.png";
                }
                else
                {
                    switch (node.Modifier)
                    {
                        case AccessModifier.Private:
                            return "icons/field_private.png";
                            break;
                        case AccessModifier.Friend:
                            return "icons/field_friend.png";
                            break;
                        default:
                            return "icons/field.png";
                            break;
                    }
                }
            }
            else if (_node is InterfaceNode)
            {
                return "icons/interface.png";
            }
            else if (_node is ReferenceNode)
            {
                var node = _node as ReferenceNode;
                if (new[]{"String", "Byte", "Boolean", "Integer", "Date", "Long", "Double", "Single", "Variant", "Currency"}.Contains(node.Name))
                {
                    return "icons/library.png";
                }
                else
                {
                    return "icons/class_reference.png";
                }
            }
            else if(_node is CodeBlockNode)
            {
                return "icons/scope.png";
            }

            return "icons/misc_field.png";
        }

        public string Name
        {
            get { return _node.Name + _description; }
        }

        public IEnumerable<VB6Extensions.Lexer.Attributes.IAttribute> Attributes
        {
            get { return _node.Attributes; }
        }

        public IList<ISyntaxTree> Nodes
        {
            get 
            {
                var result = new List<ISyntaxTree>();
                foreach (var node in _node.Nodes)
                {
                    result.Add(new TreeNodeViewModel(node));
                }
                return result;
            }
        }

        public override string ToString()
        {
            return _node.ToString();
        }
    }
}
