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
    public class TreeNodeViewModel : ISyntaxTreeNode
    {
        private readonly ISyntaxTreeNode _node;
        private string _description;

        public TreeNodeViewModel(ISyntaxTreeNode node)
        {
            var moduleNode = node as CodeModule;

            _node = node;
            Icon = SetIcon();

            AttributeVisibility = Visibility.Collapsed;

            if (node is AttributeNode)
            {
                AttributeValue = (node as AttributeNode).Value;
                AttributeVisibility = Visibility.Visible;
            }
        }

        public Visibility AttributeVisibility { get; private set; }
        public string AttributeValue { get; private set; }

        public string Icon { get; private set; }

        private string SetIcon()
        {
            if (_node is AttributeNode)
            {
                return "icons/assembly.png";
            }
            else if (_node is MethodNode)
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
            else if (_node is EnumMemberNode)
            {
                return "icons/enum_member.png";
            }
            else if (_node is DeclarationNode)
            {
                var node = _node as DeclarationNode;
                if (node.Modifier == null)
                {
                    if (node.NodeName == "Private")
                        return "icons/field_private.png";
                    if (node.NodeName == "Friend")
                        return "icons/field_friend.png";
                    return "icons/field.png";
                }
                else
                {
                    switch (node.Modifier)
                    {
                        case AccessModifier.Private:
                            return node.NodeName == "Enum" ? "icons/enum_private.png"
                                                       : node.NodeName == "Type" ? "icons/struct_private.png"
                                                                             : "icons/field_private.png";
                            break;
                        case AccessModifier.Friend:
                            return node.NodeName == "Enum" ? "icons/enum_friend.png"
                                                       : node.NodeName == "Type" ? "icons/struct_friend.png"
                                                                             : "icons/field_friend.png";
                            break;
                        default:
                            return node.NodeName == "Enum" ? "icons/enum.png"
                                                       : node.NodeName == "Type" ? "icons/struct.png" 
                                                                             : "icons/field.png";
                            break;
                    }
                }
            }
            else if (_node is InterfaceNode)
            {
                return "icons/interface.png";
            }
            else if (_node is TypeReferenceNode)
            {
                var node = _node as TypeReferenceNode;
                if (new[]{"String", "Byte", "Boolean", "Integer", "Date", "Long", "Double", "Single", "Variant", "Currency"}.Contains(node.NodeName))
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

        public string NodeName
        {
            get { return _node.NodeName + _description; }
            set { _node.NodeName = value; }
        }

        public IList<ISyntaxTreeNode> Nodes
        {
            get 
            {
                var result = new List<ISyntaxTreeNode>();
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

        public SyntaxTreeNodeType NodeType
        {
            get { return _node.NodeType; }
        }
    }
}
