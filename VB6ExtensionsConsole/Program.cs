using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VB6Extensions.Parser;

namespace VB6ExtensionsConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var tokenizer = new CodeFileParser();
            var tree = tokenizer.Parse(@"VB6\SqlResult.cls");

            Console.WriteLine(tree.Name);
            foreach (var node in tree.Nodes)
            {
                RecursivePrint(node, 1);
            }

            Console.ReadLine();
        }

        static void RecursivePrint(ISyntaxTree node, int depth)
        {
            if (node is DeclarationNode && ((DeclarationNode)node).Modifier.HasValue)
            {
                Console.WriteLine(new string('\t', depth) + "{0} {1}", ((DeclarationNode)node).Modifier, node.Name);
            }
            else if (node is PropertyNode)
	        {
                Console.WriteLine(new string('\t', depth) + "{0}: {1} ({2})", node.GetType().Name, node.Name, ((PropertyNode)node).Accessor);
            }
            else
            {
                Console.WriteLine(new string('\t', depth) + "{0}: {1}", node.GetType().Name, node.Name);
            }

            foreach (var item in node.Nodes)
            {
                Console.WriteLine(new string('\t', depth) + "{0}: {1}", item.GetType().Name, item.Name);

                if (item.Nodes != null && item.Nodes.Any())
                {
                    RecursivePrint(item, depth + 1);
                }

                Console.WriteLine();
            }
        }
    }
}
