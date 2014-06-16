using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VB6Extensions.Lexer;

namespace VB6ExtensionsConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var tokenizer = new Tokenizer();
            var tokens = tokenizer.Tokenize(File.ReadAllLines(@"VB6\SqlCommand.cls"));

            foreach (var token in tokens)
            {
                Console.WriteLine(string.Format("{0}: [{1}] {2}", token.ToString(), token.Keyword, token.Instruction));
            }

            Console.ReadLine();
        }
    }
}
