using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VB6Extensions.Lexer
{
    public interface ILexer
    {
        IEnumerable<IToken> Tokenize(string[] content);
    }

    public class Tokenizer : ILexer
    {
        public static readonly string LineContinuation = "_";
        public static readonly string InstructionSeparator = ":";

        private static readonly Comment _commentLexer = new Comment(string.Empty);
        private static readonly Label _labelLexer = new Label(string.Empty);
        private static readonly Declaration _declarationLexer = new Declaration(string.Empty, string.Empty);
        private static readonly Instruction _instructionLexer = new Instruction(string.Empty, string.Empty);
        private static readonly Statement _statementLexer = new Statement(string.Empty, string.Empty);

        public IEnumerable<IToken> Tokenize(string[] content)
        {
            var builder = new StringBuilder();
            foreach (var line in content.Select(code => code.Trim()))
            {
                if (string.IsNullOrEmpty(line))
                    continue;

                var allButLastCharacter = line.Substring(0, line.Length - 1);
                if (allButLastCharacter.Contains(InstructionSeparator)) // todo: escape strings
                {
                    var instructions = allButLastCharacter.Split(InstructionSeparator[0]);
                    foreach (var instruction in instructions)
                    {
                        builder.AppendLine(instruction); // todo: test when last instruction has line continuation character.
                    }
                }
                else if (line.EndsWith(LineContinuation))
                {
                    builder.Append(allButLastCharacter);
                    continue;
                }
                else
                {
                    builder.AppendLine(line);
                }
                
                var text = builder.ToString();
                IToken token;
                if (!_commentLexer.TryParse(text, out token)
                    && !_labelLexer.TryParse(text, out token)
                    && !_declarationLexer.TryParse(text, out token)
                    && !_instructionLexer.TryParse(text, out token)
                    && !_statementLexer.TryParse(text, out token))
                {
                    token = null;
                }

                if (token != null)
                {
                    yield return token;
                }

                builder.Clear();
            }
        }
    }
}
