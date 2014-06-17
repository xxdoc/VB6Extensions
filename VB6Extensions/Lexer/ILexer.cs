using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VB6Extensions.Properties;

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

        private static readonly AttributeToken _attributeLexer = new AttributeToken(string.Empty, string.Empty);
        private static readonly CommentLineToken _commentLexer = new CommentLineToken(string.Empty);
        private static readonly LabelToken _labelLexer = new LabelToken(string.Empty);
        private static readonly DeclarationToken _declarationLexer = new DeclarationToken(string.Empty, string.Empty);
        private static readonly InstructionToken _instructionLexer = new InstructionToken(string.Empty, string.Empty);
        private static readonly StatementToken _statementLexer = new StatementToken(string.Empty, string.Empty);
        private static readonly ExpressionToken _expressionLexer = new ExpressionToken(string.Empty);
        private static readonly ProcedureCallToken _callLexer = new ProcedureCallToken(string.Empty);

        public IEnumerable<IToken> Tokenize(string[] content)
        {
            var builder = new StringBuilder();
            var lineCount = 0;
            foreach (var line in content.Select(code => code.Trim()))
            {
                lineCount++;
                if (string.IsNullOrEmpty(line) || lineCount <= 13) // todo: find a better way to tokenize or skip file header
                {
                    continue;
                }

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
                    && !_attributeLexer.TryParse(text, out token)
                    && !_labelLexer.TryParse(text, out token)
                    && !_declarationLexer.TryParse(text, out token)
                    && !_instructionLexer.TryParse(text, out token)
                    && !_statementLexer.TryParse(text, out token)
                    && !_expressionLexer.TryParse(text, out token)
                    && !_callLexer.TryParse(text, out token))
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
