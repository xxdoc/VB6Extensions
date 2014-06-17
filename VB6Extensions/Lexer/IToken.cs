using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Keywords = VB6Extensions.Properties.ReservedKeywords;

namespace VB6Extensions.Lexer
{
    public interface IToken
    {
        string Instruction { get; }
        string Keyword { get; }
        bool TryParse(string instruction, out IToken token);
    }

    public abstract class Token : IToken
    {
        protected Token(string keyword, string instruction)
        {
            _keyword = keyword;
            _instruction = instruction;
        }

        private readonly string _instruction;
        public virtual string Instruction { get { return _instruction; } }

        private readonly string _keyword;
        public virtual string Keyword { get { return _keyword; } }

        public abstract bool TryParse(string instruction, out IToken token);
    }

    public class ProcedureCallToken : Token
    {
        public ProcedureCallToken(string instruction)
            : base(instruction.Trim().Split(' ')[0], instruction)
        { }

        public override bool TryParse(string instruction, out IToken token)
        {
            token = new ProcedureCallToken(instruction);
            return true;
        }
    }

    public class ExpressionToken : Token
    {
        public ExpressionToken(string instruction)
            : base("=", instruction)
        { }

        public override bool TryParse(string instruction, out IToken token)
        {
            if (!instruction.Contains('=') || instruction.StartsWith("'"))
            {
                token = null;
                return false;
            }

            token = new ExpressionToken(instruction);
            return true;
        }
    }

    public class CommentLineToken : Token
    {
        public static readonly string CommentMarker = "'";

        public CommentLineToken(string instruction)
            : base(CommentMarker, instruction)
        {
            _instruction = instruction;
        }

        private string _instruction;
        public override string Instruction
        {
            get { return _instruction; }
        }

        public void SetContent(string comment)
        {
            if (!comment.TrimStart().StartsWith(CommentMarker))
                throw new ArgumentException("Comment must start with comment marker.");

            _instruction = comment;
        }

        public override bool TryParse(string instruction, out IToken token)
        {
            var noIndent = instruction.TrimStart();
            if (!noIndent.StartsWith(CommentMarker))
            {
                token = null;
                return false;
            }

            token = new CommentLineToken(instruction);
            return true;
        }
    }

    public class LabelToken : Token
    {
        public static readonly string LabelMarker = ":";

        public LabelToken(string instruction)
            : base(LabelMarker, instruction)
        {
        }

        public bool IsCaseLabel 
        { 
            get 
            { 
                return _instruction.TrimStart().StartsWith(Keywords.Case + " "); 
            } 
        }

        private string _instruction;
        public override string Instruction
        {
            get 
            { 
                return _instruction; 
            }
        }

        public void SetContent(string label)
        {
            if (!label.TrimEnd().EndsWith(LabelMarker))
                throw new ArgumentException("Label must end with label marker.");

            _instruction = label;
        }

        public override bool TryParse(string instruction, out IToken token)
        {
            var trimmed = instruction.TrimEnd();
            if (!trimmed.EndsWith(LabelMarker))
            {
                token = null;
                return false;
            }

            token = new LabelToken(instruction);
            return true;
        }
    }

    public class AttributeToken : Token
    {
        public AttributeToken(string keyword, string instruction)
            : base(Keywords.Attribute, instruction)
        { }

        public override bool TryParse(string instruction, out IToken token)
        {
            var noIndent = instruction.TrimStart();
            if (!noIndent.StartsWith(Keyword))
            {
                token = null;
                return false;
            }

            token = new AttributeToken(Keyword, instruction);
            return true;
        }
    }

    public class DeclarationToken : Token
    {
        private readonly string[] _keywords = new[] 
                                {
                                    Keywords.Const,
                                    Keywords.Dim, 
                                    Keywords.Static,
                                    Keywords.Public, 
                                    Keywords.Private, 
                                    Keywords.Friend, 
                                    Keywords.Global,
                                    Keywords.Declare,
                                    Keywords.Type,
                                    Keywords.Enum,
                                    Keywords.Property,
                                    Keywords.Function,
                                    Keywords.Sub
                                };

        public DeclarationToken(string keyword, string instruction)
            : base(keyword, instruction)
        { }

        public override bool TryParse(string instruction, out IToken token)
        {
            var noIndent = instruction.TrimStart();
            var keyword = _keywords.FirstOrDefault(k => noIndent.StartsWith(k + " "));
            if (keyword == null && !noIndent.Contains(Keywords.As))
            {
                token = null;
                return false;
            }
            else if (noIndent.Contains(Keywords.As) && !noIndent.Contains('#') && !noIndent.Contains('='))
            {
                // member declaration. "Public" keyword used only for semantics.
                token = new DeclarationToken(Keywords.Public, instruction);
                return true;
            }

            token = new DeclarationToken(keyword, instruction);
            return true;
        }
    }

    public class InstructionToken : Token
    {
        private readonly string[] _keywords = new[] 
                                {
                                    Keywords.Set,
                                    Keywords.Let,
                                    Keywords.Call,
                                    Keywords.Beep,
                                    Keywords.Open,
                                    Keywords.Option,
                                    Keywords.On,
                                    Keywords.Close,
                                    Keywords.ChDir,
                                    Keywords.ChDrive,
                                    Keywords.Debug,
                                    Keywords.DoEvents,
                                    Keywords.Exit,
                                    Keywords.End,
                                    Keywords.GoTo,
                                    Keywords.GoSub,
                                    Keywords.Input,
                                    Keywords.Kill,
                                    Keywords.MkDir,
                                    Keywords.MsgBox,
                                    Keywords.Output,
                                    Keywords.Option,
                                    Keywords.Print,
                                    Keywords.Put,
                                    Keywords.Randomize,
                                    Keywords.Read,
                                    Keywords.ReDim,
                                    Keywords.Resume,
                                    Keywords.Return,
                                    Keywords.RmDir,
                                    Keywords.Shell,
                                    Keywords.Stop,
                                    Keywords.Write
                                };

        public InstructionToken(string keyword, string instruction)
            : base(keyword, instruction)
        { }

        public override bool TryParse(string instruction, out IToken token)
        {
            var noIndent = instruction.TrimStart();
            var keyword = _keywords.FirstOrDefault(k => noIndent.StartsWith(k));
            if (keyword == null)
            {
                token = null;
                return false;
            }

            token = new InstructionToken(keyword, instruction);
            return true;
        }
    }

    public class StatementToken : Token
    {
        private readonly string[] _keywords = new[] 
                                {
                                    Keywords.Do,
                                    Keywords.Else,
                                    Keywords.ElseIf,
                                    Keywords.For,
                                    Keywords.If + " ",
                                    Keywords.Loop,
                                    Keywords.Next,
                                    Keywords.Select + " ",
                                    Keywords.Wend,
                                    Keywords.While,
                                    Keywords.With + " "
                                };

        public StatementToken(string keyword, string instruction)
            : base(keyword, instruction)
        { }

        public override bool TryParse(string instruction, out IToken token)
        {
            var noIndent = instruction.TrimStart();
            var keyword = _keywords.FirstOrDefault(k => noIndent.StartsWith(k));
            if (keyword == null)
            {
                token = null;
                return false;
            }

            token = new StatementToken(keyword, instruction);
            return true;
        }
    }
}
