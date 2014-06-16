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

    public class Comment : Token
    {
        public static readonly string CommentMarker = "'";

        public Comment(string instruction)
            : base(CommentMarker, instruction)
        {
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

            token = new Comment(instruction);
            return true;
        }
    }

    public class Label : Token
    {
        public static readonly string LabelMarker = ":";

        public Label(string instruction)
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

            token = new Label(instruction);
            return true;
        }
    }

    public class Declaration : Token
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

        public Declaration(string keyword, string instruction)
            : base(keyword, instruction)
        { }

        public override bool TryParse(string instruction, out IToken token)
        {
            var noIndent = instruction.TrimStart();
            var keyword = _keywords.FirstOrDefault(k => noIndent.StartsWith(k + " "));
            if (keyword == null)
            {
                token = null;
                return false;
            }

            token = new Declaration(keyword, instruction);
            return true;
        }
    }

    public class Instruction : Token
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

        public Instruction(string keyword, string instruction)
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

            token = new Instruction(keyword, instruction);
            return true;
        }
    }

    public class Statement : Token
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

        public Statement(string keyword, string instruction)
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

            token = new Statement(keyword, instruction);
            return true;
        }
    }
}
