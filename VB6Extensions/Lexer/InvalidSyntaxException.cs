using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VB6Extensions.Lexer
{
    public class InvalidSyntaxException : Exception
    {
        public InvalidSyntaxException(string instruction)
            :base(string.Format("Instruction '{0}' did not match expected syntax.", instruction))
        { }
    }
}
