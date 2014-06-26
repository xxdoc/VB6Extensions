using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VB6Extensions.Lexer
{
    public class InvalidTokenException : Exception
    {
        public InvalidTokenException(string keyword)
            :base(string.Format("Invalid token. Expected '{0}' token.", keyword))
        {
            Keyword = keyword;
        }

        public string Keyword { get; private set; }
    }
}
