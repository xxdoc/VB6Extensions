using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VB6Extensions.Lexer
{
    public interface IIdentifier
    {
        IVBType Type { get; set; }
        string Name { get; set; }
    }

    public interface IMember
    {
        AccessModifier Accessibility { get; set; }
        IIdentifier Identifier { get; set; }
    }
}
