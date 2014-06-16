using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Keywords = VB6Extensions.Properties.ReservedKeywords;

namespace VB6Extensions.Lexer
{
    public interface IVBType
    {
        string Library { get; }
        string TypeName { get; }
        bool? IsReference { get; }
    }

    public abstract class VBTypeBase : IVBType
    {
        public static readonly string DefaultLibrary = "VBA";

        protected VBTypeBase(string name, bool? isReference)
            : this(DefaultLibrary, name, isReference)
        {
        }

        protected VBTypeBase(string library, string name, bool? isReference)
        {
            _library = library;
            _name = name;
            _isReference = isReference;
        }

        private readonly string _library;
        public virtual string Library { get { return _library; } }

        private readonly string _name;
        public virtual string TypeName { get { return _name; } }

        private readonly bool? _isReference;
        public virtual bool? IsReference { get { return _isReference; } }

        public override string ToString()
        {
            return string.Format("{0}.{1}", _library, _name);
        }
    }

    public class ObjectType : VBTypeBase
    {
        public ObjectType(string name)
            : base(name, true)
        {
        }
    }

    public class UserDefinedType : VBTypeBase
    {
        private readonly string _name;
        private readonly IEnumerable<IVBType> _members;

        public UserDefinedType(string library, string name, IEnumerable<IVBType> members)
            :base(library, Keywords.Type, false)
        {
            _name = name;
            _members = members;
        }

        public override string TypeName { get { return _name; } }
    }

    public class Array : VBTypeBase
    {
        private readonly IDictionary<int, int> _dimensions;
        private readonly bool _isTypeSpecified;

        public Array(IDictionary<int, int> dimensions)
            :this(new Variant(), dimensions)
        {
            _isTypeSpecified = false;
        }

        public Array(IVBType type, IDictionary<int, int> dimensions)
            :base(type.Library, type.TypeName, type.IsReference)
        {
            _isTypeSpecified = true;
            _dimensions = dimensions;
        }

        public override string ToString()
        {
            return base.ToString() + "()";
        }
    }

    public class Boolean : VBTypeBase
    {
        public Boolean() : base(Keywords.Boolean, false) { }
    }

    public class Byte : VBTypeBase
    {
        public Byte() : base(Keywords.Byte, false) { }
    }

    public class Currency : VBTypeBase
    {
        public Currency() : base(Keywords.Currency, false) { }
    }

    public class Date : VBTypeBase
    {
        public Date() : base(Keywords.Date, false) { }
    }

    public class Double : VBTypeBase
    {
        public Double() : base(Keywords.Double, false) { }
    }

    public class Integer : VBTypeBase
    {
        public Integer() : base(Keywords.Integer, false) { }
    }

    public class Long : VBTypeBase
    {
        public Long() : base(Keywords.Long, false) { }
    }

    public class Single : VBTypeBase
    {
        public Single() : base(Keywords.Single, false) { }
    }

    public class String : VBTypeBase
    {
        public String() : base(Keywords.String, false) { }
    }

    public class Variant : VBTypeBase
    {
        public Variant() : base(Keywords.Variant, null) { }
    }
}
