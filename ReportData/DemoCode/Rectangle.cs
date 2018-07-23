using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemoCode
{
    class Rectangle
    {
        private int _Width;
        private int _Height;

        public virtual void SetWidth(int Value)
        {
            this._Width = Value;
        }

        public virtual void SetHeight(int Value)
        {
            this._Height = Value;
        }

        public int GetWidth()
        {
            return _Width;
        }

        public int GetHeight()
        {
            return _Height;
        }

    }

    class Sqaure : Rectangle
    {
        public override void SetHeight(int Value)
        {
            base.SetHeight(Value);
            base.SetWidth(Value);
        }

        public override void SetWidth(int Value)
        {
            base.SetWidth(Value);
            base.SetHeight(Value);
        }

    }
}
