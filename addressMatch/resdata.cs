using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace convertPointToPoint
{
    class resdata
    {
        string _name;
        string _addr;
        string _x;
        string _y;
        public string name
        {
            set { _name = value; }
            get { return _name; }
        }
        public string addr
        {
            set { _addr = value; }
            get { return _addr; }
        }
        public string x
        {
            set { _x = value; }
            get { return _x; }
        }
        public string y
        {
            set { _y = value; }
            get { return _y; }
        }
    }
}
