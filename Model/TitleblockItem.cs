using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPI_Quyen.Model
{
    public class TitleblockItem
    {
        public string Name
        {
            set;
            get;
        }

        public Element Value
        {
            set;
            get;
        }

        public TitleblockItem() { }

        public override string ToString()
        {
            return Name;
        }
    }
}
