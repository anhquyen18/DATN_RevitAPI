using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPI_Quyen.Util
{
    class ElementUtil
    {
        public static string ElementInfo(Element ele)
        {
            if (ele == null)
                return "";

            string typeName = ele.GetType().Name;

            if (ele is FamilyInstance)
            {
                FamilyInstance fi = ele as FamilyInstance;
                typeName = "FamilyInstance: " + fi.StructuralType.ToString();
            }

            string categoryName = (null == ele.Category) ? String.Empty : ele.Category.Name + " ";

            return string.Format("{0} <{1} | {2} | {3}>", typeName, ele.Id.IntegerValue, ele.Name, categoryName);

        }
    }
}
