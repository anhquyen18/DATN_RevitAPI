using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitAPI_Quyen.Model
{
    public class MyScheduleField
    {
        public BuiltInParameter BIParameter { get; set; }
        public string ScheduleFieldName { get; set; }
        public ForgeTypeId Unit { get; set; }
        public MyScheduleField(BuiltInParameter bip, string sfName, ForgeTypeId unit)
        {
            BIParameter = bip;
            ScheduleFieldName = sfName;
            Unit = unit;
        }

    }
}
