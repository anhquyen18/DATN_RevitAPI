using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPI_Quyen.Model
{
    class FieldItem
    {
        public string FieldName { get; set; }

        public int ParamId { get; set; }

        public bool IsEnable { get; set; }

        public int FieldOrder { get; set; }

        public int TotalColumn { get; set; }

        public FieldItem(int id, String fieldName, int fieldOrder, int totalCol, bool isEnable = true)
        {
            ParamId = id;
            FieldName = fieldName;
            IsEnable = isEnable;
            FieldOrder = fieldOrder;
            TotalColumn = totalCol;
        }

        public override string ToString()
        {
            return ParamId + " | " + FieldName + " | " + IsEnable + " | " + FieldOrder;
        }
    }
}
