using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPI_Quyen.ViewModel
{
    public class ScheduleItem
    {
        public string ScheduleName
        {
            get;
            set;
        }
        public int Id
        {
            get;
            set;
        }
        public ScheduleItem(int id, string name)
        {
            Id = id;
            ScheduleName = name;
        }
        public MyScheduleField[] ScheduleFieldList { get; set; }
        public ScheduleItem(string ScheduleName, MyScheduleField[] ScheduleFieldList)
        {
            this.ScheduleName = ScheduleName;
            this.ScheduleFieldList = ScheduleFieldList;
        }



    }
}
