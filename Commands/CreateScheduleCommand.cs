using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using MaterialDesignColors;
using MaterialDesignThemes.Wpf;
using RevitAPI_Quyen.MyWindows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace RevitAPI_Quyen.Commands
{
    [TransactionAttribute(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    class CreateScheduleCommand : IExternalCommand
    {
        //Revit API 2022 chưa cho tạo caculated field
        //Chỉ có thể tạo được những field mặc định có sẵn
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Application app = uiApp.Application;
            Document doc = uiDoc.Document;
            Selection selec = uiDoc.Selection;

            try
            {
                ColorZoneAssist.SetMode(new CheckBox(), ColorZoneMode.Standard);
                Hue hue = new Hue("name", System.Windows.Media.Color.FromArgb(1, 2, 3, 4), System.Windows.Media.Color.FromArgb(1, 5, 6, 7));

                TaskDialog noti = new TaskDialog("Chú ý");
                noti.MainContent = "Để đảm bảo không lỗi tên cấu kiện cấu kiện và các chú thích khác nếu có phải là TIẾNG VIỆT IN HOA CÓ DẤU";
                noti.Show();

                CreateScheduleWindow wd = new CreateScheduleWindow();
                wd.App = app;
                wd.Doc = doc;
                wd.ShowDialog();



                //t.Commit();
                //uiDoc.ActiveView = wallSchedule;

                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }
        }

    }
}
