using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;

namespace RevitAPI_Quyen
{
    [TransactionAttribute(TransactionMode.Manual)]
    class RibbonTab : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {

            string nameTab = "BIM_DUT";
            application.CreateRibbonTab(nameTab);
            RibbonPanel panel = application.CreateRibbonPanel(nameTab, "Ex/Import Excel");
            RibbonPanel panel2 = application.CreateRibbonPanel(nameTab, "Test Panel");

            string path = Assembly.GetExecutingAssembly().Location;
            PushButtonData toExcelBt = new PushButtonData("toExcelBt", "Schedule to Excel", path, "RevitAPI_Quyen.Commands.ScheduleToExcelCommand");

            //panel.AddItem(toExcelBt);

            PushButton toExcel = panel.AddItem(toExcelBt) as PushButton;

            Uri wrefImageUri = new Uri(@"F:\University\Hoc ky 9\DATN\RevitAPI_Quyen\Resources\logoCTT29x32.png");
            BitmapImage wrefImage = new BitmapImage(wrefImageUri);
            toExcel.LargeImage = wrefImage;


            return Result.Succeeded;
        }
    }
}
