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
            string tabName = "BIM_DUT";
            application.CreateRibbonTab(tabName);
            RibbonPanel ExcelPn = application.CreateRibbonPanel(tabName, "Ex/Import Excel");
            RibbonPanel createSchedulePn = application.CreateRibbonPanel(tabName, "Create Schedule");
            RibbonPanel sheetEditorPn = application.CreateRibbonPanel(tabName, "Sheet Editor");

            string path = Assembly.GetExecutingAssembly().Location;

            PushButtonData toExcelBt = new PushButtonData("toExcelBt", "Schedule to Excel", path, "RevitAPI_Quyen.Commands.ScheduleToExcelCommand");
            PushButtonData createScheduleBt = new PushButtonData("creatExcelBt", "Create schedules", path, "RevitAPI_Quyen.Commands.CreateScheduleCommand");
            PushButtonData sheetCreationBt = new PushButtonData("sheetCreationBt", "Create sheets", path, "RevitAPI_Quyen.Commands.SheetCreationCommand");

            //panel.AddItem(toExcelBt);

            PushButton toExcel = ExcelPn.AddItem(toExcelBt) as PushButton;
            PushButton createSchedule = createSchedulePn.AddItem(createScheduleBt) as PushButton;
            PushButton sheetCreation = sheetEditorPn.AddItem(sheetCreationBt) as PushButton;

            Uri wrefImageUri = new Uri(@"F:\University\Hoc ky 9\DATN\RevitAPI_Quyen\Resources\logoCTT29x32.png");
            BitmapImage wrefImage = new BitmapImage(wrefImageUri);

            toExcel.LargeImage = wrefImage;
            createSchedule.LargeImage = wrefImage;
            sheetCreation.LargeImage = wrefImage;

            return Result.Succeeded;
        }
    }
}
