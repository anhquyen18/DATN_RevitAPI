using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using MaterialDesignColors;
using MaterialDesignThemes.Wpf;
using RevitAPI_Quyen.MyWindows;

namespace RevitAPI_Quyen.Commands
{
    [TransactionAttribute(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    class SheetCreationCommand:IExternalCommand
    {
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

                SheetCreationWindow wd = new SheetCreationWindow();
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
