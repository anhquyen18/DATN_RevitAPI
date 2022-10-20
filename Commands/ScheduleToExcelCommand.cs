using MaterialDesignColors;
using MaterialDesignThemes.Wpf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.UI.Selection;
using System.Collections.ObjectModel;
using RevitAPI_Quyen.ViewModel;
using RevitAPI_Quyen.MyWindows;

namespace RevitAPI_Quyen.Commands
{
    [TransactionAttribute(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ScheduleToExcelCommand : IExternalCommand
    {

        //private ObservableCollection<ScheduleItem> _scheduleList;
        public static string path { get; set; } = @"C:\RebarShape";
        public ObservableCollection<ScheduleItem> scheduleList { get; set; }
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Application app = uiApp.Application;
            Document doc = uiDoc.Document;
            Selection selec = uiDoc.Selection;

            scheduleList = FilterScheduleNameList(doc);
            try
            {
                ColorZoneAssist.SetMode(new CheckBox(), ColorZoneMode.Standard);
                Hue hue = new Hue("name", System.Windows.Media.Color.FromArgb(1, 2, 3, 4), System.Windows.Media.Color.FromArgb(1, 5, 6, 7));


                ScheduleToExcelWindow wd = new ScheduleToExcelWindow();
                wd.RevitListView.ItemsSource = scheduleList;
                wd.Doc = doc;
                wd.App = app;
                wd.ShowDialog();


                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }


        }

        public ObservableCollection<ScheduleItem> FilterScheduleNameList(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.WhereElementIsNotElementType();
            collector.WhereElementIsViewIndependent();
            collector.OfClass(typeof(ViewSchedule));

            var sql = from e in collector orderby e.Name ascending select e;

            ObservableCollection<ScheduleItem> list = new ObservableCollection<ScheduleItem>();

            foreach (Element e in sql.ToList<Element>())
            {
                ViewSchedule vs = e as ViewSchedule;
                if (!vs.IsTemplate)
                    list.Add(new ScheduleItem(vs.Id.IntegerValue, vs.Name));
            }

            return list;
        }

    }
}
