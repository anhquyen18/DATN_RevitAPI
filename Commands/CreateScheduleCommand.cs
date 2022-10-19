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
        BuiltInParameter[] bipList = new BuiltInParameter[] {BuiltInParameter.WALL_BASE_CONSTRAINT,BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM,
            BuiltInParameter.HOST_AREA_COMPUTED, BuiltInParameter.CURVE_ELEM_LENGTH};
        Dictionary<BuiltInParameter, string> wallHeadingList = new Dictionary<BuiltInParameter, string>() {
           { BuiltInParameter.WALL_BASE_CONSTRAINT, "TẦNG" }, { BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM, "LOẠI FAMILY" },
        { BuiltInParameter.HOST_AREA_COMPUTED, "DIỆN TÍCH (m2)" }, { BuiltInParameter.CURVE_ELEM_LENGTH, "CHIỀU DÀI TƯỜNG (mm)" },
        { BuiltInParameter.HOST_VOLUME_COMPUTED, "THỂ TÍCH (m3)" }
        };
        string[] availableScheduleList = new string[] { "THỐNG KÊ TƯỜNG XÂY", "TỔNG HỢP THỐNG KÊ BÊTÔNG DẦM",
            "TỔNG HỢP THỐNG KÊ BÊTÔNG CỘT", "TỔNG HỢP THỐNG KÊ BÊTÔNG MÓNG" };
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Application app = uiApp.Application;
            Document doc = uiDoc.Document;
            Selection selec = uiDoc.Selection;

            //UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            //Document doc = uiDoc.Document;
            
            //Transaction t = new Transaction(doc, "Create Schedule");
            //t.Start();
            // Thử tạo footer thêm lần nữa;
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

                // THỐNG KÊ TƯỜNG XÂY
                ElementId wallCategoryId = new ElementId(BuiltInCategory.OST_Walls);
                ViewSchedule wallSchedule = ViewSchedule.CreateSchedule(doc, wallCategoryId);
                ScheduleSortGroupField baseConstraintSorting = null;
                //ScheduleFilter BaseConstraintFilter = null;
                //string levelNameFilter = "TẦNG 1";
                wallSchedule.Name = "THỐNG KẾ TƯỜNG XÂY";
                wallSchedule.Definition.ShowGrandTotal = true;
                //foreach (KeyValuePair<BuiltInParameter, string> pair in wallHeadingList)
                //{
                //    foreach (SchedulableField sf in wallSchedule.Definition.GetSchedulableFields())
                //    {
                //        if (CheckField(pair.Key, sf))
                //        {
                //            ScheduleField scheduleField = wallSchedule.Definition.AddField(sf);
                //            scheduleField.ColumnHeading = pair.Value;
                //            if (CheckField(BuiltInParameter.WALL_BASE_CONSTRAINT, sf))
                //            {
                //                baseConstraintSorting = new ScheduleSortGroupField(scheduleField.FieldId);
                //                //baseConstraintSorting.ShowHeader = true;
                //                wallSchedule.Definition.AddSortGroupField(baseConstraintSorting);
                //            }
                //            ShowTotal(BuiltInParameter.HOST_VOLUME_COMPUTED, UnitTypeId.CubicMeters, sf, scheduleField);
                //        }
                //    }
                //}



                //t.Commit();
                //uiDoc.ActiveView = wallSchedule;

                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }
        }

        public bool CheckField(BuiltInParameter bip, SchedulableField sf)
        {
            if (new ElementId(bip) == sf.ParameterId)
            {
                return true;
            }
            return false;
        }

        public void ShowTotal(BuiltInParameter bip, ForgeTypeId unitTypeId, SchedulableField sf, ScheduleField sF)
        {
            if (CheckField(bip, sf))
            {
                FormatOptions fo = new FormatOptions();
                fo.UseDefault = false;
                fo.SetUnitTypeId(unitTypeId);
                fo.Accuracy = 0.01;
                sF.SetFormatOptions(fo);
                sF.DisplayType = ScheduleFieldDisplayType.Totals;
            }
        }

        public Element GetLevelByName(Document doc, string name)
        {
            Level level = new FilteredElementCollector(doc).OfClass(typeof(Level))
                .Cast<Level>().Where(x => x.Name == name).FirstOrDefault();

            return level;
        }

    }
}
