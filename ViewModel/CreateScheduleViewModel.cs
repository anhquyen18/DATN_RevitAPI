using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace RevitAPI_Quyen.ViewModel
{
    [TransactionAttribute(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CreateScheduleViewModel : BaseViewModel
    {
        #region commands
        public ICommand SelectFileCommand { get; set; }
        public ICommand BrowsePathCommand { get; set; }
        public ICommand AddToLeftCommand { get; set; }
        public ICommand AddToRightCommand { get; set; }
        public ICommand SelectAllRevitScheduleCommand { get; set; }
        public ICommand ExportCommand { get; set; }
        #endregion

        #region binding variables

        #endregion

        #region revit variables
        private Autodesk.Revit.ApplicationServices.Application _App;
        public Autodesk.Revit.ApplicationServices.Application App { set { _App = value; } }
        private Document _Doc;
        public Document Doc { get => _Doc; set { _Doc = value; } }
        BuiltInParameter[] bipList = new BuiltInParameter[] {BuiltInParameter.WALL_BASE_CONSTRAINT,BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM,
            BuiltInParameter.HOST_AREA_COMPUTED, BuiltInParameter.CURVE_ELEM_LENGTH};
        Dictionary<BuiltInParameter, string> wallHeadingList = new Dictionary<BuiltInParameter, string>() {
           { BuiltInParameter.WALL_BASE_CONSTRAINT, "TẦNG" }, { BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM, "LOẠI FAMILY" },
        { BuiltInParameter.HOST_AREA_COMPUTED, "DIỆN TÍCH (m2)" }, { BuiltInParameter.CURVE_ELEM_LENGTH, "CHIỀU DÀI TƯỜNG (mm)" },
        { BuiltInParameter.HOST_VOLUME_COMPUTED, "THỂ TÍCH (m3)" }
        };
        string[] availableScheduleList = new string[] { "THỐNG KÊ TƯỜNG XÂY", "TỔNG HỢP THỐNG KÊ BÊTÔNG DẦM",
            "TỔNG HỢP THỐNG KÊ BÊTÔNG CỘT", "TỔNG HỢP THỐNG KÊ BÊTÔNG MÓNG" };
        #endregion
        public CreateScheduleViewModel()
        {
            ExportCommand = new RelayCommand<object>((p) => { return true; }, (p) =>
            {
                Transaction t = new Transaction(Doc, "Create Schedule");
                t.Start();
                TaskDialog noti = new TaskDialog("Chú ý");
                noti.MainContent = "this is notificaftion";
                noti.Show();
                try
                {
                    ElementId wallCategoryId = new ElementId(BuiltInCategory.OST_Walls);
                    ViewSchedule wallSchedule = ViewSchedule.CreateSchedule(Doc, wallCategoryId);
                    ScheduleSortGroupField baseConstraintSorting = null;
                    //ScheduleFilter BaseConstraintFilter = null;
                    //string levelNameFilter = "TẦNG 1";
                    wallSchedule.Name = "THỐNG KẾ TƯỜNG XÂY";
                    wallSchedule.Definition.ShowGrandTotal = true;
                    foreach (KeyValuePair<BuiltInParameter, string> pair in wallHeadingList)
                    {
                        foreach (SchedulableField sf in wallSchedule.Definition.GetSchedulableFields())
                        {
                            if (CheckField(pair.Key, sf))
                            {
                                ScheduleField scheduleField = wallSchedule.Definition.AddField(sf);
                                scheduleField.ColumnHeading = pair.Value;
                                if (CheckField(BuiltInParameter.WALL_BASE_CONSTRAINT, sf))
                                {
                                    baseConstraintSorting = new ScheduleSortGroupField(scheduleField.FieldId);
                                    //baseConstraintSorting.ShowHeader = true;
                                    wallSchedule.Definition.AddSortGroupField(baseConstraintSorting);
                                }
                                ShowTotal(BuiltInParameter.HOST_VOLUME_COMPUTED, UnitTypeId.CubicMeters, sf, scheduleField);

                            }
                        }
                    }
                }
                catch
                {

                }

                t.Commit();
            });
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


    }
}
