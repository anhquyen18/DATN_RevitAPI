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
        public ICommand AddToLeftCommand { get; set; }
        public ICommand AddToRightCommand { get; set; }
        public ICommand ExportCommand { get; set; }
        #endregion

        #region binding variables
        private ObservableCollection<ScheduleItem> _ActiveScheduleList;
        public ObservableCollection<ScheduleItem> ActiveScheduleList { get => _ActiveScheduleList; set { _ActiveScheduleList = value; OnPropertyChanged(); } }
        private ObservableCollection<ScheduleItem> _CreationScheduleList;
        public ObservableCollection<ScheduleItem> CreationScheduleList { get => _CreationScheduleList; set { _CreationScheduleList = value; OnPropertyChanged(); } }
        private ObservableCollection<ScheduleItem> _ActiveSelectedItems;
        public ObservableCollection<ScheduleItem> ActiveSelectedItems { get => _ActiveSelectedItems; set { _ActiveSelectedItems = value; OnPropertyChanged(); } }
        private ObservableCollection<ScheduleItem> _CreationSelectedItems;
        public ObservableCollection<ScheduleItem> CreationSelectedItems { get => _CreationSelectedItems; set { _CreationSelectedItems = value; OnPropertyChanged(); } }
        private int _ActiveSelectedIndex;
        public int ActiveSelectedIndex { get => _ActiveSelectedIndex; set { _ActiveSelectedIndex = value; OnPropertyChanged(); } }
        private int _CreationSelectedIndex;
        public int CreationSelectedIndex { get => _CreationSelectedIndex; set { _CreationSelectedIndex = value; OnPropertyChanged(); } }
        #endregion

        #region revit variables
        private Autodesk.Revit.ApplicationServices.Application _App;
        public Autodesk.Revit.ApplicationServices.Application App { set { _App = value; } }
        private Document _Doc;
        public Document Doc { get => _Doc; set { _Doc = value; } }
        private Dictionary<BuiltInParameter, string> wallHeadingList = new Dictionary<BuiltInParameter, string>() {
           { BuiltInParameter.WALL_BASE_CONSTRAINT, "TẦNG" }, { BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM, "LOẠI FAMILY" },
        { BuiltInParameter.HOST_AREA_COMPUTED, "DIỆN TÍCH (m2)" }, { BuiltInParameter.CURVE_ELEM_LENGTH, "CHIỀU DÀI TƯỜNG (mm)" },
        { BuiltInParameter.HOST_VOLUME_COMPUTED, "THỂ TÍCH (m3)" }
        };
        private MyScheduleField[] wallScheduleFieldList = new MyScheduleField[] {
            new MyScheduleField(BuiltInParameter.WALL_BASE_CONSTRAINT, "TẦNG", null),
            new MyScheduleField(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM, "LOẠI FAMILY", null),
            new MyScheduleField(BuiltInParameter.HOST_AREA_COMPUTED, "DIỆN TÍCH (m2)", UnitTypeId.SquareMeters),
            new MyScheduleField(BuiltInParameter.CURVE_ELEM_LENGTH, "CHIỀU DÀI TƯỜNG (mm)", UnitTypeId.Millimeters),
            new MyScheduleField(BuiltInParameter.HOST_VOLUME_COMPUTED, "THỂ TÍCH (m3)", UnitTypeId.CubicMeters),
        };
        private BuiltInParameter[] showTotalList = new BuiltInParameter[] { 
            BuiltInParameter.CURVE_ELEM_LENGTH,
            BuiltInParameter.HOST_VOLUME_COMPUTED
        };
        #endregion

        public CreateScheduleViewModel()
        {
            ActiveScheduleList = new ObservableCollection<ScheduleItem>() {
                new ScheduleItem( "THỐNG KÊ TƯỜNG XÂY", wallScheduleFieldList),
                new ScheduleItem( "TỔNG HỢP THỐNG KÊ BÊTÔNG DẦM", wallScheduleFieldList),
                new ScheduleItem( "TỔNG HỢP THỐNG KÊ BÊTÔNG CỘT", wallScheduleFieldList),
                new ScheduleItem( "TỔNG HỢP THỐNG KÊ BÊTÔNG MÓNG", wallScheduleFieldList),
            };
            CreationScheduleList = new ObservableCollection<ScheduleItem>();

            AddToRightCommand = new RelayCommand<object>((p) => { return true; }, (p) =>
            {
                int n = ActiveScheduleList.Count;
                if (n > 0)
                {
                    if (ActiveSelectedIndex == -1)
                        return;
                    foreach (ScheduleItem item in ActiveSelectedItems.ToList<ScheduleItem>())
                    {
                        CreationScheduleList.Add(item);
                        ActiveScheduleList.Remove(item);
                    }
                }
                // Mỗi lần chuyển là sắp sếp lại List nhưng chưa sắp xếp được
                CreationScheduleList.OrderBy(s => s.ScheduleName);
                ActiveScheduleList.OrderBy(s => s.ScheduleName);
            });
            AddToLeftCommand = new RelayCommand<object>((p) => { return true; }, (p) =>
            {
                int n = CreationScheduleList.Count;
                if (n > 0)
                {
                    if (CreationSelectedIndex == -1)
                        return;
                    foreach (ScheduleItem item in CreationSelectedItems.ToList<ScheduleItem>())
                    {
                        ActiveScheduleList.Add(item);
                        CreationScheduleList.Remove(item);
                    }
                }
                // Mỗi lần chuyển là sắp sếp lại List nhưng chưa sắp xếp được
                CreationScheduleList.OrderBy(s => s.ScheduleName);
                ActiveScheduleList.OrderBy(s => s.ScheduleName);
            });
            ExportCommand = new RelayCommand<object>((p) => { return true; }, (p) =>
            {
                Transaction t = new Transaction(Doc, "Create Schedule");
                t.Start();
                TaskDialog noti = new TaskDialog("Chú ý");
                noti.MainContent = "this is notificaftion";
                noti.Show();
                ElementId wallCategoryId = new ElementId(BuiltInCategory.OST_Walls);
                ViewSchedule wallSchedule = ViewSchedule.CreateSchedule(Doc, wallCategoryId);
                ScheduleSortGroupField baseConstraintSorting = null;
                //ScheduleFilter BaseConstraintFilter = null;
                //string levelNameFilter = "TẦNG 1";
                wallSchedule.Definition.ShowGrandTotal = true;
                foreach (ScheduleItem item in CreationScheduleList)
                {
                    if (item.ScheduleName == "THỐNG KÊ TƯỜNG XÂY")
                    {
                        wallSchedule.Name = item.ScheduleName;
                        foreach (MyScheduleField schedulefield in item.ScheduleFieldList)
                        {
                            foreach (SchedulableField schedulablefield in wallSchedule.Definition.GetSchedulableFields())
                            {
                                if (CheckField(schedulefield.BIParameter, schedulablefield))
                                {
                                    ScheduleField scheduleField = wallSchedule.Definition.AddField(schedulablefield);
                                    scheduleField.ColumnHeading = schedulefield.ScheduleFieldName;
                                    if (CheckField(BuiltInParameter.WALL_BASE_CONSTRAINT, schedulablefield))
                                    {
                                        baseConstraintSorting = new ScheduleSortGroupField(scheduleField.FieldId);
                                        baseConstraintSorting.ShowHeader = true;
                                        wallSchedule.Definition.AddSortGroupField(baseConstraintSorting);
                                    }
                                    if (showTotalList.Contains(schedulefield.BIParameter))
                                        ShowTotal(schedulefield.BIParameter, schedulefield.Unit, schedulablefield, scheduleField);

                                }
                            }
                        }
                    }
                }

                t.Commit();
            });
        }

        //public void CreateWallSchedule(ElementId categoryId, )

        public bool CheckField(BuiltInParameter bip, SchedulableField sf)
        {
            if (new ElementId(bip) == sf.ParameterId)
            {
                return true;
            }
            return false;
        }

        public void ShowTotal(BuiltInParameter bip, ForgeTypeId unit, SchedulableField sf, ScheduleField sF)
        {
            if (CheckField(bip, sf))
            {
                FormatOptions fo = new FormatOptions();
                fo.UseDefault = false;
                if (unit != null)
                    fo.SetUnitTypeId(unit);
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
