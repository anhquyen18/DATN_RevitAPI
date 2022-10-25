using Autodesk.Revit.DB;
using RevitAPI_Quyen.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace RevitAPI_Quyen.ViewModel
{
    public class AddElementToSheetViewModel : BaseViewModel
    {
        #region commands
        public ICommand AddToLeftCommand { get; set; }
        public ICommand AddToRightCommand { get; set; }
        public ICommand ExportCommand { get; set; }
        #endregion
        #region binding variables
        private ObservableCollection<ViewportTypeItem> _ViewportTypeList;
        public ObservableCollection<ViewportTypeItem> ViewportTypeList { get => _ViewportTypeList; set { _ViewportTypeList = value; OnPropertyChanged(); } }
        private ObservableCollection<SheetItem> _SheetList;
        public ObservableCollection<SheetItem> SheetList { get => _SheetList; set { _SheetList = value; OnPropertyChanged(); } }
        private string _SearchText;
        public string SearchText { get => _SearchText; set { _SearchText = value; OnPropertyChanged(); } }
        #endregion

        #region revit variables
        private Autodesk.Revit.ApplicationServices.Application _App;
        public Autodesk.Revit.ApplicationServices.Application App { get => _App; set { _App = value; } }
        private Document _Doc;
        public Document Doc { get => _Doc; set { _Doc = value; } }
        #endregion

        public AddElementToSheetViewModel()
        {
            
        }

        public enum FilterViewOption
        {
            VIEW = 1,
            SCHEDULE = 2,
            LEGEND = 3,
            SCHEDULE_OR_LENGEND = 4
        }
    }
}
