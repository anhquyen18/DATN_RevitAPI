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
        private int _CreationSelectedIndex;
        public int CreationSelectedIndex { get => _CreationSelectedIndex; set { _CreationSelectedIndex = value; OnPropertyChanged(); } }
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
    }
}
