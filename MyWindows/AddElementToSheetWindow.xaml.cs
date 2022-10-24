using Autodesk.Revit.DB;
using MaterialDesignColors;
using MaterialDesignThemes.Wpf;
using RevitAPI_Quyen.Model;
using RevitAPI_Quyen.ViewModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace RevitAPI_Quyen.MyWindows
{
    /// <summary>
    /// Interaction logic for AddElementToSheetWindow.xaml
    /// </summary>
    public partial class AddElementToSheetWindow : Window
    {
        private Autodesk.Revit.ApplicationServices.Application _App;
        public Autodesk.Revit.ApplicationServices.Application App { get => _App; set { _App = value; } }
        private Document _Doc;
        public Document Doc { get => _Doc; set { _Doc = value; } }
        AddElementToSheetViewModel Viewmodel { get; set; }
        public AddElementToSheetWindow()
        {
            ColorZoneAssist.SetMode(new CheckBox(), ColorZoneMode.Standard);
            Hue hue = new Hue("name", System.Windows.Media.Color.FromArgb(1, 2, 3, 4), System.Windows.Media.Color.FromArgb(1, 5, 6, 7));

            this.DataContext = Viewmodel = new AddElementToSheetViewModel();
            InitializeComponent();

            this.Loaded += new RoutedEventHandler(MainWindow_Loaded);
        }

        void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            Viewmodel.App = App;
            Viewmodel.Doc = Doc;
            InitComboboxViewportType();
        }

        void InitComboboxViewportType()
        {
            FilteredElementCollector col = new FilteredElementCollector(Doc);
            IList<ElementType> viewportTypes = col.OfClass(typeof(ElementType)).Cast<ElementType>().Where(q => q.FamilyName == "Viewport").ToList();

            Viewmodel.ViewportTypeList = new ObservableCollection<ViewportTypeItem>();
            Viewmodel.ViewportTypeList.Clear();
            foreach (ElementType ele in viewportTypes)
            {
                ViewportTypeItem item = new ViewportTypeItem();
                item.Name = ele.Name;
                item.Value = ele;
                Viewmodel.ViewportTypeList.Add(item);
            }
        }


    }
}
