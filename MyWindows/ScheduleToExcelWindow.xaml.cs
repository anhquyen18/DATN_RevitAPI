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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using MaterialDesignColors;
using MaterialDesignThemes.Wpf;
using Excel = Microsoft.Office.Interop.Excel;
using RevitAPI_Quyen.ViewModel;
using RevitAPI_Quyen.Commands;
using RevitAPI_Quyen.Model;

namespace RevitAPI_Quyen.MyWindows
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class ScheduleToExcelWindow : Window
    {
        public ScheduleToExcelViewModel Viewmodel { get; set; }
        private Autodesk.Revit.ApplicationServices.Application _App;
        public Autodesk.Revit.ApplicationServices.Application App { get => _App; set { _App = value; } }
        private Document _Doc;
        public Document Doc { get => _Doc; set { _Doc = value; } }
        public ScheduleToExcelWindow()
        {
            ColorZoneAssist.SetMode(new CheckBox(), ColorZoneMode.Light);
            Hue hue = new Hue("xyz", System.Windows.Media.Color.FromArgb(1, 2, 3, 4), System.Windows.Media.Color.FromArgb(1, 5, 6, 7));
            this.DataContext = Viewmodel = new ScheduleToExcelViewModel();

            InitializeComponent();

            this.Loaded += new RoutedEventHandler(MainWindow_Loaded);
        }

        void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            Viewmodel.App = this.App;
            Viewmodel.Doc = this.Doc;
        }

        private void RevitList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Viewmodel.RevitSelectedItems = new ObservableCollection<ScheduleItem>();
            foreach (ScheduleItem item in RevitListView.SelectedItems)
            {
                Viewmodel.RevitSelectedItems.Add(item);
            }
        }
        private void ToExcelList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Viewmodel.ToExcelSelectedItems = new ObservableCollection<ScheduleItem>();
            foreach (ScheduleItem item in ToExcelListView.SelectedItems)
            {
                Viewmodel.ToExcelSelectedItems.Add(item);
            }
        }


    }

    
}
