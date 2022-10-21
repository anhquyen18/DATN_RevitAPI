using Autodesk.Revit.DB;
using MaterialDesignColors;
using MaterialDesignThemes.Wpf;
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
    
    public partial class CreateScheduleWindow : Window
    {
        private Autodesk.Revit.ApplicationServices.Application _App;
        public Autodesk.Revit.ApplicationServices.Application App { get => _App; set { _App = value; } }
        private Document _Doc;
        public Document Doc { get => _Doc; set { _Doc = value; } }

        public CreateScheduleViewModel Viewmodel { get; set; }
        public CreateScheduleWindow()
        {
            ColorZoneAssist.SetMode(new CheckBox(), ColorZoneMode.Standard);
            Hue hue = new Hue("name", System.Windows.Media.Color.FromArgb(1, 2, 3, 4), System.Windows.Media.Color.FromArgb(1, 5, 6, 7));
            
            this.DataContext = Viewmodel = new CreateScheduleViewModel();

            InitializeComponent();

            this.Loaded += new RoutedEventHandler(MainWindow_Loaded);
        }

        void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            Viewmodel.App = this.App;
            Viewmodel.Doc = this.Doc;
        }

        private void ActiveListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Viewmodel.ActiveSelectedItems = new ObservableCollection<ScheduleItem>();
            foreach (ScheduleItem item in ActiveListView.SelectedItems)
            {
                Viewmodel.ActiveSelectedItems.Add(item);
            }
        }

        private void CreationListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Viewmodel.CreationSelectedItems = new ObservableCollection<ScheduleItem>();
            foreach (ScheduleItem item in CreationListView.SelectedItems)
            {
                Viewmodel.CreationSelectedItems.Add(item);
            }
        }
    }
}
