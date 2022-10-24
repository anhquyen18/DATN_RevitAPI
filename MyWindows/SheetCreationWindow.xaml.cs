using Autodesk.Revit.DB;
using MaterialDesignColors;
using MaterialDesignThemes.Wpf;
using RevitAPI_Quyen.ViewModel;
using RevitAPI_Quyen.Model;
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
    /// Interaction logic for SheetCreationWindow.xaml
    /// </summary>
    public partial class SheetCreationWindow : Window
    {
        private Autodesk.Revit.ApplicationServices.Application _App;
        public Autodesk.Revit.ApplicationServices.Application App { get => _App; set { _App = value; } }
        private Document _Doc;
        public Document Doc { get => _Doc; set { _Doc = value; } }

        public SheetCreationViewModel Viewmodel { get; set; }
        public SheetCreationWindow()
        {
            ColorZoneAssist.SetMode(new CheckBox(), ColorZoneMode.Standard);
            Hue hue = new Hue("name", System.Windows.Media.Color.FromArgb(1, 2, 3, 4), System.Windows.Media.Color.FromArgb(1, 5, 6, 7));
            this.DataContext = Viewmodel = new SheetCreationViewModel();

            InitializeComponent();

            this.Loaded += new RoutedEventHandler(MainWindow_Loaded);
            
        }

        void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            Viewmodel.App = this.App;
            Viewmodel.Doc = this.Doc;
            InitComboBoxTitleblock();
        }

        public void InitComboBoxTitleblock()
        {
            FilteredElementCollector col = new FilteredElementCollector(Doc);
            col.WhereElementIsElementType();
            col.OfCategory(BuiltInCategory.OST_TitleBlocks);

            Viewmodel.TitleblockList = new ObservableCollection<TitleblockItem>();
            Viewmodel.TitleblockList.Clear();
            foreach (Element e in col.ToElements())
            {
                TitleblockItem item = new TitleblockItem();
                item.Name = e.Name;
                item.Value = e;
                Viewmodel.TitleblockList.Add(item);
            }

        }

    }
}
