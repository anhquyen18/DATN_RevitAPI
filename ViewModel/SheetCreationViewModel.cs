using Autodesk.Revit.DB;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using Excel = Microsoft.Office.Interop.Excel;
using RevitAPI_Quyen.Model;

namespace RevitAPI_Quyen.ViewModel
{
    public class SheetCreationViewModel : BaseViewModel
    {
        #region commands
        public ICommand SelectFileCommand { get; set; }
        public ICommand CreateSheetCommand { get; set; }
        #endregion

        #region binding variables
        private string _ExcelFilePath;
        public string ExcelFilePath { get => _ExcelFilePath; set { _ExcelFilePath = value; OnPropertyChanged(); } }
        private string _LogText;
        public string LogText { get => _LogText; set { _LogText = value; OnPropertyChanged(); } }
        private int _TitleblockComboSelectedIndex;
        public int TitleblockComboSelectedIndex { get => _TitleblockComboSelectedIndex; set { _TitleblockComboSelectedIndex = value; OnPropertyChanged(); } }
        private ObservableCollection<TitleblockItem> _TitleblockList;
        public ObservableCollection<TitleblockItem> TitleblockList { get => _TitleblockList; set { _TitleblockList = value; OnPropertyChanged(); } }
        private TitleblockItem _TitleblockComboSelectedItem;
        public TitleblockItem TitleblockComboSelectedItem { get => _TitleblockComboSelectedItem; set { _TitleblockComboSelectedItem = value; OnPropertyChanged(); } }
        #endregion

        #region revit variables
        private Autodesk.Revit.ApplicationServices.Application _App;
        public Autodesk.Revit.ApplicationServices.Application App { set { _App = value; } }
        private Document _Doc;
        public Document Doc { get => _Doc; set { _Doc = value; } }
        private Excel.Application xlApp;
        #endregion

        public SheetCreationViewModel()
        {
            SelectFileCommand = new RelayCommand<object>((p) => { return true; }, (p) =>
            {
                OpenFileDialog selectFileDialog = new OpenFileDialog();
                selectFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                if (selectFileDialog.ShowDialog() == true)
                    ExcelFilePath = selectFileDialog.FileName;
            });

            
            
            CreateSheetCommand = new RelayCommand<object>((p) => { return true; }, (p) =>
            {
                if (ExcelFilePath == "")
                {
                    MessageBox.Show("Please select a Excel file.", "Excel File");
                    return;
                }
                if(TitleblockComboSelectedIndex == -1)
                {
                    MessageBox.Show("Please select a Titleblock.", "Titleblock");
                    return;
                }

                MessageBoxResult result = MessageBox.Show("Do you want to create sheets?", "Confirm Dialog", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.No)
                    return;

                try
                {
                    LogText = "";
                    CreateNewSheets();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {


                    if (xlApp != null)
                    {
                        xlApp.Quit();
                        ReleaseObject(xlApp);
                    }
                }

            });
        }

        public void AddLog(string log)
        {
            LogText += log;
        }

        private void CreateNewSheets()
        {
            xlApp = new Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Microsoft Excel can't be found. Please install it before run this add-in.");
                return;
            }

            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Open(ExcelFilePath,
                0,
                true,
                5,
                "",
                "",
                true,
                Excel.XlPlatform.xlWindows,
                "\t",
                false,
                false,
                0,
                true,
                1,
                0);

            Excel.Worksheet worksheet = xlWorkBook.Worksheets[1] as Excel.Worksheet;

            int start_excel_row = 2;
            int last_excel_row = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, misValue).Row;

            Transaction trans = new Transaction(Doc, "Create Sheet");

            ViewSheet curViewSheet = null;
            TitleblockItem titleBlock = TitleblockComboSelectedItem;

            string viewSheetName = "";
            string viewSheetNumber = "";
            //List<ViewSheet> sheetList = this.getAllViewSheets();

            for (int i = start_excel_row; i <= last_excel_row; i++)
            {
                try
                {
                    Excel.Range rangeSheetNumber = worksheet.Cells[i, 1] as Excel.Range;
                    Excel.Range rangePostfixNumber = worksheet.Cells[i, 2] as Excel.Range;
                    Excel.Range rangeSheetName = worksheet.Cells[i, 3] as Excel.Range;

                    viewSheetNumber = rangeSheetNumber.Value2;

                    if (!string.IsNullOrEmpty(Convert.ToString(rangePostfixNumber.Value2)))
                    {
                        int n = 0;
                        if (int.TryParse(Convert.ToString(rangePostfixNumber.Value2), out n))
                        {
                            string sn = "";
                            if (n < 10)
                            {
                                sn = "00" + n;
                            }
                            else if (n > 9 && n < 100)
                            {
                                sn = "0" + n;
                            }
                            else
                            {
                                sn = n.ToString();
                            }

                            viewSheetNumber += "-" + sn;
                        }
                        else
                        {
                            viewSheetNumber += "-" + rangePostfixNumber.Value2;
                        }
                    }

                    viewSheetName = viewSheetNumber;
                    if (!string.IsNullOrEmpty(Convert.ToString(rangeSheetName.Value2)))
                    {
                        viewSheetName += " - " + rangeSheetName.Value2;
                    }

                    //Ignore if the sheet number is already existed
                    if (IsSheetExisted(viewSheetNumber))
                    {
                        AddLog(viewSheetName + ": ");
                        AddLog("[ALREADY EXISTED]" + Environment.NewLine);
                        continue;
                    }

                    trans.Start();
                    curViewSheet = ViewSheet.Create(Doc, titleBlock.Value.Id);
                    curViewSheet.SheetNumber = viewSheetNumber;
                    if (!string.IsNullOrEmpty(Convert.ToString(rangeSheetName.Value2)))
                        curViewSheet.Name = rangeSheetName.Value2;

                    trans.Commit();

                    AddLog(viewSheetName + ": ");
                    AddLog("[CREATED]" + Environment.NewLine);
                }
                catch (Exception ex)
                {
                    AddLog(viewSheetName + ": ");
                    AddLog("[FAILED]" + Environment.NewLine);
                    AddLog(ex.ToString() + Environment.NewLine);
                    if (trans.HasStarted())
                        trans.RollBack();
                }
                finally
                {

                }
            }
        }


        private bool IsSheetExisted(string pSheetNumber)
        {
            FilteredElementCollector col = new FilteredElementCollector(Doc);
            col.OfCategory(BuiltInCategory.OST_Sheets);

            foreach (ViewSheet item in col.ToElements())
            {
                if (item.SheetNumber.Equals(pSheetNumber))
                    return true;
            }
            return false;
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

    }
}
