using System.Linq;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System;
using RevitAPI_Quyen.Commands;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.UI.Selection;
using System.Windows;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace RevitAPI_Quyen.ViewModel
{
    public class ScheduleToExcelViewModel : BaseViewModel
    {
        #region commands
        public ICommand SelectFileCommand { get; set; }
        public ICommand BrowsePathCommand { get; set; }
        public ICommand AddToLeftCommand { get; set; }
        public ICommand AddToRightCommand { get; set; }
        public ICommand MoveDownCommand { get; set; }
        public ICommand MoveUpCommand { get; set; }
        public ICommand ExportCommand { get; set; }
        #endregion

        #region binding variables
        private ObservableCollection<ScheduleItem> _ToExcelSelectedItems;
        public ObservableCollection<ScheduleItem> ToExcelSelectedItems { get => _ToExcelSelectedItems; set { _ToExcelSelectedItems = value; OnPropertyChanged(); } }
        private ObservableCollection<ScheduleItem> _RevitSelectedItems;
        public ObservableCollection<ScheduleItem> RevitSelectedItems { get => _RevitSelectedItems; set { _RevitSelectedItems = value; OnPropertyChanged(); } }
        private string _TemplateFilePath;
        public string TemplateFilePath { get => _TemplateFilePath; set { _TemplateFilePath = value; OnPropertyChanged(); } }
        private string _SaveFilePath;
        public string SaveFilePath { get => _SaveFilePath; set { _SaveFilePath = value; OnPropertyChanged(); } }
        private ObservableCollection<ScheduleItem> _RevitScheduleList;
        public ObservableCollection<ScheduleItem> RevitScheduleList { get => _RevitScheduleList; set { _RevitScheduleList = value; OnPropertyChanged(); } }
        private ObservableCollection<ScheduleItem> _ToExcelScheduleList;
        public ObservableCollection<ScheduleItem> ToExcelScheduleList { get => _ToExcelScheduleList; set { _ToExcelScheduleList = value; OnPropertyChanged(); } }
        private int _ToExcelSelectedIndex;
        public int ToExcelSelectedIndex { get => _ToExcelSelectedIndex; set { _ToExcelSelectedIndex = value; OnPropertyChanged(); } }
        private int _RevitScheduleSelectedIndex;
        public int RevitScheduleSelectedIndex { get => _RevitScheduleSelectedIndex; set { _RevitScheduleSelectedIndex = value; OnPropertyChanged(); } }
        #endregion

        #region revit variables
        private Autodesk.Revit.ApplicationServices.Application _App;
        public Autodesk.Revit.ApplicationServices.Application App { get => _App; set { _App = value; } }
        private Document _Doc;
        public Document Doc { get => _Doc; set { _Doc = value; } }
        private const string TAG = "Schedule To Excel";
        private Microsoft.Office.Interop.Excel.Application xlApp;
        public const int DEFAULT_TITLE_ROW = 5;
        public const int DEFAULT_HEADER_ROW = 7;
        public const int DEFAULT_DATA_ROW = 8;
        public const string ELEMENT_UNIQUE_ID_PARAM_NAME = "Element_Unique_Id";
        int TitleRow = 0;
        int HeaderRow = 0;
        int DataRow = 0;
        #endregion

        public ScheduleToExcelViewModel()
        {

            #region file options panel
            TemplateFilePath = @"F:\University\Hoc ky 9\DATN\RevitAPI_Quyen\Resources\template.xlsx";
            SaveFilePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\exported_schedule.xlsx";
            //SaveFilePath = @"F:\University\Hoc ky 9\DATN\RevitAPI_Quyen" + @"\exported_schedule.xlsx";

            SelectFileCommand = new RelayCommand<object>((p) => { return true; }, (p) =>
            {
                OpenFileDialog selectFileDialog = new OpenFileDialog();
                selectFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                if (selectFileDialog.ShowDialog() == true)
                    TemplateFilePath = selectFileDialog.FileName;
            });

            #endregion

            #region revit schedule list panel
            RevitScheduleList = new ObservableCollection<ScheduleItem>();
            ToExcelScheduleList = new ObservableCollection<ScheduleItem>();

            AddToLeftCommand = new RelayCommand<object>((p) => { return true; }, (p) =>
            {
                int n = ToExcelScheduleList.Count;
                if (n > 0)
                {
                    if (ToExcelSelectedIndex == -1)
                        return;
                    foreach (ScheduleItem item in ToExcelSelectedItems.ToList<ScheduleItem>())
                    {
                        RevitScheduleList.Add(item);
                        ToExcelScheduleList.Remove(item);
                    }
                }
                // Mỗi lần chuyển là sắp sếp lại List nhưng chưa sắp xếp được
                ToExcelScheduleList.OrderBy(s => s.ScheduleName);
                RevitScheduleList.OrderBy(s => s.ScheduleName);
            });
            AddToRightCommand = new RelayCommand<object>((p) => { return true; }, (p) =>
            {
                int n = RevitScheduleList.Count;
                if (n > 0)
                {
                    if (RevitScheduleSelectedIndex == -1)
                        return;
                    foreach (ScheduleItem item in RevitSelectedItems.ToList<ScheduleItem>())
                    {
                        ToExcelScheduleList.Add(item);
                        RevitScheduleList.Remove(item);
                    }
                }
                ToExcelScheduleList.OrderBy(s => s.ScheduleName);
                RevitScheduleList.OrderBy(s => s.ScheduleName);
            });
            MoveDownCommand = new RelayCommand<object>((p) => { return true; }, (p) =>
            {
                int oldIndex = ToExcelSelectedIndex;
                if (oldIndex == -1)
                    return;
                ScheduleItem item = ToExcelScheduleList[oldIndex];
                if (oldIndex < ToExcelScheduleList.Count - 1)
                {
                    int newIndex = oldIndex + 1;
                    ToExcelScheduleList.Remove(item);
                    ToExcelScheduleList.Insert(newIndex, item);
                    ToExcelSelectedIndex = newIndex;
                }
            });
            MoveUpCommand = new RelayCommand<object>((p) => { return true; }, (p) =>
            {
                int oldIndex = ToExcelSelectedIndex;
                if (oldIndex == -1)
                    return;
                ScheduleItem item = ToExcelScheduleList[oldIndex];
                if (oldIndex > 0)
                {
                    int newIndex = oldIndex - 1;
                    ToExcelScheduleList.Remove(item);
                    ToExcelScheduleList.Insert(newIndex, item);
                    ToExcelSelectedIndex = newIndex;
                }
            });
            ExportCommand = new RelayCommand<object>((p) => { return true; }, (p) =>
            {

                if (!System.IO.File.Exists(TemplateFilePath))
                {
                    TaskDialog.Show(TAG, "The template file is empty or does not exist.");
                    return;
                }

                if (SaveFilePath == "")
                {
                    TaskDialog.Show(TAG, "The save as path is empty or does not exist.");
                    return;
                }
                else if (!System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(SaveFilePath)))
                {
                    TaskDialog.Show(TAG, "The save as path is empty or does not exist.");
                    return;
                }
                if (ToExcelScheduleList.Count == 0)
                {
                    TaskDialog.Show(TAG, "The exported list is empty.");
                    return;
                }

                MessageBoxResult result = MessageBox.Show("Do you want to export file?", "Confirm Dialog", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.No)
                    return;
                if (isFileInUse(new System.IO.FileInfo(SaveFilePath)))
                {
                    MessageBox.Show("The File: " + SaveFilePath + " is already in used by another process.Please close it, and try again.");
                    return;
                }

                try
                {
                    ExportToExcel();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());

                    if (xlApp != null)
                    {
                        xlApp.Quit();
                        releaseObject(xlApp);
                    }
                }

            });

            #endregion

        }



        public void ExportToExcel()
        {
            // lấy đường dẫn file
            string assemblyFilePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string assemplyDirPath = System.IO.Path.GetDirectoryName(assemblyFilePath);

            // Tạo một file excel mới
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("Microsoft Excel can't be found. Please install it before run this add-in.");
                return;
            }

            Excel.Workbook xlWorkbookTpl;
            Excel.Worksheet xlWorksheetTpl;
            Excel.Workbook xlWorkbook;

            object misValue = System.Reflection.Missing.Value;

            xlApp.DisplayAlerts = false;
            // Mở file template
            xlWorkbookTpl = xlApp.Workbooks.Open(TemplateFilePath, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorksheetTpl = xlWorkbookTpl.Worksheets[1] as Excel.Worksheet;

            ////New excel file for export
            xlWorkbook = xlApp.Workbooks.Add(misValue);
            xlWorkbook.SaveAs(SaveFilePath);

            List<string> tmpSheetName = new List<string>();

            TitleRow = DEFAULT_TITLE_ROW;
            HeaderRow = DEFAULT_HEADER_ROW;
            DataRow = HeaderRow + 1;

            int start_col = 1;
            StringBuilder sb = new StringBuilder();

            for (int k = 0; k < ToExcelScheduleList.Count; k++)
            {
                ScheduleItem item = ToExcelScheduleList[k];
                ViewSchedule OriVS = _Doc.GetElement(new ElementId(item.Id)) as ViewSchedule;

                ViewSchedule vs = OriVS;
                bool isDeleteTmp = false;
                //if (chkExportElementId.Checked)
                if (true)
                {
                    bool canDuplicate = false;
                    // Tạo ra
                    ViewSchedule tmpvs = CreateTemporarySchedule(OriVS, ELEMENT_UNIQUE_ID_PARAM_NAME, ref canDuplicate);
                    if (tmpvs != null)
                    {
                        vs = tmpvs;
                        isDeleteTmp = true;
                    }
                    else
                    {
                        //if (!canDuplicate)
                        //{
                        //    appendTextWithColor(vs.Name + ": ", System.Drawing.Color.Black);
                        //    appendTextWithColor("[WARNING]" + Environment.NewLine, System.Drawing.Color.OrangeRed);
                        //    appendTextWithColor("Cannot duplicate Schedule to create the Element_Unique_Id parameter" + Environment.NewLine, System.Drawing.Color.OrangeRed);
                        //}

                    }
                }

                TableData td = vs.GetTableData();
                TableSectionData header = td.GetSectionData(SectionType.Header);

                TableSectionData tsd = td.GetSectionData(SectionType.Body);



                string sheetName = "";

                try
                {
                    if (tsd.NumberOfColumns > 0)
                    {
                        xlWorksheetTpl.Copy(misValue, xlWorkbook.Worksheets[xlWorkbook.Worksheets.Count] as Excel.Worksheet);
                        xlWorkbook.Save();

                        Excel.Worksheet xlWorkSheetTmp = xlWorkbook.Worksheets[xlWorkbook.Worksheets.Count];// as Excel.Worksheet;

                        if (OriVS.Name.Length < 32)
                            sheetName = checkSheetNameExist(tmpSheetName, OriVS.Name);
                        else
                            sheetName = checkSheetNameExist(tmpSheetName, OriVS.Name.Substring(0, 30));

                        xlWorkSheetTmp.Name = sheetName;

                        tmpSheetName.Add(xlWorkSheetTmp.Name);

                        //int title_row = start_row;
                        //xlWorkSheetTmp.Cells[title_row, 1] = vs.Name;

                        int start_row_data = tsd.FirstRowNumber;
                        int offset_row = 0;
                        if (header.HideSection == false)
                        {
                            xlWorkSheetTmp.Cells[TitleRow, 1] = OriVS.Name;
                        }
                        else
                        {
                            xlWorkSheetTmp.Cells[TitleRow, 1] = vs.GetCellText(SectionType.Body, 0, 0);
                            //xlWorkSheetTmp.Cells[TitleRow, 1] = tsd.GetCellText(0, 0);
                            start_row_data = tsd.FirstRowNumber + 1;
                            offset_row = 1;
                        }



                        int specCol = 0;
                        int finalCol = tsd.NumberOfColumns;
                        for (int i = start_row_data; i < tsd.NumberOfRows; i++)
                        {
                            for (int j = tsd.FirstColumnNumber; j < tsd.NumberOfColumns; j++)
                            {
                                //string temp = tsd.GetCellText(i, j);
                                string temp = vs.GetCellText(SectionType.Body, i, j);
                                if (i == start_row_data && j == tsd.NumberOfColumns - 1)
                                {
                                    temp = ELEMENT_UNIQUE_ID_PARAM_NAME;
                                }

                                int row = i + HeaderRow - offset_row;
                                int col = j + start_col;
                                xlWorkSheetTmp.Cells[row, col] = temp;
                                Excel.Range oRange = xlWorkSheetTmp.Cells[row, col];
                                string link = ScheduleToExcelCommand.path + @"\" + temp;
                                //bool check = cbIsMap.Checked && File.Exists(link);
                                //bool extenImage = temp.EndsWith(".jpg") || temp.EndsWith(".png") || temp.EndsWith(".jpeg") || temp.EndsWith(".tiff") || temp.EndsWith(".gif") || temp.EndsWith(".bmp");
                                //if (extenImage && check)
                                //{
                                //    Image image = Image.FromFile(link);
                                //    float width = image.Width;
                                //    float height = image.Height;
                                //float Left = (float)((double)oRange.Left) + 1;
                                //float Top = (float)((double)oRange.Top) + 2;
                                //xlWorkSheetTmp.Rows[row].RowHeight = 75;
                                //xlWorkSheetTmp.Columns[col].ColumnWidth = 20;
                                //xlWorkSheetTmp.Shapes.AddPicture(link, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, (float)144 / (float)(96.0 / 72.0), (float)96 / (float)(96.0 / 72.0));

                                //    specCol = col;
                                //}
                            }
                        }

                        if (specCol > 0)
                        {
                            xlWorkSheetTmp.Columns[specCol].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }

                        //Title Range
                        string cell1 = this.getCellNameByLetter(1, TitleRow);
                        string cell2 = this.getCellNameByLetter(tsd.NumberOfColumns, TitleRow);
                        Excel.Range rangeTitle = xlWorkSheetTmp.get_Range(cell1, cell2);
                        rangeTitle.Font.Bold = true;
                        rangeTitle.Merge();
                        rangeTitle.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        rangeTitle.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        rangeTitle.Font.Size = 12;
                        //rangeTitle.Font.Color = "Red";


                        //rangeTitle.Interior.Color = "Green";
                        rangeTitle.RowHeight = 40;

                        //Header Range
                        Excel.Range rangeHeader = xlWorkSheetTmp.get_Range("A" + HeaderRow, "Z" + HeaderRow);
                        rangeHeader.Font.Bold = true;
                        if (specCol == 0)
                        {
                            rangeHeader.EntireColumn.AutoFit();
                        }
                        else if (specCol != 1)
                        {
                            string C1 = getCellNameByLetter(1, HeaderRow);
                            string C2 = getCellNameByLetter(specCol - 1, HeaderRow);
                            Excel.Range range_1 = xlWorkSheetTmp.get_Range(C1, C2);
                            range_1.EntireColumn.AutoFit();
                            string C3 = getCellNameByLetter(specCol + 1, HeaderRow);
                            string C4 = getCellNameByLetter(finalCol, HeaderRow);
                            Excel.Range range_2 = xlWorkSheetTmp.get_Range(C3, C4);
                            range_2.EntireColumn.AutoFit();
                        }
                        else
                        {
                            string C3 = getCellNameByLetter(specCol + 1, HeaderRow);
                            string C4 = getCellNameByLetter(finalCol, HeaderRow);
                            Excel.Range range_2 = xlWorkSheetTmp.get_Range(C3, C4);
                            range_2.EntireColumn.AutoFit();
                        }

                        //Scale All Image to Original Size After The Colunm is AutoFit
                        //Excel.Pictures pics = xlWorkSheetTmp.Pictures(misValue) as Excel.Pictures;
                        //scalePicturesToOriginal(pics);
                        //this.scalePictureToOriginal(pics, 1);

                        //Draw Border
                        string cell11 = getCellNameByLetter(1, HeaderRow);
                        string cell22 = getCellNameByLetter(tsd.NumberOfColumns, HeaderRow + tsd.NumberOfRows - 1);
                        Excel.Range rangeBorder = xlWorkSheetTmp.get_Range(cell11, cell22);

                        rangeBorder.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        // Đang không gán màu được, gán màu sẽ bị lỗi và hiển thị luôn excel
                        //rangeBorder.Borders.Color = "Blue";

                        //xlWorkSheetTmp.Copy(misValue, xlWorkBook.Worksheets[xlWorkBook.Worksheets.Count] as Excel.Worksheet);
                        xlWorkbook.Save();

                        //appendTextWithColor(sheetName + ": ", System.Drawing.Color.Black);
                        //appendTextWithColor("[DONE]" + Environment.NewLine, System.Drawing.Color.Green);
                    }
                    else
                    {
                        //appendTextWithColor(sheetName + ": ", System.Drawing.Color.Black);
                        //appendTextWithColor("[EMPTY DATA]" + Environment.NewLine, System.Drawing.Color.OrangeRed);
                    }

                    //Update Progress Bar


                }
                catch (Exception ex)
                {
                    //this.appendTextWithColor(vs.Name + ": ", System.Drawing.Color.Black);
                    //this.appendTextWithColor("[FAIL]" + Environment.NewLine, System.Drawing.Color.Red);
                    //this.appendTextWithColor(ex.ToString() + Environment.NewLine, System.Drawing.Color.Red);
                }

                //sb.AppendLine("=======================");

                if (isDeleteTmp)
                {
                    Transaction trans = new Transaction(_Doc, "DELETE-TEMP-SCHEDULE");
                    trans.Start();
                    _Doc.Delete(vs.Id);
                    trans.Commit();
                }

            }

            try
            {
                if (xlWorkbook.Worksheets.Count > 1)
                {
                    (xlWorkbook.Worksheets[1] as Excel.Worksheet).Delete();
                    xlWorkbook.Save();
                }
            }
            catch (Exception ex1)
            {
                //this.appendTextWithColor(ex1.ToString() + Environment.NewLine, System.Drawing.Color.Red);
            }
            finally
            {
                xlWorkbook.Close(true, misValue, misValue);
                xlWorkbookTpl.Close(false, misValue, misValue);
                xlApp.Quit();

                this.releaseObject(xlWorksheetTpl);
                this.releaseObject(xlWorkbook);
                this.releaseObject(xlWorkbookTpl);
                this.releaseObject(xlApp);
            }
            //xlWorkbook.SaveAs(assemplyDirPath + @"\excel_tpl\file_01.xlsx");

            //TaskDialog.Show(TAG, sb.ToString());

            //if (chkExportFolder.Checked)
            //{
            //    string dir = System.IO.Path.GetDirectoryName(this.txtSavePath.Text);
            //    if (System.IO.Directory.Exists(dir))
            //        System.Diagnostics.Process.Start("explorer.exe", dir);
            //}


            //if (chkOpenExportFile.Checked)
            //{
            if (System.IO.File.Exists(SaveFilePath))
                System.Diagnostics.Process.Start(SaveFilePath);
            //}
        }

        private ViewSchedule CreateTemporarySchedule(ViewSchedule pOriSchedule, string pUniqueIDParamName, ref bool pCanDuplicate)
        {
            Transaction trans = new Transaction(_Doc, "CREATE-UNIQUE-PARAM-FIELD");
            trans.Start();

            ViewSchedule vs = null;
            pCanDuplicate = true;
            try
            {
                //duplicate the schedule so that the original is not modified
                ElementId tempViewScheduleId = pOriSchedule.Duplicate(ViewDuplicateOption.Duplicate);
                vs = _Doc.GetElement(tempViewScheduleId) as ViewSchedule;
            }
            catch (Exception ex)
            {
                trans.RollBack();
                pCanDuplicate = false;
                return null;
            }


            Guid paramGuid = Guid.Empty;

            FilteredElementCollector eleCollector = new FilteredElementCollector(this._Doc, vs.Id);
            eleCollector.WhereElementIsNotElementType();
            eleCollector.OfCategory((BuiltInCategory)vs.Definition.CategoryId.IntegerValue);

            List<Element> eleList = eleCollector.ToList<Element>();

            if (eleList == null || (eleList != null && eleList.Count == 0))
            {
                trans.RollBack();
                return null;
            }


            Element firstEle = eleList[0];
            Util.ParameterUtil.createOrUpdateCustomParameter(_App, _Doc, firstEle, pUniqueIDParamName, "Element Id", ParameterType.Text, BuiltInParameterGroup.PG_IDENTITY_DATA);

            //List<ElementId> listExcludedIds = new List<ElementId>();
            Parameter paramElementId = firstEle.LookupParameter(pUniqueIDParamName);
            if (paramElementId != null)
            {
                paramGuid = paramElementId.GUID;
                //listExcludedIds.Add(paramElementId.Id);
            }

            foreach (Element ele in eleCollector)
            {
                if (BuiltInCategory.OST_Views == (BuiltInCategory)ele.Category.Id.IntegerValue)
                {
                    Autodesk.Revit.DB.View tmpView = ele as Autodesk.Revit.DB.View;
                    if (tmpView.IsTemplate)
                        continue;

                    Autodesk.Revit.DB.View tplView = _Doc.GetElement(tmpView.ViewTemplateId) as Autodesk.Revit.DB.View;
                    if (tplView != null)
                    {
                        //tplView.SetNonControlledTemplateParameterIds(listExcludedIds);

                        ICollection<ElementId> nonIds = tplView.GetNonControlledTemplateParameterIds();
                        if (!nonIds.Contains(paramElementId.Id))
                        {
                            nonIds.Add(paramElementId.Id);
                            tplView.SetNonControlledTemplateParameterIds(nonIds);
                        }
                    }

                    Parameter p = tmpView.get_Parameter(paramGuid);

                    if (p != null)
                    {
                        if (!p.IsReadOnly)
                        {
                            p.Set(ele.Id.IntegerValue.ToString());
                            //ele.LookupParameter(pUniqueIdParamName).Set(ele.Id.IntegerValue.ToString());
                        }
                    }
                }
                else
                {
                    Parameter p2 = ele.get_Parameter(paramGuid);
                    if (p2 != null)
                    {
                        if (!p2.IsReadOnly)
                        {
                            p2.Set(ele.Id.IntegerValue.ToString());
                            //ele.LookupParameter(pUniqueIdParamName).Set(ele.Id.IntegerValue.ToString());
                        }
                    }

                }

            }

            //Create new field to mapping with UniqueId
            SchedulableField schedulableField = null;
            foreach (SchedulableField sf in vs.Definition.GetSchedulableFields())
            {
                if (sf.GetName(_Doc).Equals(pUniqueIDParamName))
                {
                    schedulableField = sf;
                    break;
                }
            }

            if (schedulableField == null)
            {
                TaskDialog.Show(TAG, "Didn't find Element's IdParam as Schedulable Field");
                trans.RollBack();
                return null;
            }

            try
            {
                vs.Definition.AddField(schedulableField);
                vs.RefreshData();

            }
            catch (Exception ex)
            {
                TaskDialog.Show(TAG, ex.ToString());
                trans.RollBack();
                return null;
            }

            trans.Commit();

            return vs;
        }

        private bool isFileInUse(System.IO.FileInfo file)
        {

            bool rs = false;
            System.IO.FileStream fs = null;
            try
            {
                fs = System.IO.File.OpenWrite(file.FullName);
            }
            catch (System.IO.IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                rs = true;
            }
            finally
            {
                if (fs != null)
                    fs.Close();
            }

            return rs;
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                //MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        private string checkSheetNameExist(List<string> pSrc, string pName)
        {
            pName.Replace(':', '_').Replace('/', '_');

            if (pSrc.Count == 0)
                return pName;

            string temp = pName;
            bool b = false;
            int k = 1;
            do
            {
                b = false;
                foreach (string str in pSrc)
                {
                    if (str.Equals(temp, StringComparison.OrdinalIgnoreCase))
                    {
                        temp = (k++) + "_" + pName;
                        temp = temp.Substring(0, 31);
                        b = true;
                        break;
                    }
                }
            } while (b);

            return temp;
        }

        private string getCellNameByLetter(int col, int row)
        {
            string[] letters = new string[]{
                "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L",
                "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
                "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL",
                "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ"
            };

            return letters[col - 1] + row;

        }

        private void scalePicturesToOriginal(Excel.Pictures pics)
        {
            for (int i = 1; i <= pics.Count; i++)
            {
                Excel.Picture pic = pics.Item(i) as Excel.Picture;
                pic.ShapeRange.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoCTrue;
                pic.ShapeRange.ScaleWidth(1, Microsoft.Office.Core.MsoTriState.msoTrue);
                pic.ShapeRange.ScaleHeight(1, Microsoft.Office.Core.MsoTriState.msoTrue);
            }

        }

        private void scalePictureToOriginal(Excel.Pictures pics, int pos)
        {
            if (pics == null || pics.Count == 0 || pos < 1)
                return;

            Excel.Picture pic = pics.Item(pos) as Excel.Picture;
            if (pic == null)
                return;
            pic.ShapeRange.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoCTrue;
            pic.ShapeRange.ScaleWidth(1, Microsoft.Office.Core.MsoTriState.msoTrue);
            pic.ShapeRange.ScaleHeight(1, Microsoft.Office.Core.MsoTriState.msoTrue);


        }

        private void appendTextWithColor(string text, System.Drawing.Color color)
        {
            //this.rtbLog.Select(this.rtbLog.TextLength, 0);
            //this.rtbLog.SelectionColor = color;
            //this.rtbLog.AppendText(text);
        }

    }

}
