using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Threading;
using System.Globalization;
using System.IO;
using System.IO.Ports;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using ConvertDB; 

namespace WarehouseManager
{
    public partial class SQL_APPL
    {
        #region
        public Int32 Total_Emp;
        public Excel.Application OpenXL;
        public Excel.Workbook OpenWB;
        public Excel.Worksheet OpenSheet;
        public Excel.Range OpenRng;
        #endregion

        enum ADD_BILL_RESULT
        {
            ADD_FAIL = 0,
            ADD_NEW = 1,
            UPDATE = 2
        };

        /**************************************************
         * Create & Open Excel File
         * ************************************************/
        public bool create_excel_file(string file_path)
        {
            bool result = false;
            string[] file_info = file_path.Split('.');
            if (File.Exists(file_path) == false)
            {
                //khoi tao cac doi tuong Com Excel de lam viec
                Excel.Application xlApp;
                Excel.Worksheet xlSheet;
                Excel.Workbook xlBook;
                //doi tuong Trống để thêm  vào xlApp sau đó lưu lại sau
                object missValue = System.Reflection.Missing.Value;
                // khoi tao doi tuong Com Excel moi
                xlApp = new Excel.Application();
                xlBook = xlApp.Workbooks.Add(missValue);
                //su dung Sheet dau tien de thao tac
                xlSheet = (Excel.Worksheet)xlBook.Worksheets.get_Item(1);
                //không cho hiện ứng dụng Excel lên để tránh gây đơ máy
                xlApp.Visible = false;

                //save file
                if (file_info[1].Trim() == "xls")
                {
                    xlBook.SaveAs(file_path, Excel.XlFileFormat.xlWorkbookNormal, missValue, missValue, missValue, missValue, Excel.XlSaveAsAccessMode.xlExclusive, missValue, missValue, missValue, missValue, missValue);
                }
                else if (file_info[1].Trim() == "xlsx")
                {
                    xlBook.SaveAs(file_path, Excel.XlFileFormat.xlOpenXMLWorkbook, missValue, missValue, missValue, missValue, Excel.XlSaveAsAccessMode.xlExclusive, missValue, missValue, missValue, missValue, missValue);
                }
                xlBook.Close(true, missValue, missValue);
                xlApp.Quit();

                // release cac doi tuong COM
                releaseObject(xlSheet);
                releaseObject(xlBook);
                releaseObject(xlApp);
                result = true;
            }
            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="file_path"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public Excel.Workbook Open_excel_file(string file_path, string password)
        {
            bool next, file_opened;
            int cont;
            Excel.Workbook oWB;
            try
            {
                //Start Excel and get Application object.
                //OpenXL = new Excel.Application();
                //OpenXL.Visible = false;

                // Open New WordBook
                next = false;
                cont = 0;
                file_opened = false;
                do
                {
                    oWB = (Excel.Workbook)OpenXL.Workbooks.Open(file_path, 2, false, 5, password, password,
                                            false, Excel.XlPlatform.xlWindows, "", true, true, 0, true, false, false);
                    file_opened = true;
                    if (oWB.ReadOnly == true)
                    {
                        oWB.Close(false, false, false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oWB);
                        //OpenXL.Quit();
                        cont++;
                        file_opened = false;
                        if (cont == 5)
                        {
                            next = true;
                        }
                        else
                        {
                            next = false;
                        }
                        Random random = new Random();
                        int randomNumber = random.Next(0, 10);
                        Thread.Sleep(randomNumber * 100);
                    }
                    else
                    {
                        next = true;
                    }
                } while (next == false);

                if (file_opened != true)
                {
                    MessageBox.Show("Can not open excel File.\n Please try again later!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //OpenXL.Quit();
                    return null;
                }
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);
                MessageBox.Show(errorMessage, "Error");
                oWB = null;
            }
            return oWB;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oWB"></param>
        public void Close_WorkBook(Excel.Workbook oWB)
        {
            try
            {
                //Manipulate a variable number of columns for Quarterly Sales Data.
                if (oWB != null)
                {
                    oWB.DoNotPromptForConvert = true;
                    oWB.CheckCompatibility = false;
                    oWB.Save();
                    oWB.Close(false, false, false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oWB);
                }

                //OpenXL.Quit();
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenXL);
                //Thread.Sleep(1000);
                //Application.DoEvents();
            }
            catch (Exception ex)
            {
                // Bắt lỗi
                MessageBox.Show(ex.Message, "Error");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="obj"></param>
        static public void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                throw new Exception("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        //End Create & Open Excel File


        /************************************************************
         * Read File
         * **********************************************************/
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <param name="first_col"></param>
        /// <param name="num_col"></param>
        /// <returns></returns>
        public string Get_Excel_Line(Excel.Worksheet sheet, int row, int first_col, int num_col)
        {
            Excel.Range cur_cell;
            string ret_str = "", cel_str;
            int col;
            ConvertFont convert = new ConvertFont();

            for (col = first_col; col < num_col + first_col; col++)
            {
                if (col != first_col) ret_str += ";";
                cur_cell = (Excel.Range)sheet.Cells[row, col];
                cel_str = cur_cell.Text.ToString().Trim();
                convert.Convert(ref cel_str, ConvertDB.FontIndex.iNotKnown, ConvertDB.FontIndex.iUNI);
                ret_str += cel_str;
                if ((col == first_col) && (cel_str == "")) return "";
            }
            return ret_str.Trim();
        }

        public string Get_Text_Cell(Excel.Worksheet sheet, int row, int col, int max_length)
        {
            Excel.Range cur_cell;
            string cel_str;
            int len;
            ConvertFont convert = new ConvertFont();

            ((Excel.Range)sheet.Cells[row, col]).NumberFormat = "@";
            cur_cell = (Excel.Range)sheet.Cells[row, col];
            cel_str = (cur_cell.Text.ToString().Trim());
            convert.Convert(ref cel_str, ConvertDB.FontIndex.iTCV, ConvertDB.FontIndex.iUNI);
            if (cel_str == "#REF")
            {
                cel_str = "";
                cur_cell.Interior.Color = 5296274;
            }
            else
            {
                len = cel_str.Length;
                if (len > max_length)
                {
                    MessageBox.Show("Length of String too long", "Warning");
                    cur_cell.Interior.Color = 255;
                    cel_str = cel_str.Substring(0, max_length);
                    OpenWB.Save();
                }
                else
                {
                    cur_cell.Interior.Color = 5296274;
                }
            }
            return cel_str;
        }

        public string Get_date_str_Cell(Excel.Worksheet sheet, int row, int col)
        {
            Excel.Range cur_cell;
            string cel_str;
            ((Excel.Range)sheet.Cells[row, col]).NumberFormat = "[$-409]d-mmm-yyyy;@";
            cur_cell = (Excel.Range)sheet.Cells[row, col];
            cel_str = cur_cell.Text.ToString().Trim();
            cur_cell.Interior.Color = 5296274;
            return cel_str;
        }

        public int Get_int_Cell(Excel.Worksheet sheet, int row, int col)
        {
            Excel.Range cur_cell;
            string cel_str;
            int value;
            try
            {
                ((Excel.Range)sheet.Cells[row, col]).NumberFormat = "0";
                ((Excel.Range)sheet.Cells[row, col]).EntireColumn.AutoFit();
                cur_cell = (Excel.Range)sheet.Cells[row, col];
                cel_str = cur_cell.Text.ToString().Trim();
                if (cel_str == "") value = 0;
                else value = Convert.ToInt32(cel_str);
                cur_cell.Interior.Color = 5296274;
            }
            catch
            {
                MessageBox.Show("Error when get Int value in cell (row = " + row + "; col = " + col + ")");
                ((Excel.Range)sheet.Cells[row, col]).Interior.Color = 255;
                value = 0;
                OpenWB.Save();
            }

            return value;
        }

        public float Get_float_Cell(Excel.Worksheet sheet, int row, int col)
        {
            Excel.Range cur_cell;
            string cel_str;
            float value;
            try
            {
                ((Excel.Range)sheet.Cells[row, col]).NumberFormat = "0.0000";
                ((Excel.Range)sheet.Cells[row, col]).EntireColumn.AutoFit();
                cur_cell = (Excel.Range)sheet.Cells[row, col];
                cel_str = cur_cell.Text.ToString().Trim();
                if (cel_str == "") value = 0;
                else value = float.Parse(cel_str);
                cur_cell.Interior.Color = 5296274;
            }
            catch
            {
                MessageBox.Show("Error when get Float value in cell (row = " + row + "; col = " + col + ")");
                ((Excel.Range)sheet.Cells[row, col]).Interior.Color = 255;
                value = 0;
                OpenWB.Save();
            }

            return value;
        }

        public int Convert_StringNumber2Int(string number)
        {
            string value_num;
            int value;
            NumberStyles styles = NumberStyles.Number;

            //String[] strVals2;
            String[] strVals = number.Split(')');
            if (strVals.Length == 2)
            {
                String[] strVals2 = strVals[0].Split('(');
                value_num = "-" + strVals2[1];
                value = int.Parse(value_num.Trim(), styles);
            }
            else
            {
                if (number != "-")
                {
                    value = int.Parse(number.Trim(), styles);
                }
                else
                {
                    value = 0;
                }
            }

            return value;

        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="file_name"></param>
        /// <param name="fInfo"></param>
        /// <param name="table"></param>
        /// <param name="filter_str"></param>
        /// <param name="priKey_index"></param>
        /// <returns></returns>
        public bool ReadDataFromFile(string file_name, string fInfo, int first_row, DataTable table, string filter_str, int priKey_index)
        {
            String strLine = String.Empty;
            StreamReader myfile;
            int num_col;
            int row;


            // ProgressBar1.Visible = true;
            if ((fInfo == ".xls") || (fInfo == ".XLS") || (fInfo == ".xlsx") || (fInfo == ".XLSX"))
            {
                OpenWB = Open_excel_file(file_name, "");
                row = first_row;
                num_col = table.Columns.Count;
                while ((strLine = Get_Excel_Line((Excel.Worksheet)OpenWB.Sheets[1], row, 1, num_col)) != "")
                {
                    try
                    {
                        AddDataRowToTable(strLine, table, filter_str, priKey_index, ';');
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error At Line: " + row + ". With error message\n" + ex.ToString(), "Error");
                    }
                    // ProgressBar1.Value = row % 100;
                    row++;
                }
                Close_WorkBook(OpenWB);
            }
            else if ((fInfo == ".csv") || (fInfo == ".CSV"))
            {
                myfile = File.OpenText(file_name);
                if ((strLine = myfile.ReadLine()) != null)
                {
                    row = 0;
                    while (strLine != null)
                    {
                        try
                        {

                            AddDataRowToTable(strLine, table, filter_str, priKey_index, ',');
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error At Line: " + row + ". With error message\n" + ex.ToString(), "Error");
                        }
                        // ProgressBar1.Value = row % 100;
                        row++;
                        strLine = myfile.ReadLine();
                    }
                }
                myfile.Close();
            }

            // ProgressBar1.Visible = false;
            return true;
        }

        // End Read File
        
        /***************************************************
         * Import data from excel to Datatable
         * *************************************************/

        public bool AddDataRowToTable(String strCSVLine, DataTable dt, string filter_str, int priKey_index, char split_char)
        {
            int idx = 0;
            string priKey_var;
            string filterExpression = "";
            String[] strVals = strCSVLine.Split(split_char);
            Int32 iTotalNumberOfValues = strVals.Length;
            DataRow drow = dt.NewRow();
            idx = 0;

            if (strVals.Length > 0)
            {
                // Get Empl_ID
                if (strVals[priKey_index].Trim() != "")
                {
                    priKey_var = strVals[priKey_index].Trim();
                    // Check data in Table
                    filterExpression = filter_str + "'" + priKey_var + "'";
                    DataRow[] rows = dt.Select(filterExpression);

                    if (rows.Length == 1)
                    {
                        // @NOTE (Kien #1#): update current row
                        foreach (String strVal in strVals)
                        {
                            if (idx != 0)
                            {
                                if (strVal != "")
                                    rows[0][idx] = strVal.Trim();
                            }
                            idx++;
                        }
                    }
                    else if (rows.Length == 0)
                    {
                        // @NOTE (Kien #1#): insert new row
                        foreach (String strVal in strVals)
                        {
                            if (idx == priKey_index)
                            {
                                drow[idx] = priKey_var.Trim();
                            }
                            else
                            {
                                if (strVal != "")
                                {
                                    drow[idx] = DateTime.ParseExact((strVal.Trim()), "dd/MM/yy", System.Globalization.CultureInfo.CurrentCulture);
                                }
                            }
                            idx++;
                        }
                        dt.Rows.Add(drow);
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            return true;
        }
      

        /******************************************************
         * Export DataTable to Excel File
         * ****************************************************/
        /// <summary>
        /// 
        /// </summary>
        /// <param name="file_path"></param>
        /// <param name="tieude"></param>
        /// <param name="dt"></param>
        /// <returns></returns>
        public bool ExportDataToExcel(string file_path, string tieude, DataTable dt)
        {
            bool result = false;
            //khoi tao cac doi tuong Com Excel de lam viec
            Excel.Application xlApp;
            Excel.Worksheet xlSheet;
            Excel.Workbook xlBook;
            //doi tuong Trống để thêm  vào xlApp sau đó lưu lại sau
            object missValue = System.Reflection.Missing.Value;
            //khoi tao doi tuong Com Excel moi
            xlApp = new Excel.Application();
            xlBook = xlApp.Workbooks.Add(missValue);
            //su dung Sheet dau tien de thao tac
            xlSheet = (Excel.Worksheet)xlBook.Worksheets.get_Item(1);
            //không cho hiện ứng dụng Excel lên để tránh gây đơ máy
            xlApp.Visible = false;
            int socot = dt.Columns.Count;
            int sohang = dt.Rows.Count;
            int i, j;

            if (file_path != "")
            {
                //set thuoc tinh cho tieu de
                xlSheet.get_Range("A1", Convert.ToChar(socot + 65) + "1").Merge(false);
                Excel.Range caption = xlSheet.get_Range("A1", Convert.ToChar(socot + 65) + "1");
                caption.Select();
                caption.FormulaR1C1 = tieude;
                //căn lề cho tiêu đề
                caption.HorizontalAlignment = Excel.Constants.xlCenter;
                caption.Font.Bold = true;
                caption.VerticalAlignment = Excel.Constants.xlCenter;
                caption.Font.Size = 15;
                //màu nền cho tiêu đề
                caption.Interior.ColorIndex = 20;
                caption.RowHeight = 30;
                //set thuoc tinh cho cac header
                Excel.Range header = xlSheet.get_Range("A2", Convert.ToChar(socot + 65) + "2");
                header.Select();

                header.HorizontalAlignment = Excel.Constants.xlCenter;
                header.Font.Bold = true;
                header.Font.Size = 10;
                //điền tiêu đề cho các cột trong file excel
                for (i = 0; i < socot; i++)
                {
                    xlSheet.Cells[2, i + 2] = dt.Columns[i].ColumnName;
                }
                //dien cot stt
                xlSheet.Cells[2, 1] = "STT";
                for (i = 0; i < sohang; i++)
                {
                    xlSheet.Cells[i + 3, 1] = i + 1;
                }

                //dien du lieu vao sheet
                // ProgressBar1.Visible = true;
                for (i = 0; i < sohang; i++)
                {
                    for (j = 0; j < socot; j++)
                    {
                        ((Excel.Range)xlSheet.Cells[i + 3, j + 2]).NumberFormat = "@";
                        xlSheet.Cells[i + 3, j + 2] = dt.Rows[i][j].ToString() == "" ? dt.Rows[i][j] : dt.Rows[i][j].ToString().Trim();
                    }
                    // Update progress Bar
                    // ProgressBar1.Value = i % 100;
                }
                // ProgressBar1.Visible = false;

                //autofit độ rộng cho các cột
                for (i = 0; i < socot; i++)
                {
                    ((Excel.Range)xlSheet.Cells[1, i + 1]).EntireColumn.AutoFit();
                }

                //save file
                xlBook.SaveAs(file_path, Excel.XlFileFormat.xlWorkbookNormal, missValue, missValue, missValue, missValue, Excel.XlSaveAsAccessMode.xlExclusive, missValue, missValue, missValue, missValue, missValue);
                xlBook.Close(true, missValue, missValue);
                xlApp.Quit();

                // release cac doi tuong COM
                releaseObject(xlSheet);
                releaseObject(xlBook);
                releaseObject(xlApp);
                result = true;
            }
            return result;
        }

        public bool Export_Data2Excel(Button bt, DataTable dt, string title)
        {
            SaveFileDialog save_diaglog = new SaveFileDialog();
            string file_name, fInfo;
            string temp;

            save_diaglog.Filter = "Excel File (*.xls)|*.xls|All files (*.*)|*.*";
            if (save_diaglog.ShowDialog() == DialogResult.OK)
            {
                file_name = save_diaglog.FileName;
                fInfo = Path.GetExtension(save_diaglog.FileName);
                temp = bt.Text;
                bt.Text = "Exporting ...";
                bt.Enabled = false;
                if ((fInfo == ".xls") || (fInfo == ".xlsx"))
                {
                    ExportDataToExcel(file_name, title, dt);
                }
                bt.Enabled = true;
                bt.Text = temp;
            }
            MessageBox.Show("Export File thành công", "Thông báo");
            return true;
        }
        //End Export DataTable to Excel File
    }
}