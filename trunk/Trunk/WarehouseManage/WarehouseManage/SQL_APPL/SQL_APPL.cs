using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.IO.Ports;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace WarehouseManager
{
    public partial class SQL_APPL: Form
    {
        public DataTable Get_SQL_Data(string connString, string cmd_str, ref SqlDataAdapter dataAdapter, ref DataSet input_dataset)
        {
            DataTable dtbTmp = new DataTable();

            System.Data.SqlClient.SqlConnection conn = new SqlConnection(connString);
            try
            {
                conn.Open();
                dataAdapter = new SqlDataAdapter(cmd_str, conn);
                dataAdapter.Fill(input_dataset);
                dtbTmp = input_dataset.Tables[0];
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error");
            }
            finally
            {
                conn.Close();
            }
            return dtbTmp;
        }

        public bool Update_SQL_Data(SqlDataAdapter dataAdapter, DataTable dtbTmp)
        {
            // Create SQL Command builder
            SqlCommandBuilder cb = new SqlCommandBuilder(dataAdapter);
            try
            {
                //dataAdapter.Fill(dtbTmp);
                cb.GetUpdateCommand();
                dataAdapter.Update(dtbTmp);
                dtbTmp.AcceptChanges();
            }
            catch (Exception ex)
            {
                // Bắt lỗi
                MessageBox.Show(ex.Message);
                return false;
            }
            return true;
        }

        public bool Update_Data_Info(string connString, string sql_cmd)
        {
            // Tạo connection
            System.Data.SqlClient.SqlConnection conn = new SqlConnection(connString);
            try
            {
                conn.Open();
                // Get data from Database
                SqlCommand update_sql = new SqlCommand(sql_cmd, conn);
                update_sql.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            return true;
        }


        public bool Check_ItemExistTable(string item, DataTable table, string col)
        {
            //foreach (DataRow row in Card_List_Each_Provider_dtb.Rows)
            foreach (DataRow row in table.Rows)
            {
                if (item.Trim() == row[col].ToString().Trim())
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="file_path"></param>
        /// <param name="tieude"></param>
        /// <param name="dt"></param>
        /// <returns></returns>

        // Temp: Lam add
        public bool ExportDataToExcel(string file_path, string tieude, DataTable dt)
        {
            bool result = false;
            //khoi tao cac doi tuong Com Excel de lam viec
            Excel.ApplicationClass xlApp;
            Excel.Worksheet xlSheet;
            Excel.Workbook xlBook;
            //doi tuong Trống để thêm  vào xlApp sau đó lưu lại sau
            object missValue = System.Reflection.Missing.Value;
            //khoi tao doi tuong Com Excel moi
            xlApp = new Excel.ApplicationClass();
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
                releaseObject_2(xlSheet);
                releaseObject_2(xlBook);
                releaseObject_2(xlApp);
                result = true;
            }
            return result;
        }

        static public void releaseObject_2(object obj)
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
    }
}
