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
