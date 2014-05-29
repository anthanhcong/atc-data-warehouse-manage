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

namespace WarehouseManager
{
    public partial class User_Login
    {
        private string User_Database_Conn = @"server=ADMIN\SQLEXPRESS;database=USER_DB;uid=sa;pwd=123456";

        private DataTable Get_SQL_Data(string connString, string cmd_str, ref SqlDataAdapter dataAdapter, ref DataSet input_dataset)
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

        private DataTable Get_UserInfo(string username, ref SqlDataAdapter dataAdapter, ref DataSet input_dataset)
        {
            DataTable dtbTmp = new DataTable();
            string get_user_info_cmd = @"SELECT * FROM USER_DASHBOARD_tb WHERE username = '" + username + "'";
            dtbTmp = Get_SQL_Data(User_Database_Conn, get_user_info_cmd, ref dataAdapter, ref input_dataset);
            return dtbTmp;
        }

        private string Map_UserName(string loginname)
        {
            string alias_name, user_name;
            DataTable dtbTmp = new DataTable();
            DataSet datasettmp = new DataSet();
            SqlDataAdapter adaptertmp = new SqlDataAdapter();
            string get_user_info_cmd = @"SELECT * FROM USER_DASHBOARD_tb";//User_AliasName_TBL";
            dtbTmp = Get_SQL_Data(User_Database_Conn, get_user_info_cmd, ref adaptertmp, ref datasettmp);


            foreach (DataRow row in dtbTmp.Rows)
            {
                alias_name = row["AliasName"].ToString().Trim();
                user_name = row["UserName"].ToString().Trim();
                if (alias_name == loginname) return user_name;
            }
            return loginname;
        }

    }
}