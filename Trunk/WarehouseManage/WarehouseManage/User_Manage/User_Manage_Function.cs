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
using System.Diagnostics;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;


namespace WarehouseManager
{
    public partial class Form1
    {
        #region table User Manage contend

        int UserName_col;
        int Password_col;
        int MSNV_col;
        int AliasName_col;
        int PermissionID_col;
        int BomManage_col;
        int ImportMaterial_col;

        const string UserName_col_str = "UserName";
        const string Password_col_str = "Password";
        const string MSNV_col_str = "MSNV";
        const string AliasName_col_str = "Alias Name";
        const string PermissionID_col_str = "Permission ID";
        const string BomManage_col_str = "Allow Bom Manage";
        const string ImportMaterial_col_str = "Allow Import Material";

        #endregion

        private bool Import_User_Manage_file(string file_name)
        {
            int row;
            bool ret_val = false;
            string cell_value, username;

            User_Table_Form.Load_DataBase(User_Conn, @"SELECT * FROM dbo.USER_DASHBOARD_tb");
            Permission_Table_Form.Load_DataBase(User_Conn, @"SELECT * FROM dbo.PERMISSION_tb");

            row = 1;
            ProgressBar1.Visible = true;
            StatusLabel.Text = "Loading File";
            row = 3;
            OpenWB = Open_excel_file(file_name, "");
            if (User_Manage_Get_Col_info(OpenWB) == true)
            {
                StatusLabel.Text = "Process File";
                cell_value = Get_Excel_Line((Excel.Worksheet)OpenWB.Sheets[1], row, UserName_col, 1);
                while (cell_value.Trim() != "")
                {
                    username = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, UserName_col, 20);
                    if (username != "")
                    {
                        ret_val = Add_User_Item_Row(username, row);
                    }
                    row++;
                    cell_value = Get_Excel_Line((Excel.Worksheet)OpenWB.Sheets[1], row, UserName_col, 1);
                    ProgressBar1.Value = row % 100;
                }
            }
            else
            {
                MessageBox.Show(Error_log, "Error");
                ret_val = false;
            }
            Close_WorkBook(OpenWB);
            Clear_Stt_Column_HQ_Item();
            ProgressBar1.Value = 0;
            StatusLabel.Text = "Status";
            ProgressBar1.Visible = false;
            return ret_val;
        }

        private bool Add_User_Item_Row(string username, int row)
        {
            bool ret_var = false;
            string permission_id, bom_manage, import_material;
            string filterExpression = "";
            filterExpression = "UserName =" + "'" + username + "'";
            string filter_per = "";
            DataRow[] rows = User_Table_Form.Data_dtb.Select(filterExpression);

            if (rows.Length == 0)
            {
                DataRow new_row = User_Table_Form.Data_dtb.NewRow();
                try
                {
                    new_row["UserName"] = username.Trim();
                    new_row["Password"] = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, Password_col, 20);
                    new_row["MSNV"] = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, MSNV_col, 20);
                    new_row["AliasName"] = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, AliasName_col, 30);
                    permission_id = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, PermissionID_col, 10);
                    new_row["PermissionID"] = permission_id;
                    filter_per = "PermissionID =" + "'" + permission_id + "'";
                    DataRow[] rows1 = Permission_Table_Form.Data_dtb.Select(filter_per);
                    //User_Table_Form.Data_dtb.Rows.Add(new_row);
                    if (rows1.Length == 0)
                    {
                        User_Table_Form.Data_dtb.Rows.Add(new_row);
                        DataRow new_row1 = Permission_Table_Form.Data_dtb.NewRow();
                        new_row1["PermissionID"] = permission_id.Trim();
                        bom_manage = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, BomManage_col, 10);
                        import_material = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, ImportMaterial_col, 10);
                        if ((bom_manage.Trim() == "True") || (bom_manage.Trim() == "TRUE") || (bom_manage.Trim() == "true"))
                        {
                            new_row1["BomManage"] = 1;// "True";
                        }
                        else if ((bom_manage.Trim() == "False") || (bom_manage.Trim() == "FALSE") || (bom_manage.Trim() == "false")) 
                        {
                            new_row1["BomManage"] = 0;// "False";
                        }
                        else
                        {
                            new_row1["BomManage"] = 0;// "False";
                        }
                        if ((import_material.Trim() == "True") || (import_material.Trim() == "TRUE") || (import_material.Trim() == "true"))
                        {
                            new_row1["ImportMaterial"] = 1;// "True";
                        }
                        else if ((import_material.Trim() == "False") || (import_material.Trim() == "FALSE") || (import_material.Trim() == "false")) 
                        {
                            new_row1["ImportMaterial"] = 0;// "False";
                        }
                        else
                        {
                            new_row1["ImportMaterial"] = 0;// "False";
                        }
                        Permission_Table_Form.Data_dtb.Rows.Add(new_row1);
                        ret_var = true;
                    }
                    else
                    {
                        MessageBox.Show("Permission_ID: " + permission_id + " is exist.\nPlease fill new permission_id", "Warning");
                        ((Excel.Range)((Excel.Worksheet)OpenWB.Sheets[1]).Cells[row, PermissionID_col]).Interior.Color = 250;
                    }
                }
                catch
                {
                    MessageBox.Show("Import Item fail, row: " + row, "Error");
                    ((Excel.Range)((Excel.Worksheet)OpenWB.Sheets[1]).Cells[row, 1]).Interior.Color = 250;
                }
            }
            else
            {
                string message = "Has duplicate for item: " + username
                                     + "\nDo you want to update?";
                if (MessageBox.Show(message, "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    try
                    {
                        rows[0]["UserName"] = username.Trim();
                        rows[0]["Password"] = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, Password_col, 20);
                        rows[0]["MSNV"] = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, MSNV_col, 20);
                        rows[0]["AliasName"] = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, AliasName_col, 30);
                        permission_id = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, PermissionID_col, 10);
                        //if ( permission_id.Trim() != rows[0]["PermissionID"].ToString().Trim())
                        //{
                            //rows[0]["PermissionID"] = permission_id;
                            filter_per = "PermissionID =" + "'" + rows[0]["PermissionID"].ToString().Trim() + "'";
                            DataRow[] rows1 = Permission_Table_Form.Data_dtb.Select(filter_per);
                            if (rows1.Length != 0)
                            {
                                rows1[0]["PermissionID"] = permission_id;
                                bom_manage = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, BomManage_col, 10);
                                import_material = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, ImportMaterial_col, 10);
                                if ((bom_manage.Trim() == "True") || (bom_manage.Trim() == "TRUE") || (bom_manage.Trim() == "true"))
                                {
                                    rows1[0]["BomManage"] = 1;// "True";
                                }
                                else if ((bom_manage.Trim() == "False") || (bom_manage.Trim() == "FALSE") || (bom_manage.Trim() == "false"))
                                {
                                    rows1[0]["BomManage"] = 0;// "False";
                                }
                                else
                                {
                                    rows1[0]["BomManage"] = 0;// "False";
                                }
                                if ((import_material.Trim() == "True") || (import_material.Trim() == "TRUE") || (import_material.Trim() == "true"))
                                {
                                    rows1[0]["ImportMaterial"] = 1;// "True";
                                }
                                else if ((import_material.Trim() == "False") || (import_material.Trim() == "FALSE") || (import_material.Trim() == "false"))
                                {
                                    rows1[0]["ImportMaterial"] = 0;// "False";
                                }
                                else
                                {
                                    rows1[0]["ImportMaterial"] = 0;// "False";
                                }
                            }
                            rows[0]["PermissionID"] = permission_id;
                        //}
                    }
                    catch
                    {
                        MessageBox.Show("Import Item fail, row: " + row, "Error");
                        ((Excel.Range)((Excel.Worksheet)OpenWB.Sheets[1]).Cells[row, 1]).Interior.Color = 250;
                    }
                }
                else
                {
                    //ret_var = false;
                }
            }
            return ret_var;
        }

        private bool User_Manage_Get_Col_info(Excel.Workbook cur_wbook)
        {
            int row, col;
            string cell_val;
            string error_log = "";
            bool error = false;

            // Find Cell need to check
            row = 2;
            ProgressBar1.Visible = true;
            StatusLabel.Text = "Check Column";
            for (col = 1; col < 50; col++)
            {
                cell_val = Get_Excel_Line((Excel.Worksheet)OpenWB.Sheets[1], row, col, 1);
                if (cell_val == UserName_col_str) UserName_col = col;
                if (cell_val == Password_col_str) Password_col = col;
                if (cell_val == MSNV_col_str) MSNV_col = col;
                if (cell_val == AliasName_col_str) AliasName_col = col;
                if (cell_val == PermissionID_col_str) PermissionID_col = col;
                if (cell_val == BomManage_col_str) BomManage_col = col;
                if (cell_val == ImportMaterial_col_str) ImportMaterial_col = col;

            }
            if (UserName_col == 0)
            {
                error_log += "\nCan not find 'UserName' Colum.\nImport Fail";
                error = true;
            }
            if (Password_col == 0)
            {
                error_log += "\nCan not find 'Mã Password' Colum.\nImport Fail";
                error = true;
            }
            if (MSNV_col == 0)
            {
                error_log += "\nCan not find 'MSNV' Colum.\nImport Fail";
                error = true;
            }
            if (AliasName_col == 0)
            {
                error_log += "\nCan not find 'AliasName' Colum.\nImport Fail";
                error = true;
            }
            if (PermissionID_col == 0)
            {
                error_log += "\nCan not find 'Permission' Colum.\nImport Fail";
                error = true;
            }
            if (BomManage_col == 0)
            {
                error_log += "\nCan not find 'Allow Bom Manage' Colum.\nImport Fail";
                error = true;
            }
            if (ImportMaterial_col == 0)
            {
                error_log += "\nCan not find 'Allow Import Material' Colum.\nImport Fail";
                error = true;
            }
            if (error == true)
            {
                Error_log = error_log;
                return false;
            }
            else
            {
                Error_log = "";
                return true;
            }
        }



        private void Clear_Stt_Column_HQ_Item()
        {
            UserName_col = 0;
            Password_col = 0;
            MSNV_col = 0;
            AliasName_col = 0;
            PermissionID_col = 0;
            BomManage_col = 0;
            ImportMaterial_col = 0;
        }

        void User_Table_Form_CellDoubleClick(Object sender, EventArgs e)
        {
            string permission;

            //if (Permission_Table_Form.Data_dtb != null)
            //{
            //    Permission_Table_Form.Data_dtb.Clear();
            //}
            if (User_Table_Form.dataGridView_View.CurrentCell == User_Table_Form.dataGridView_View.CurrentRow.Cells["Username"])
            {
                permission = User_Table_Form.dataGridView_View.CurrentRow.Cells["PermissionID"].Value.ToString().Trim();
                Permission_Table_Form.Load_DataBase(User_Conn, @"SELECT * FROM dbo.PERMISSION_tb where PermissionID = '" + permission + "'");
            }
        }

        private string Encrypt_Pass(string pass)
        {
            return StringToHexString(pass.Trim());
        }

        //private string Decrypt_Pass(string encrypt_pass)
        //{
        //    return Change_HexString2String(encrypt_pass.Trim());

        //}

        /// <summary> Converts an array of bytes into a formatted string of hex digits (ex: E4 CA B2)</summary>
        /// <param name="data"> The array of bytes to be translated into a string of hex digits. </param>
        /// <returns> Returns a well formatted string of hex digits with spacing. </returns>
        private string StringToHexString(String data)
        {
            StringBuilder sb = new StringBuilder(data.Length * 3);
            Random random = new Random();
            int addnew;

            foreach (char b in data)
            {
                sb.Append(Convert.ToString(Convert.ToByte(b), 16));
                addnew = random.Next('0', '9');
                sb.Append(addnew - '0');
            }
            return sb.ToString().ToUpper();
        }
        //private string Change_HexString2String(string indata)
        //{
        //    int i, in_len;
        //    string char_str;
        //    Int32 value;

        //    // check correct data
        //    if (indata == "") return "";
        //    in_len = indata.Length;
        //    StringBuilder sb = new StringBuilder(in_len);
        //    for (i = 0; i < in_len - 2; i = i + 3)
        //    {
        //        char_str = indata.Substring(i, 2);
        //        value = Int32.Parse(char_str, System.Globalization.NumberStyles.HexNumber);
        //        if ((value < 127) || (value > 0))
        //        {
        //            sb.Append(Convert.ToString(Convert.ToChar(Int32.Parse(char_str, System.Globalization.NumberStyles.HexNumber))));
        //        }
        //        else return "";
        //    }
        //    return sb.ToString();
        //}
    }
}