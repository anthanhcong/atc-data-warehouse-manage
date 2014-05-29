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
        
        private void User_Manage_Import_BT_Click_event(object sender, EventArgs e)
        {
            OpenFileDialog open_dialog = new OpenFileDialog();

            string file_name;
            string fInfo;
            string temp;

            open_dialog.Filter = "Excel file (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|" + "All files (*.*)|*.*";

            if (open_dialog.ShowDialog() == DialogResult.OK)
            {
                file_name = open_dialog.FileName;
                fInfo = Path.GetExtension(open_dialog.FileName);
                temp = User_Manage_Import_BT.Text;
                User_Manage_Import_BT.Text = "Importing ...";
                User_Manage_Import_BT.Enabled = false;

                //Import_User_Manage_file(file_name);
                if (Import_User_Manage_file(file_name) == true)
                {
                    User_Manage_Store_BT_Click(null, null);
                }
                User_Manage_Import_BT.Enabled = true;
                User_Manage_Import_BT.Text = temp;
            }
        }

        private void User_Manage_Store_BT_Click(object sender, EventArgs e)
        {
            if (Permission_Table_Form.dataGridView_View.DataSource == null)
            {
                Permission_Table_Form.Load_DataBase(User_Conn, @"SELECT * FROM dbo.PERMISSION_tb");
            }
            if (Update_SQL_Data(Permission_Table_Form.Data_da, Permission_Table_Form.Data_dtb) == true)
            {
                if (Update_SQL_Data(User_Table_Form.Data_da, User_Table_Form.Data_dtb) == true)
                {
                    MessageBox.Show("Store Data Complete", "Successful");
                    //RELOAD_DB = 1;
                }
                else
                {
                    MessageBox.Show("Store Data Fail", "Failed");
                }
            }
        }


        private void User_Manage_Search_Component_BT_Click(object sender, EventArgs e)
        {
            string component_id, component_cell;
            bool component_exist;

            component_id = User_Manage_Search_Component.My_TextBox.Text.ToString().Trim();
            int i, max_row;

            if (component_id != "")
            {
                component_exist = false;
                max_row = User_Table_Form.dataGridView_View.RowCount;
                for (i = 0; i < max_row - 1; i++)
                {
                    component_cell = User_Table_Form.dataGridView_View.Rows[i].Cells["UserName"].Value.ToString().Trim();
                    if (component_id == component_cell)
                    {
                        User_Table_Form.dataGridView_View.CurrentCell = User_Table_Form.dataGridView_View.Rows[i].Cells["Username"];
                        User_Table_Form.dataGridView_View.CurrentCell.Selected = true;
                        component_exist = true;
                        break;
                    }
                }
                if (component_exist == false)
                {
                    MessageBox.Show("Component number : " + component_id + " isn't exist", "Error");
                }
            }
            else
            {
                MessageBox.Show("Please Fill Name !", "Warning");
            }
        }

        private void User_Manage_Add_Part_number_BT_Click(object sender, EventArgs e)
        {
            string username = "", password = "", msnv = "", aliasname = "", permission_id = "",
                   bom_manage = "", import_material = "";
            Permission_Table_Form.Load_DataBase(User_Conn, @"SELECT * FROM dbo.PERMISSION_tb");
            //User_Table_Form.Load_DataBase(User_Conn, @"SELECT * FROM dbo.USER_DASHBOARD_tb");

            username = User_Manage_UserName.My_TextBox.Text.ToString().Trim();
            password = User_Manage_Password.My_TextBox.Text.ToString().Trim();
            msnv = User_Manage_MSNV.My_TextBox.Text.ToString().Trim();
            aliasname = User_Manage_AliasName.My_TextBox.Text.ToString().Trim();
            permission_id = User_Manage_PermissionID.My_TextBox.Text.ToString().Trim();
            bom_manage = User_Manage_BomManage.My_TextBox.Text.ToString().Trim();
            import_material = User_Manage_ImportMaterial.My_TextBox.Text.ToString().Trim();

            if (username == "" || password == "" || msnv == "" || aliasname == "" || permission_id == "")
            {
                MessageBox.Show("Please fill cells(*)", "Warning");
            }
            else
            {
                if (Username_Is_New_item(username) == true)
                {
                    DataRow new_row = User_Table_Form.Data_dtb.NewRow();
                    new_row["UserName"] = username;
                    new_row["Password"] = Encrypt_Pass(password);
                    new_row["MSNV"] = msnv;
                    new_row["AliasName"] = aliasname;
                    new_row["PermissionID"] = permission_id;
                    //new_row["fill"] =  ;
                    //new_row["fill"] = ;
                    //new_row["fill"] = ;
                    //User_Table_Form.Data_dtb.Rows.Add(new_row);
                    if (Permission_ID_Is_New_item(permission_id) == true)
                    {
                        DataRow new_row1 = Permission_Table_Form.Data_dtb.NewRow();
                        new_row1["PermissionID"] = permission_id;
                        if ((bom_manage == "true") || (bom_manage == "True") || (bom_manage == "TRUE"))
                        {
                            new_row1["BomManage"] = 1;
                        }
                        else if ((bom_manage == "false") || (bom_manage == "False") || (bom_manage == "FALSE"))
                        {
                            new_row1["BomManage"] = 0;
                        }
                        else
                        {
                            new_row1["BomManage"] = 0;
                        }
                        if ((import_material == "true") || (import_material == "True") || (import_material == "TRUE"))
                        {
                            new_row1["ImportMaterial"] = 1;
                        }
                        else if ((import_material == "false") || (import_material == "False") || (import_material == "FALSE"))
                        {
                            new_row1["ImportMaterial"] = 0;
                        }
                        else
                        {
                            new_row1["ImportMaterial"] = 0;
                        }
                        User_Table_Form.Data_dtb.Rows.Add(new_row);
                        Permission_Table_Form.Data_dtb.Rows.Add(new_row1);
                        User_Manage_Search_Component.My_TextBox.Text = username;
                        User_Manage_Search_Component_BT_Click(null, null);
                        DialogResult result = MessageBox.Show("Do you want to save this username ?", "Attention", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (result == DialogResult.Yes)
                        {
                            User_Manage_Store_BT_Click(null, null);

                        }
                        else
                        {
                            new_row.Delete();
                            new_row1.Delete();
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Permission_ID: " + permission_id + " is exist.\nPlease fill new permission_id", "Warning");
                        return;
                    }
                    //User_Manage_Search_Component.My_TextBox.Text = username;
                    //User_Manage_Search_Component_BT_Click(null, null);
                    //DialogResult result = MessageBox.Show("Do you want to save this username ?", "Attention", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    //if (result == DialogResult.Yes)
                    //{
                    //    User_Manage_Store_BT_Click(null, null);
                    //}
                    //else
                    //{
                    //    new_row.Delete();
                    //    return;
                    //}
                }
                else
                {
                    User_Manage_Search_Component.My_TextBox.Text = username;
                    MessageBox.Show("UserName: " + username + " is exist", "Warning");
                    User_Manage_Search_Component_BT_Click(null, null);

                }
            }

        }

        private bool Username_Is_New_item(string username)
        {
            bool ret_var = true;
            string filterExpression = "";
            filterExpression = "UserName =" + "'" + username + "'";

            DataRow[] rows = User_Table_Form.Data_dtb.Select(filterExpression);
            if (rows.Length > 0)
            {
                ret_var = false;
            }
            return ret_var;
        }

        private bool Permission_ID_Is_New_item(string permission_id)
        {
            bool ret_var = true;
            string filterExpression = "";
            filterExpression = "PermissionID =" + "'" + permission_id + "'";

            DataRow[] rows = Permission_Table_Form.Data_dtb.Select(filterExpression);
            if (rows.Length > 0)
            {
                ret_var = false;
            }
            return ret_var;
        }
    }
}