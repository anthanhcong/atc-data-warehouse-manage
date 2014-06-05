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
        public DataTable Recent_day_TBL;
        DataSet Recent_day_ds = new DataSet();
        SqlDataAdapter Recent_day_da;

                /**********************************/
                //        Input Nhap Khau         //
                /*********************************/

        private void INPUT_NK_Search_BT_Click(object sender, EventArgs e)
        {
            //string component_id, component_cell;
            //bool component_exist;

            //component_id = INPUT_NK_Search_Txt_Lb.My_TextBox.Text.ToString().Trim();
            //int i, max_row;

            //if (component_id != "")
            //{
            //    component_exist = false;
            //    max_row = INPUT_NK_Table_Form.dataGridView_View.RowCount;
            //    for (i = 0; i < max_row - 1; i++)
            //    {
            //        component_cell = INPUT_NK_Table_Form.dataGridView_View.Rows[i].Cells["So_TK"].Value.ToString().Trim();
            //        if (component_id == component_cell)
            //        {
            //            INPUT_NK_Table_Form.dataGridView_View.CurrentCell = INPUT_NK_Table_Form.dataGridView_View.Rows[i].Cells["So_TK"];
            //            INPUT_NK_Table_Form.dataGridView_View.CurrentCell.Selected = true;
            //            component_exist = true;
            //            break;
            //        }
            //    }
            //    if (component_exist == false)
            //    {
            //        MessageBox.Show("Component number : " + component_id + " isn't exist", "Error");
            //    }
            //}
            //else
            //{
            //    MessageBox.Show("Please Fill Name !", "Warning");
            //}
        }

        private void INPUT_NK_Find_BtL_Click(object sender, EventArgs e)
        {
            string so_tk, ma_lh, ma_hs, ma_hang, start_day, end_day, recent_day;
            bool dk_where = false;

            string sql_cmd = @"SELECT * FROM [WHM_INFOMATION_DB].[dbo].[INPUT_NK_tb]";
            if (INPUT_NK_Check_So_TK.My_CheckBox.Checked == true)
            {
                so_tk = INPUT_NK_So_TK_TxbL.My_TextBox.Text.ToString().Trim();
                if (so_tk != "")
                {
                    sql_cmd += " where So_TK = " + "'" + so_tk + "'";
                    dk_where = true;
                }
            }
            if (INPUT_NK_Check_Ngay_DK.My_CheckBox.Checked == true)
            {
                start_day = INPUT_NK_Start_Date.My_picker.Value.Date.ToString("yyyy-MM-dd");
                end_day = INPUT_NK_End_Date.My_picker.Value.Date.ToString("yyyy-MM-dd");
                if (dk_where == true)
                {
                    sql_cmd += " and Ngay_DK >= " + "'" + start_day + "'";
                    sql_cmd += " and Ngay_DK <= " + "'" + end_day + "'";
                }
                else
                {
                    sql_cmd += " where Ngay_DK >= " + "'" + start_day + "'";
                    sql_cmd += " and Ngay_DK <= " + "'" + end_day + "'";
                    dk_where = true;
                }
            }
            if (INPUT_NK_Check_Ma_loai_hinh.My_CheckBox.Checked == true)
            {
                ma_lh = INPUT_NK_Ma_loai_hinh_TxbL.My_TextBox.Text.ToString().Trim();
                if (dk_where == true)
                {
                    sql_cmd += " and Ma_loai_hinh = " + "'" + ma_lh + "'";
                }
                else
                {
                    sql_cmd += " where Ma_loai_hinh = " + "'" + ma_lh + "'";
                    dk_where = true;
                }
            }
            if (INPUT_NK_Check_Ma_HS.My_CheckBox.Checked == true)
            {
                ma_hs = INPUT_NK_Ma_HS_TxbL.My_TextBox.Text.ToString().Trim();
                if (dk_where == true)
                {
                    sql_cmd += " and Ma_HS = " + "'" + ma_hs + "'";
                }
                else
                {
                    sql_cmd += " where Ma_HS = " + "'" + ma_hs + "'";
                    dk_where = true;
                }
            }
            if (INPUT_NK_Check_Ma_hang.My_CheckBox.Checked == true)
            {
                ma_hang = INPUT_NK_Ma_hang_TxbL.My_TextBox.Text.ToString().Trim();
                if (dk_where == true)
                {
                    sql_cmd += " and Ma_hang = " + "'" + ma_hang + "'";
                }
                else
                {
                    sql_cmd += " where Ma_hang = " + "'" + ma_hang + "'";
                    dk_where = true;
                }
            }
            if (INPUT_NK_Check_Recent_Day.My_CheckBox.Checked == true)
            {
                string sql_recent_day = @"SELECT Ngay_DK = (Max(Ngay_DK)) FROM [WHM_INFOMATION_DB].[dbo].[INPUT_NK_tb]";

                Recent_day_TBL = Get_SQL_Data(Database_WHM_Info_Con_Str, sql_recent_day, ref Recent_day_da, ref Recent_day_ds);
                //INPUT_NK_Table_Form.Load_DataBase(Database_WHM_Info_Con_Str, sql_recent_day);
                if (Recent_day_TBL != null)
                {
                    recent_day = Convert.ToDateTime(Recent_day_TBL.Rows[0]["Ngay_DK"].ToString().Trim()).Date.ToString("yyyy-MM-dd");
                    if (dk_where == true)
                    {
                        sql_cmd += " and Ngay_DK = " + "'" + recent_day + "'";
                    }
                    else
                    {
                        sql_cmd += " where Ngay_DK = " + "'" + recent_day + "'";
                        dk_where = true;
                    }
                }
                
            }
            INPUT_NK_Table_Form.Load_DataBase(Database_WHM_Info_Con_Str, sql_cmd);
        }

        private void INPUT_NK_KD_Import_BT_Click(object sender, EventArgs e)
        {
            OpenFileDialog open_dialog = new OpenFileDialog();

            string file_name;
            string fInfo;
            string temp;

            Ma_KD_Or_SX = "N-KD";

            open_dialog.Filter = "Excel file (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|" + "All files (*.*)|*.*";
            open_dialog.Multiselect = true;
            if (open_dialog.ShowDialog() == DialogResult.OK)
            {
                file_name = open_dialog.FileName;
                fInfo = Path.GetExtension(open_dialog.FileName);
                temp = INPUT_NK_KD_Import_BT.Text;
                INPUT_NK_KD_Import_BT.Text = "Running ...";
                INPUT_NK_KD_Import_BT.Enabled = false;
                //Import_INPUT_SXXK_NK_Table_in_file(file_name);
                Import_INPUT_NK_Table_in_file(file_name);
                INPUT_NK_KD_Import_BT.Enabled = true;
                INPUT_NK_KD_Import_BT.Text = temp;
            }

        }

        private void INPUT_NK_SX_Import_BT_Click(object sender, EventArgs e)
        {
            OpenFileDialog open_dialog = new OpenFileDialog();

            string file_name;
            string fInfo;
            string temp;

            Ma_KD_Or_SX = "N-SX";

            open_dialog.Filter = "Excel file (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|" + "All files (*.*)|*.*";
            open_dialog.Multiselect = true;
            if (open_dialog.ShowDialog() == DialogResult.OK)
            {
                file_name = open_dialog.FileName;
                fInfo = Path.GetExtension(open_dialog.FileName);
                temp = INPUT_NK_SX_Import_BT.Text;
                INPUT_NK_SX_Import_BT.Text = "Running ...";
                INPUT_NK_SX_Import_BT.Enabled = false;
                //Import_INPUT_SXXK_NK_Table_in_file(file_name);
                Import_INPUT_NK_Table_in_file(file_name);
                INPUT_NK_SX_Import_BT.Enabled = true;
                INPUT_NK_SX_Import_BT.Text = temp;
            }

        }

        private void INPUT_NK_Store_BT_Click(object sender, EventArgs e)
        {
            if (Update_SQL_Data(INPUT_NK_Table_Form.Data_da, INPUT_NK_Table_Form.Data_dtb) == true)
            {
                MessageBox.Show("Store Data Complete", "Successful");
                //RELOAD_DB = 3;
                //Load_WHM_Info_DB();
            }
            else
            {
                MessageBox.Show("Store Data Fail", "Failed");
            }
        }

        private void INPUT_NK_So_TK_CbxL_Click(object sender, EventArgs e)
        {
            Load_List_TK_NK(Database_WHM_Info_Con_Str);
        }

        private void INPUT_NK_So_TK_CbxL_Text_Change(object sender, EventArgs e)
        {
            Load_Ma_LH_NK(Database_WHM_Info_Con_Str);
            NK_Find_So_TK();
        }

        /**************************************************************************************/

                    /*******************************************/
                    //             Input Xuất khẩu             //
                    /******************************************/

        private void INPUT_XK_Search_BT_Click(object sender, EventArgs e)
        {
            string component_id, component_cell;
            bool component_exist;

            component_id = INPUT_XK_Search_Txt_Lb.My_TextBox.Text.ToString().Trim();
            int i, max_row;

            if (component_id != "")
            {
                component_exist = false;
                max_row = INPUT_XK_Table_Form.dataGridView_View.RowCount;
                for (i = 0; i < max_row - 1; i++)
                {
                    component_cell = INPUT_XK_Table_Form.dataGridView_View.Rows[i].Cells["So_TK"].Value.ToString().Trim();
                    if (component_id == component_cell)
                    {
                        INPUT_XK_Table_Form.dataGridView_View.CurrentCell = INPUT_XK_Table_Form.dataGridView_View.Rows[i].Cells["So_TK"];
                        INPUT_XK_Table_Form.dataGridView_View.CurrentCell.Selected = true;
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

        private void INPUT_XK_SX_Import_BT_Click(object sender, EventArgs e)
        {
            OpenFileDialog open_dialog = new OpenFileDialog();

            string file_name;
            string fInfo;
            string temp;

            Ma_KD_Or_SX = "X-SX";

            open_dialog.Filter = "Excel file (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|" + "All files (*.*)|*.*";
            open_dialog.Multiselect = true;
            if (open_dialog.ShowDialog() == DialogResult.OK)
            {
                file_name = open_dialog.FileName;
                fInfo = Path.GetExtension(open_dialog.FileName);
                temp = INPUT_XK_SX_Import_BT.Text;
                INPUT_XK_SX_Import_BT.Text = "Running ...";
                INPUT_XK_SX_Import_BT.Enabled = false;
                Import_INPUT_SXXK_XK_Table_in_file(file_name);
                INPUT_XK_SX_Import_BT.Enabled = true;
                INPUT_XK_SX_Import_BT.Text = temp;
            }

        }

        private void INPUT_XK_KD_Import_BT_Click(object sender, EventArgs e)
        {
            OpenFileDialog open_dialog = new OpenFileDialog();

            string file_name;
            string fInfo;
            string temp;

            Ma_KD_Or_SX = "X-KD";

            open_dialog.Filter = "Excel file (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|" + "All files (*.*)|*.*";
            open_dialog.Multiselect = true;
            if (open_dialog.ShowDialog() == DialogResult.OK)
            {
                file_name = open_dialog.FileName;
                fInfo = Path.GetExtension(open_dialog.FileName);
                temp = INPUT_XK_KD_Import_BT.Text;
                INPUT_XK_KD_Import_BT.Text = "Running ...";
                INPUT_XK_KD_Import_BT.Enabled = false;
                Import_INPUT_SXXK_XK_Table_in_file(file_name);
                INPUT_XK_KD_Import_BT.Enabled = true;
                INPUT_XK_KD_Import_BT.Text = temp;
            }

        }

        private void INPUT_XK_Store_BT_Click(object sender, EventArgs e)
        {
            if (Update_SQL_Data(INPUT_XK_Table_Form.Data_da, INPUT_XK_Table_Form.Data_dtb) == true)
            {
                MessageBox.Show("Store Data Complete", "Successful");
                //RELOAD_DB = 3;
                //Load_WHM_Info_DB();
            }
            else
            {
                MessageBox.Show("Store Data Fail", "Failed");
            }
        }

      
    }
}