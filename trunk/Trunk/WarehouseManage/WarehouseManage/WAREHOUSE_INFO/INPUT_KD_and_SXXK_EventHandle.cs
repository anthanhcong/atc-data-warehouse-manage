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
                /**********************************/
                //        Input Nhap Khau         //
                /*********************************/

        private void INPUT_NK_Search_BT_Click(object sender, EventArgs e)
        {
            string component_id, component_cell;
            bool component_exist;

            component_id = INPUT_NK_Search_Txt_Lb.My_TextBox.Text.ToString().Trim();
            int i, max_row;

            if (component_id != "")
            {
                component_exist = false;
                max_row = INPUT_NK_Table_Form.dataGridView_View.RowCount;
                for (i = 0; i < max_row - 1; i++)
                {
                    component_cell = INPUT_NK_Table_Form.dataGridView_View.Rows[i].Cells["So_TK"].Value.ToString().Trim();
                    if (component_id == component_cell)
                    {
                        INPUT_NK_Table_Form.dataGridView_View.CurrentCell = INPUT_NK_Table_Form.dataGridView_View.Rows[i].Cells["So_TK"];
                        INPUT_NK_Table_Form.dataGridView_View.CurrentCell.Selected = true;
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
                Import_INPUT_SXXK_NK_Table_in_file(file_name);
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
                Import_INPUT_SXXK_NK_Table_in_file(file_name);
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