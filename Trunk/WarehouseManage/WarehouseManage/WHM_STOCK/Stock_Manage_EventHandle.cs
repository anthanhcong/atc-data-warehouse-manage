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
        /*************************************************/
        //            Stock Manage EventHandle           //
        /************************************************/

        private void Stock_Search_BT_Click(object sender, EventArgs e)
        {
            string component_id, component_cell;
            bool component_exist;

            component_id = Stock_Search_Txt_Lb.My_TextBox.Text.ToString().Trim();
            int i, max_row;

            if (component_id != "")
            {
                component_exist = false;
                max_row = Stock_Table_Form.dataGridView_View.RowCount;
                for (i = 0; i < max_row - 1; i++)
                {
                    component_cell = Stock_Table_Form.dataGridView_View.Rows[i].Cells["Item_ID"].Value.ToString().Trim();
                    if (component_id == component_cell)
                    {
                        Stock_Table_Form.dataGridView_View.CurrentCell = Stock_Table_Form.dataGridView_View.Rows[i].Cells["Item_ID"];
                        Stock_Table_Form.dataGridView_View.CurrentCell.Selected = true;
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

        private void Stock_Import_BT_Click(object sender, EventArgs e)
        {
            OpenFileDialog open_dialog = new OpenFileDialog();

            string file_name;
            string fInfo;
            string temp;

            open_dialog.Filter = "Excel file (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|" + "All files (*.*)|*.*";
            open_dialog.Multiselect = true;
            if (open_dialog.ShowDialog() == DialogResult.OK)
            {
                file_name = open_dialog.FileName;
                fInfo = Path.GetExtension(open_dialog.FileName);
                temp = Stock_Import_BT.Text;
                Stock_Import_BT.Text = "Running ...";
                Stock_Import_BT.Enabled = false;
                Import_Stock_Manage_Table_in_file(file_name);
                Stock_Import_BT.Enabled = true;
                Stock_Import_BT.Text = temp;
            }
        }

        private void Stock_Process_BT_Click(object sender, EventArgs e)
        {

        }

        private void Stock_Store_BT_Click(object sender, EventArgs e)
        {
            if (Update_SQL_Data(Stock_Table_Form.Data_da, Stock_Table_Form.Data_dtb) == true)
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

                 /************************************************/
                //          WH_ID_List Manage EventHandle        //
                /************************************************/

        private void WH_ID_List_Search_BT_Click(object sender, EventArgs e)
        {
            string component_id, component_cell;
            bool component_exist;

            component_id = WH_ID_List_Search_Txt_Lb.My_TextBox.Text.ToString().Trim();
            int i, max_row;

            if (component_id != "")
            {
                component_exist = false;
                max_row = WH_ID_List_Table_Form.dataGridView_View.RowCount;
                for (i = 0; i < max_row - 1; i++)
                {
                    component_cell = WH_ID_List_Table_Form.dataGridView_View.Rows[i].Cells["WareHouse_ID"].Value.ToString().Trim();
                    if (component_id == component_cell)
                    {
                        WH_ID_List_Table_Form.dataGridView_View.CurrentCell = WH_ID_List_Table_Form.dataGridView_View.Rows[i].Cells["WareHouse_ID"];
                        WH_ID_List_Table_Form.dataGridView_View.CurrentCell.Selected = true;
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

        private void WH_ID_List_Import_BT_Click(object sender, EventArgs e)
        {
            OpenFileDialog open_dialog = new OpenFileDialog();

            string file_name;
            string fInfo;
            string temp;

            open_dialog.Filter = "Excel file (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|" + "All files (*.*)|*.*";
            open_dialog.Multiselect = true;
            if (open_dialog.ShowDialog() == DialogResult.OK)
            {
                file_name = open_dialog.FileName;
                fInfo = Path.GetExtension(open_dialog.FileName);
                temp = WH_ID_List_Import_BT.Text;
                WH_ID_List_Import_BT.Text = "Running ...";
                WH_ID_List_Import_BT.Enabled = false;
                Import_WH_List_Manage_Table_in_file(file_name);
                WH_ID_List_Import_BT.Enabled = true;
                WH_ID_List_Import_BT.Text = temp;
            }
        }

        private void WH_ID_List_Process_BT_Click(object sender, EventArgs e)
        {

        }

        private void WH_ID_List_Store_BT_Click(object sender, EventArgs e)
        {
            if (Update_SQL_Data(WH_ID_List_Table_Form.Data_da, WH_ID_List_Table_Form.Data_dtb) == true)
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

        private void Ma_List_Import_BT_Click(object sender, EventArgs e)
        {
            OpenFileDialog open_dialog = new OpenFileDialog();

            string file_name;
            string fInfo;
            string temp;

            open_dialog.Filter = "Excel file (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|" + "All files (*.*)|*.*";
            open_dialog.Multiselect = true;
            if (open_dialog.ShowDialog() == DialogResult.OK)
            {
                file_name = open_dialog.FileName;
                fInfo = Path.GetExtension(open_dialog.FileName);
                temp = Ma_List_Import_BT.Text;
                Ma_List_Import_BT.Text = "Running ...";
                Ma_List_Import_BT.Enabled = false;
                Import_Ma_List_Manage_Table_in_file(file_name);
                Ma_List_Import_BT.Enabled = true;
                Ma_List_Import_BT.Text = temp;
            }
        }

        private void Ma_List_Store_BT_Click(object sender, EventArgs e)
        {
            if (Update_SQL_Data(Material_List_Table_Form.Data_da, Material_List_Table_Form.Data_dtb) == true)
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