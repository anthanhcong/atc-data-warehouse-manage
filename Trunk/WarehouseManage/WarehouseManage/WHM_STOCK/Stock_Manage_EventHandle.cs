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

        private void Stock_WH_ID_List_cbx_SelectedValueChanged(Object sender, EventArgs e)
        {
            string wh_id, load_cmd;
            if (Stock_WH_ID_List_cbx.My_Combo.SelectedValue != null)
            {
                wh_id = Stock_WH_ID_List_cbx.My_Combo.SelectedValue.ToString().Trim();
                load_cmd = Load_Stock_cmd;
                load_cmd += " where WareHouse_ID =" + "'" + wh_id + "'";
                if (wh_id != "")
                {
                    Load_Stock_Table(load_cmd);
                    if (Stock_Table_Form.Data_dtb.Rows.Count == 0)
                    {
                        MessageBox.Show("Can not find WareHouse_ID: '" + wh_id + "'.", "Warning");
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Please fill WareHouse ID.", "Warning");
                    return;
                }
            }
        }

        private void Stock_Part_Number_cbx_SelectedValueChanged(Object sender, EventArgs e)
        {
            string part, load_cmd;
            if (Stock_WH_ID_List_cbx.My_Combo.SelectedValue != null)
            {
                part = Stock_WH_ID_List_cbx.My_Combo.SelectedValue.ToString().Trim();
                load_cmd = Load_Stock_cmd;
                load_cmd += " where Part_Number =" + "'" + part + "'";
                if (part != "")
                {
                    Load_Stock_Table(load_cmd);
                    if (Stock_Table_Form.Data_dtb.Rows.Count == 0)
                    {
                        MessageBox.Show("Can not find Part Number: '" + part + "'.", "Warning");
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Please fill Part Number.", "Warning");
                    return;
                }
            }
        }
        
        private void Stock_WH_ID_List_cbx_KeyDown(object sender, KeyEventArgs e)
        {
            string wh_id, load_cmd;

            wh_id = Stock_WH_ID_List_cbx.My_Combo.Text.ToString().Trim();
            if (e.KeyCode == Keys.Enter)
            {
                load_cmd = Load_Stock_cmd;
                load_cmd += " where WareHouse_ID =" + "'" + wh_id + "'";
                if (wh_id != "")
                {
                    Load_Stock_Table(wh_id);
                    if (Stock_Table_Form.Data_dtb.Rows.Count == 0)
                    {
                        MessageBox.Show("Can not find WH ID: '" + wh_id + "'.", "Warning");
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Please fill WH ID.", "Warning");
                    return;
                }
            }
        }

        private void Stock_Part_Number_cbx_KeyDown(object sender, KeyEventArgs e)
        {
            string wh_id, load_cmd;

            wh_id = Stock_WH_ID_List_cbx.My_Combo.Text.ToString().Trim();
            if (e.KeyCode == Keys.Enter)
            {
                load_cmd = Load_Stock_cmd;
                load_cmd += " where Part_Number =" + "'" + wh_id + "'";
                if (wh_id != "")
                {
                    Load_Stock_Table(load_cmd);
                    if (Stock_Table_Form.Data_dtb.Rows.Count == 0)
                    {
                        MessageBox.Show("Can not find Part_Number: '" + wh_id + "'.", "Warning");
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Please fill Part Number.", "Warning");
                    return;
                }
            }
        }

        private void Stock_Bin_Txt_KeyDown(object sender, KeyEventArgs e)
        {
            string bin, load_cmd;

            bin = Stock_Bin_Txt.My_TextBox.Text.ToString().Trim();
            if (e.KeyCode == Keys.Enter)
            {
                load_cmd = Load_Stock_cmd;
                load_cmd += " where Bin =" + "'" + bin + "'";
                if (bin != "")
                {
                    Load_Stock_Table(load_cmd);
                    if (Stock_Table_Form.Data_dtb.Rows.Count == 0)
                    {
                        MessageBox.Show("Can not find Bin: '" + bin + "'.", "Warning");
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Please fill Bin.", "Warning");
                    return;
                }
            }
        }

        private void Stock_Plant_Txt_KeyDown(object sender, KeyEventArgs e)
        {
            string plant, load_cmd;

            plant = Stock_Plant_Txt.My_TextBox.Text.ToString().Trim();
            if (e.KeyCode == Keys.Enter)
            {
                load_cmd = Load_Stock_cmd;
                load_cmd += " where Plant =" + "'" + plant + "'";
                if (plant != "")
                {
                    Load_Stock_Table(load_cmd);
                    if (Stock_Table_Form.Data_dtb.Rows.Count == 0)
                    {
                        MessageBox.Show("Can not find Bin: '" + plant + "'.", "Warning");
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Please fill Bin.", "Warning");
                    return;
                }
            }
        }

        void Stock_Table_Form_CellDoubleClick(Object sender, EventArgs e)
        {
            string cell_value, load_cmd;

            load_cmd = @"SELECT * FROM [WHM_STOCK_DB].[dbo].[Material_Stock_tb]";
            if (Stock_Table_Form.dataGridView_View.CurrentCell == Stock_Table_Form.dataGridView_View.CurrentRow.Cells["Part_Number"])
            {
                Stock_Manage_Single_View_All.My_CheckBox.Checked = false;
                cell_value = Stock_Table_Form.dataGridView_View.CurrentRow.Cells["Part_Number"].Value.ToString().Trim();
                load_cmd += " where Part_Number = " + "'" + cell_value + "'";
                Load_Stock_Table(load_cmd);
                for (int row = 0; row < Stock_Table_Form.dataGridView_View.Rows.Count - 1; row++)
                {
                    Stock_Table_Form.dataGridView_View["Part_Number", row].Style.BackColor = Color.YellowGreen;
                }
            }
            else if (Stock_Table_Form.dataGridView_View.CurrentCell == Stock_Table_Form.dataGridView_View.CurrentRow.Cells["Bin"])
            {
                Stock_Manage_Single_View_All.My_CheckBox.Checked = false;
                cell_value = Stock_Table_Form.dataGridView_View.CurrentRow.Cells["Bin"].Value.ToString().Trim();
                load_cmd += " where Bin = " + "'" + cell_value + "'";
                Load_Stock_Table(load_cmd);
                for (int row = 0; row < Stock_Table_Form.dataGridView_View.Rows.Count - 1; row++)
                {
                    Stock_Table_Form.dataGridView_View["Bin", row].Style.BackColor = Color.YellowGreen;
                }
            }
            else if (Stock_Table_Form.dataGridView_View.CurrentCell == Stock_Table_Form.dataGridView_View.CurrentRow.Cells["Plant"])
            {
                Stock_Manage_Single_View_All.My_CheckBox.Checked = false;
                cell_value = Stock_Table_Form.dataGridView_View.CurrentRow.Cells["Plant"].Value.ToString().Trim();
                load_cmd += " where Plant = " + "'" + cell_value + "'";
                Load_Stock_Table(load_cmd);
                for (int row = 0; row < Stock_Table_Form.dataGridView_View.Rows.Count - 1; row++)
                {
                    Stock_Table_Form.dataGridView_View["Plant", row].Style.BackColor = Color.YellowGreen;
                }
            }
            else if (Stock_Table_Form.dataGridView_View.CurrentCell == Stock_Table_Form.dataGridView_View.CurrentRow.Cells["WareHouse_ID"])
            {
                Stock_Manage_Single_View_All.My_CheckBox.Checked = false;
                cell_value = Stock_Table_Form.dataGridView_View.CurrentRow.Cells["WareHouse_ID"].Value.ToString().Trim();
                load_cmd += " where WareHouse_ID = " + "'" + cell_value + "'";
                Load_Stock_Table(load_cmd);
                for (int row = 0; row < Stock_Table_Form.dataGridView_View.Rows.Count - 1; row++)
                {
                    Stock_Table_Form.dataGridView_View["WareHouse_ID", row].Style.BackColor = Color.YellowGreen;
                }

            }
        }

        private void Stock_Manage_Single_View_All_CheckedChanged(object sender, EventArgs e)
        {
            if (Stock_Manage_Single_View_All.My_CheckBox.Checked == true)
            {
                Load_Stock_Table_All();
                Stock_WH_ID_List_cbx.My_Combo.Enabled = false;
                Stock_Part_Number_cbx.My_Combo.Enabled = false;
                Stock_Bin_Txt.My_TextBox.Enabled = false;
                Stock_Plant_Txt.My_TextBox.Enabled = false;
                Stock_Manage_Single_Check_WH_ID.My_CheckBox.Enabled = false;
                Stock_Manage_Single_Check_Part_Number.My_CheckBox.Enabled = false;
                Stock_Manage_Single_Check_Plant.My_CheckBox.Enabled = false;
                Stock_Manage_Single_Check_Bin.My_CheckBox.Enabled = false;
                Stock_Import_BT.Enabled = false;
                Stock_Store_BT.Enabled = false;
                Stock_Search_BT.Enabled = false;
            }
            else
            {
                Stock_WH_ID_List_cbx.My_Combo.Enabled = true;
                Stock_Part_Number_cbx.My_Combo.Enabled = true;
                Stock_Bin_Txt.My_TextBox.Enabled = true;
                Stock_Plant_Txt.My_TextBox.Enabled = true;
                Stock_Manage_Single_Check_WH_ID.My_CheckBox.Enabled = true;
                Stock_Manage_Single_Check_Part_Number.My_CheckBox.Enabled = true;
                Stock_Manage_Single_Check_Plant.My_CheckBox.Enabled = true;
                Stock_Manage_Single_Check_Bin.My_CheckBox.Enabled = true;
                Stock_Import_BT.Enabled = true;
                Stock_Store_BT.Enabled = true;
                Stock_Search_BT.Enabled = true;
            }
        }
        
        private void Stock_Search_BT_Click(object sender, EventArgs e)
        {
            //string component_id, component_cell;
            //bool component_exist;

            //component_id = Stock_Search_Txt_Lb.My_TextBox.Text.ToString().Trim();
            //int i, max_row;

            //if (component_id != "")
            //{
            //    component_exist = false;
            //    max_row = Stock_Table_Form.dataGridView_View.RowCount;
            //    for (i = 0; i < max_row - 1; i++)
            //    {
            //        component_cell = Stock_Table_Form.dataGridView_View.Rows[i].Cells["Item_ID"].Value.ToString().Trim();
            //        if (component_id == component_cell)
            //        {
            //            Stock_Table_Form.dataGridView_View.CurrentCell = Stock_Table_Form.dataGridView_View.Rows[i].Cells["Item_ID"];
            //            Stock_Table_Form.dataGridView_View.CurrentCell.Selected = true;
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
            string part, wh_id, plant, bin, sql_cmd, sql_cmd_temp;

            Show_Sort_log = "";
            if (INPUT_XK_Check_So_TK.My_CheckBox.Checked == true)
            {
                part = INPUT_XK_So_TK_TxbL.My_TextBox.Text.ToString().Trim();
                if (part != "")
                {
                    Show_Sort_log = "Part_Number: " + part;
                }
                else
                {
                    Show_Sort_log = "Hãy điền vào Số TK";
                }
                sql_cmd_temp = " where Part_Number = " + "'" + part + "'";
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