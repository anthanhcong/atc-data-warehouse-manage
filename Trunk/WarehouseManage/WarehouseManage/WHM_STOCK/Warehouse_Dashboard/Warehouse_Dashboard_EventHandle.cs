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
        void WH_Daskboard_Mother_WH_CbxL_SelectedValueChanged(Object sender, EventArgs e)
        {
            string mother_wh;
            if (WH_Daskboard_Mother_WH_CbxL.My_Combo.SelectedValue != null)
            {
                mother_wh = WH_Daskboard_Mother_WH_CbxL.My_Combo.SelectedValue.ToString().Trim();
                WH_Dashboard_Load_with_Mother_wh(mother_wh);
            }
        }
        
        void WH_ID_with_MaLH_WH_ID_CbxL_SelectedValueChanged(Object sender, EventArgs e)
        {
            string wh_id;
            if (WH_ID_with_MaLH_WH_ID_CbxL.My_Combo.SelectedValue != null)
            {
                wh_id = WH_ID_with_MaLH_WH_ID_CbxL.My_Combo.SelectedValue.ToString().Trim();
                WH_ID_with_MaLH_Load_with_WH_ID(wh_id);
            }
        }

        void WH_ID_with_MaLH_Table_Form_CellDoubleClick(Object sender, EventArgs e)
        {
            if (WH_ID_with_MaLH_Table_Form.dataGridView_View.CurrentCell == WH_ID_with_MaLH_Table_Form.dataGridView_View.CurrentRow.Cells["ID"])
            {
                WH_ID_with_MaLH_View_All.My_CheckBox.Checked = false;
                WH_ID_with_MLH_Store_BT.My_Button.Enabled = false;
                WH_ID_with_MaLH_ID_TbxL.My_TextBox.Text = WH_ID_with_MaLH_Table_Form.dataGridView_View.CurrentRow.Cells["ID"].Value.ToString().Trim();
                WH_ID_with_MaLH_Ma_LH_TbxL.My_TextBox.Text = WH_ID_with_MaLH_Table_Form.dataGridView_View.CurrentRow.Cells["Ma_loai_hinh"].Value.ToString().Trim();
                WH_ID_with_MaLH_WH_ID_CbxL.My_Combo.Text = WH_ID_with_MaLH_Table_Form.dataGridView_View.CurrentRow.Cells["WareHouse_ID"].Value.ToString().Trim();
                WH_ID_with_MaLH_Im_or_Ex_TbxL.My_TextBox.Text = WH_ID_with_MaLH_Table_Form.dataGridView_View.CurrentRow.Cells["Import_Export"].Value.ToString().Trim();
                WH_ID_with_MaLH_Ty_le_TbxL.My_TextBox.Text = WH_ID_with_MaLH_Table_Form.dataGridView_View.CurrentRow.Cells["Ty_le"].Value.ToString().Trim();
       
            }
        }
        void WH_Daskboard_Table_Form_CellDoubleClick(Object sender, EventArgs e)
        {
            string import;
            if (WH_Daskboard_Table_Form.dataGridView_View.CurrentCell == WH_Daskboard_Table_Form.dataGridView_View.CurrentRow.Cells["WareHouse_ID"])
            {
                WH_Daskboard_View_All.My_CheckBox.Checked = false;
                WH_Daskboard_Store_BT.My_Button.Enabled = false;
                WH_Daskboard_WH_ID_TbxL.My_TextBox.Text = WH_Daskboard_Table_Form.dataGridView_View.CurrentRow.Cells["WareHouse_ID"].Value.ToString().Trim();
                WH_Daskboard_WH_Name_TbxL.My_TextBox.Text = WH_Daskboard_Table_Form.dataGridView_View.CurrentRow.Cells["WareHouse_Name"].Value.ToString().Trim();
                WH_Daskboard_Mother_WH_CbxL.My_Combo.Text = WH_Daskboard_Table_Form.dataGridView_View.CurrentRow.Cells["Mother_WHID"].Value.ToString().Trim();
                WH_Daskboard_Note_Txt_Lb.My_TextBox.Text = WH_Daskboard_Table_Form.dataGridView_View.CurrentRow.Cells["Note"].Value.ToString().Trim();
                import = WH_Daskboard_Table_Form.dataGridView_View.CurrentRow.Cells["Import_allow"].Value.ToString().Trim();
                if (import == "True")
                {
                    WH_Daskboard_Check_Import.My_CheckBox.Checked = true;
                }
                else
                {
                    WH_Daskboard_Check_Import.My_CheckBox.Checked = false;
                }
            }
        }

        private void WH_ID_with_MaLH_View_All_CheckedChanged(object sender, EventArgs e)
        {
            if (WH_ID_with_MaLH_View_All.My_CheckBox.Checked == true)
            {
                WH_ID_with_MaLH_Load_All();
                WH_ID_with_MaLH_WH_ID_CbxL.My_Combo.Enabled = false;
                WH_ID_with_MaLH_ID_TbxL.My_TextBox.Enabled = false;
                WH_ID_with_MaLH_Ma_LH_TbxL.My_TextBox.Enabled = false;
                WH_ID_with_MaLH_Im_or_Ex_TbxL.My_TextBox.Enabled = false;
                WH_ID_with_MaLH_Ty_le_TbxL.My_TextBox.Enabled = false;
                WH_ID_with_MLH_Create_BT.My_Button.Enabled = false;
                WH_ID_with_MLH_Store_BT.My_Button.Enabled = false;
            }
            else
            {
                WH_ID_with_MaLH_WH_ID_CbxL.My_Combo.Enabled = true;
                WH_ID_with_MaLH_ID_TbxL.My_TextBox.Enabled = true;
                WH_ID_with_MaLH_Ma_LH_TbxL.My_TextBox.Enabled = true;
                WH_ID_with_MaLH_Im_or_Ex_TbxL.My_TextBox.Enabled = true;
                WH_ID_with_MaLH_Ty_le_TbxL.My_TextBox.Enabled = true;
                WH_ID_with_MLH_Create_BT.My_Button.Enabled = true;
                WH_ID_with_MLH_Store_BT.My_Button.Enabled = true;
            }
        }
        private void WH_Daskboard_View_All_CheckedChanged(object sender, EventArgs e)
        {
            if (WH_Daskboard_View_All.My_CheckBox.Checked == true)
            {
                WH_Dashboard_Load_WH_ID("All");
                WH_Daskboard_WH_ID_TbxL.My_TextBox.Enabled = false;
                WH_Daskboard_WH_Name_TbxL.My_TextBox.Enabled = false;
                WH_Daskboard_Check_Import.My_CheckBox.Enabled = false;
                WH_Daskboard_Mother_WH_CbxL.My_Combo.Enabled = false;
                WH_Daskboard_Note_Txt_Lb.My_TextBox.Enabled = false;
                WH_Daskboard_Create_BT.My_Button.Enabled = false;
                WH_Daskboard_Store_BT.My_Button.Enabled = false;
            }
            else
            {
                WH_Daskboard_WH_ID_TbxL.My_TextBox.Enabled = true;
                WH_Daskboard_WH_Name_TbxL.My_TextBox.Enabled = true;
                WH_Daskboard_Check_Import.My_CheckBox.Enabled = true;
                WH_Daskboard_Mother_WH_CbxL.My_Combo.Enabled = true;
                WH_Daskboard_Note_Txt_Lb.My_TextBox.Enabled = true;
                WH_Daskboard_Create_BT.My_Button.Enabled = true;
                WH_Daskboard_Store_BT.My_Button.Enabled = true;
            }

        }

        private void WH_Daskboard_WH_ID_Leave(object sender, EventArgs e)
        {
            string import;
            string wh_id = WH_Daskboard_WH_ID_TbxL.My_TextBox.Text.ToString().Trim();
            
            WH_Dashboard_Load_WH_ID(wh_id);
            if ((WH_Daskboard_Table_Form.Data_dtb != null) && (WH_Daskboard_Table_Form.Data_dtb.Rows.Count > 0))
            {
                // Reload infor
                WH_Daskboard_Create_BT.My_Button.Enabled = false;
                WH_Daskboard_Store_BT.My_Button.Enabled = true;

                WH_Daskboard_WH_Name_TbxL.My_TextBox.Text = WH_Daskboard_Table_Form.Data_dtb.Rows[0]["WareHouse_Name"].ToString().Trim();
                WH_Daskboard_Mother_WH_CbxL.My_Combo.Text = WH_Daskboard_Table_Form.Data_dtb.Rows[0]["Mother_WHID"].ToString().Trim();
                WH_Daskboard_Note_Txt_Lb.My_TextBox.Text = WH_Daskboard_Table_Form.Data_dtb.Rows[0]["Note"].ToString().Trim();
                import = WH_Daskboard_Table_Form.Data_dtb.Rows[0]["Import_allow"].ToString().Trim();
                if (import == "True")
                {
                    WH_Daskboard_Check_Import.My_CheckBox.Checked = true;
                }
                else
                {
                    WH_Daskboard_Check_Import.My_CheckBox.Checked = false;
                }
            }
            else
            {
                // Create New Produc
                WH_Daskboard_Create_BT.My_Button.Enabled = true;
                WH_Daskboard_Store_BT.My_Button.Enabled = false;
                WH_Daskboard_WH_Name_TbxL.My_TextBox.Text = "";
                WH_Daskboard_Note_Txt_Lb.My_TextBox.Text = "";
            }
        }

        private void WH_Daskboard_Mother_WH_CbxL_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                WH_Dashboard_Load_Table();
            }
        }

        private void WH_ID_with_MaLH_WH_ID_CbxL_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                WH_Dashboard_Load_Table();
            }
        }
        
        private void WH_Daskboard_Mother_WH_CbxL_Click(object sender, EventArgs e)
        {
            WH_Dashboard_Load_Table();
        }
        private void WH_Daskboard_Create_BT_Click(object sender, EventArgs e)
        {
            string cur_wh_id, wh_id, wh_name, mother_wh, note;
            int import_allow = 0;
            //DataRow newrow = WH_Daskboard_Table_Form.Data_dtb.NewRow();

            wh_id = WH_Daskboard_WH_ID_TbxL.My_TextBox.Text.ToString().Trim();
            wh_name = WH_Daskboard_WH_Name_TbxL.My_TextBox.Text.ToString().Trim();
            mother_wh = WH_Daskboard_Mother_WH_CbxL.My_Combo.Text.ToString().Trim();
            if (WH_Daskboard_Check_Import.My_CheckBox.Checked == true)
            {
                import_allow = 1;
            }
            note = WH_Daskboard_Note_Txt_Lb.My_TextBox.Text.ToString().Trim();
            if (wh_id == "")
            {
                MessageBox.Show("Please select WH ID", "Error");
                return;
            }
            else if (wh_name == "")
            {
                MessageBox.Show("Please select WH Name", "Error");
                return;
            }
            foreach (DataRow row in WH_Daskboard_Table_Form.Data_dtb.Rows)
            {
                cur_wh_id = row["WareHouse_ID"].ToString().Trim();
                if (cur_wh_id == wh_id)
                {
                    DialogResult thongbao;

                    thongbao = (MessageBox.Show("WH ID was created.\n"
                                                + "WH ID = " + wh_id 
                                                + "\nDo you want to update ?", " Attention ",
                                                            MessageBoxButtons.YesNo, MessageBoxIcon.Warning));
                    if (thongbao == DialogResult.Yes)
                    {
                        row["WareHouse_ID"] = wh_id;
                        row["WareHouse_Name"] = wh_name;
                        row["Import_allow"] = import_allow;
                        row["Mother_WHID"] = mother_wh;
                        row["Note"] = note;
                        WH_Daskboard_Store_BT_Click(null, null);
                    }
                    return;
                }
            }
            DataRow newrow = WH_Daskboard_Table_Form.Data_dtb.NewRow();
            newrow["WareHouse_ID"] = wh_id;
            newrow["WareHouse_Name"] = wh_name;
            newrow["Import_allow"] = import_allow;
            newrow["Mother_WHID"] = mother_wh;
            newrow["Note"] = note;
            WH_Daskboard_Table_Form.Data_dtb.Rows.Add(newrow);
            WH_Daskboard_Store_BT.My_Button.Enabled = true;
        }

        private void WH_Daskboard_Store_BT_Click(object sender, EventArgs e)
        {
            if (Update_SQL_Data(WH_Daskboard_Table_Form.Data_da, WH_Daskboard_Table_Form.Data_dtb) == true)
            {
                MessageBox.Show("Store Data Complete", "Successful");
            }
            else
            {
                MessageBox.Show("Store Data Fail", "Failed");
            }
        }

        private void WH_ID_with_MLH_Create_BT_Click(object sender, EventArgs e)
        {
            string cur_wh_id, cur_ma_lh, id, ma_lh, wh_id, im_or_ex, ty_le;
            decimal tyle;

            WH_ID_with_MaLH_Load_All();
            id = WH_ID_with_MaLH_ID_TbxL.My_TextBox.Text.ToString().Trim();
            ma_lh = WH_ID_with_MaLH_Ma_LH_TbxL.My_TextBox.Text.ToString().Trim();
            wh_id = WH_ID_with_MaLH_WH_ID_CbxL.My_Combo.Text.ToString().Trim();
            im_or_ex = WH_ID_with_MaLH_Im_or_Ex_TbxL.My_TextBox.Text.ToString().Trim();
            ty_le = WH_ID_with_MaLH_Ty_le_TbxL.My_TextBox.Text.ToString().Trim();
            //if (id == "")
            //{
            //    MessageBox.Show("Please select ID", "Error");
            //    return;
            //}else
            if (ma_lh == "")
            {
                MessageBox.Show("Please select 'Mã loại hình'", "Error");
                return;
            }
            else if (im_or_ex == "")
            {
                MessageBox.Show("Please select 'Import/Export'", "Error");
                return;
            }
            if (!isNumeric(ty_le, System.Globalization.NumberStyles.Number))
            {
                MessageBox.Show("Please enter a valid 'Tỷ lệ %'", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            foreach (DataRow row in WH_ID_with_MaLH_Table_Form.Data_dtb.Rows)
            {
                cur_wh_id = row["WareHouse_ID"].ToString().Trim();
                cur_ma_lh = row["Ma_loai_hinh"].ToString().Trim();
                if ((cur_wh_id == wh_id) && (cur_ma_lh == ma_lh))
                {
                    DialogResult thongbao;

                    thongbao = (MessageBox.Show("Item was created.\n"
                                                + "Mã loại hình = " + ma_lh
                                                + "\nWH ID = " + wh_id
                                                + "\nDo you want to update ?", " Attention ",
                                                            MessageBoxButtons.YesNo, MessageBoxIcon.Warning));
                    if (thongbao == DialogResult.Yes)
                    {
                        row["ID"] = id;
                        row["Ma_loai_hinh"] = ma_lh;
                        row["WareHouse_ID"] = wh_id;
                        row["Import_Export"] = im_or_ex;
                        row["Ty_le"] = decimal.Parse(ty_le);
                        WH_ID_with_MLH_Store_BT_Click(null, null);
                    }
                    return;
                }
            }
            tyle = decimal.Parse(ty_le);
            DataRow newrow = WH_ID_with_MaLH_Table_Form.Data_dtb.NewRow();
            newrow["ID"] = id;
            newrow["Ma_loai_hinh"] = ma_lh;
            newrow["WareHouse_ID"] = wh_id;
            newrow["Import_Export"] = im_or_ex;
            newrow["Ty_le"] = tyle;
            WH_ID_with_MaLH_Table_Form.Data_dtb.Rows.Add(newrow);
        }

        private void WH_ID_with_MLH_Store_BT_Click(object sender, EventArgs e)
        {
            if (Update_SQL_Data(WH_ID_with_MaLH_Table_Form.Data_da, WH_ID_with_MaLH_Table_Form.Data_dtb) == true)
            {
                MessageBox.Show("Store Data Complete", "Successful");
            }
            else
            {
                MessageBox.Show("Store Data Fail", "Failed");
            }
        }
        
    }
}