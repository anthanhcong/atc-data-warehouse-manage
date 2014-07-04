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
        private void Mother_Stock_Table_Form_CellDoubleClick(object sender, EventArgs e)
        {

        }

        private void Mother_Stock_Part_Number_cbx_SelectedValueChanged(Object sender, EventArgs e)
        {
            decimal total_qty = 0;
            Mother_WH_List_WH_with_part();
            foreach (DataRow row in List_WH_wiht_part_TBL.Rows)
            {
                total_qty += decimal.Parse(row["Qty"].ToString().Trim());
            }
            Mother_Stock_Sum_Qty_Txt.My_TextBox.Text = total_qty.ToString();
        }
        
        private void Mother_Stock_WH_ID_List_cbx_SelectedValueChanged(Object sender, EventArgs e)
        { }

        private void Mother_Stock_With_wh_id_cbx_SelectedValueChanged(Object sender, EventArgs e)
        {
            if (Mother_Stock_Table_Form.Data_dtb != null)
            {
                Mother_Stock_Table_Form.Data_dtb.Clear();
            }
            Mother_Stock_Table_Form.Data_dtb = Load_Form_List_WH_with_part(ref Mother_Stock_Table_Form.Data_ds, ref Mother_Stock_Table_Form.Data_da);
            
        }

        private void Mother_Stock_Part_Number_cbx_KeyDown(object sender, KeyEventArgs e)
        { }

        private void Mother_Stock_WH_ID_List_cbx_KeyDown(object sender, KeyEventArgs e)
        { }

        private void Mother_Stock_Bin_Txt_KeyDown(object sender, KeyEventArgs e)
        { }

        private void Mother_Stock_Plant_Txt_KeyDown(object sender, KeyEventArgs e)
        { }

        private void Mother_Stock_With_wh_id_cbx_KeyDown(object sender, KeyEventArgs e)
        { }

        private void Mother_Stock_Search_BT_Click(object sender, EventArgs e)
        {

        }

        private void Mother_Stock_Store_BT_Click(object sender, EventArgs e)
        {

        }
    }
}