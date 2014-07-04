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
        public DataTable Mother_WH_Load_Stock_TBL;
        public DataSet Mother_WH_Load_Stock_ds = new DataSet();
        public SqlDataAdapter Mother_WH_Load_Stock_da;

        public DataTable List_WH_wiht_part_TBL;
        public DataSet List_WH_wiht_part_ds = new DataSet();
        public SqlDataAdapter List_WH_wiht_part_da;

        public DataTable Mother_WH_List_Part_Tbl;
        public DataSet Mother_WH_List_Part_Tbl_ds = new DataSet();
        public SqlDataAdapter Mother_WH_List_Part_Tbl_da;

        public DataTable Stock_Total_Qty_tbl;

        private void Calc_Total_Qty_Part()
        { 
            //Stock_Total_Qty_tbl = 
        }

        private void Mother_WH_Load_Material_Stock()
        {
            string load_cmd = @"SELECT * FROM [WHM_STOCK_DB].[dbo].[Material_Stock_tb]";
            if (Mother_WH_Load_Stock_TBL != null)
            {
                Mother_WH_Load_Stock_TBL.Clear();
            }
            Mother_WH_Load_Stock_TBL = Get_SQL_Data(Database_WHM_Stock_Con_Str, load_cmd, ref Mother_WH_Load_Stock_da, ref Mother_WH_Load_Stock_ds);
        }

        private void Mother_WH_Load_Part_Number()
        {
            string load_cmd = @"SELECT distinct[Part_Number] FROM [WHM_STOCK_DB].[dbo].[Material_Stock_tb]";
            if (Mother_WH_List_Part_Tbl != null)
            {
                Mother_WH_List_Part_Tbl.Clear();
            }
            Mother_WH_List_Part_Tbl = Get_SQL_Data(Database_WHM_Stock_Con_Str, load_cmd, ref Mother_WH_List_Part_Tbl_da, ref Mother_WH_List_Part_Tbl_ds);
        }

        private void Mother_WH_List_WH_with_part()
        {
            string part_number = "";

            if (Mother_Stock_Part_Number_cbx.My_Combo.SelectedValue != null)
            {
                part_number = Mother_Stock_Part_Number_cbx.My_Combo.SelectedValue.ToString().Trim();
                
            }
            string load_cmd = @"SELECT * FROM [WHM_STOCK_DB].[dbo].[Material_Stock_tb]";
            load_cmd += " WHERE Part_Number = '" + part_number + "'";

            if (List_WH_wiht_part_TBL != null)
            {
                List_WH_wiht_part_TBL.Clear();
            }
            List_WH_wiht_part_TBL = Get_SQL_Data(Database_WHM_Stock_Con_Str, load_cmd, ref List_WH_wiht_part_da, ref List_WH_wiht_part_ds);

            if (Mother_Stock_Table_Form.Data_dtb != null)
            {
                Mother_Stock_Table_Form.Data_dtb.Clear();
            }
            Mother_Stock_Table_Form.Load_DataBase(Database_WHM_Stock_Con_Str, @"SELECT * FROM dbo.Material_Stock_tb where Part_Number = '" + part_number + "'");
        }

        private DataTable Load_Form_List_WH_with_part(ref DataSet ds, ref SqlDataAdapter da)
        {
            DataTable table;
            string part_number = "", wh_id = "";

            if (Mother_Stock_Part_Number_cbx.My_Combo.SelectedValue != null)
            {
                part_number = Mother_Stock_Part_Number_cbx.My_Combo.SelectedValue.ToString().Trim();

            }
            if (Mother_Stock_With_wh_id_cbx.My_Combo.SelectedValue != null)
            {
                wh_id = Mother_Stock_With_wh_id_cbx.My_Combo.SelectedValue.ToString().Trim();

            }
            string load_cmd = @"SELECT * FROM [WHM_STOCK_DB].[dbo].[Material_Stock_tb]";
            load_cmd += " WHERE Part_Number = '" + part_number + "'";
            load_cmd += " AND WareHouse_ID = '" + wh_id + "'";

            table = Get_SQL_Data(Database_WHM_Stock_Con_Str, load_cmd, ref da, ref ds);

            return table;
        }
    }
}