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
        public DataTable Warehouse_Stock_Load_Table(string load_cmd, ref DataSet ds, ref SqlDataAdapter da)
        {
            DataTable table;

            table = Get_SQL_Data(Database_WHM_Stock_Con_Str, load_cmd, ref da, ref ds);
            return table;
        }

        public void  WH_Dashboard_Load_Table()
        {
            string load_table_cmd = @"SELECT * FROM [WHM_STOCK_DB].[dbo].[Warehouse_Dashboard_tb]";

            if (Load_WH_Daskboard_TBL != null)
            {
                Load_WH_Daskboard_TBL.Clear();
            }
            Load_WH_Daskboard_TBL = Warehouse_Stock_Load_Table(load_table_cmd, ref Load_WH_Daskboard_ds, ref Load_WH_Daskboard_da);
        }

        public void WH_Dashboard_Load_WH_ID(string wh_id)
        {
            //distinct[WareHouse_ID]
            string load_table_cmd = @"SELECT * FROM [WHM_STOCK_DB].[dbo].[Warehouse_Dashboard_tb]";
            if (wh_id != "All")
            {
                load_table_cmd += " where WareHouse_ID =" + "'" + wh_id + "'";
            }
            if (WH_Daskboard_Table_Form.Data_dtb != null)
            {
                WH_Daskboard_Table_Form.Data_dtb.Clear();
            }
            WH_Daskboard_Table_Form.Load_DataBase(Database_WHM_Stock_Con_Str, load_table_cmd);
        }

        public void WH_Dashboard_Load_with_Mother_wh(string mother_wh)
        {
            string load_table_cmd = @"SELECT * FROM [WHM_STOCK_DB].[dbo].[Warehouse_Dashboard_tb]";
            load_table_cmd += " where WareHouse_ID =" + "'" + mother_wh + "'";
            load_table_cmd += " or Mother_WHID =" + "'" + mother_wh + "'";

            if (WH_Daskboard_Table_Form.Data_dtb != null)
            {
                WH_Daskboard_Table_Form.Data_dtb.Clear();
            }
            WH_Daskboard_Table_Form.Load_DataBase(Database_WHM_Stock_Con_Str, load_table_cmd);
        }

        public void WH_ID_with_MaLH_Load_All()
        {
            string load_table_cmd = @"SELECT * FROM [WHM_STOCK_DB].[dbo].[List_WH_ID_MaLH_tb]";

            if (WH_ID_with_MaLH_Table_Form.Data_dtb != null)
            {
                WH_ID_with_MaLH_Table_Form.Data_dtb.Clear();
            }
            WH_ID_with_MaLH_Table_Form.Load_DataBase(Database_WHM_Stock_Con_Str, load_table_cmd);
        }

        public void WH_ID_with_MaLH_Load_with_WH_ID( string wh_id)
        {
            string load_table_cmd = @"SELECT * FROM [WHM_STOCK_DB].[dbo].[List_WH_ID_MaLH_tb]";
            load_table_cmd += "where WareHouse_ID = " + "'" + wh_id + "'";

            if (WH_ID_with_MaLH_Table_Form.Data_dtb != null)
            {
                WH_ID_with_MaLH_Table_Form.Data_dtb.Clear();
            }
            WH_ID_with_MaLH_Table_Form.Load_DataBase(Database_WHM_Stock_Con_Str, load_table_cmd);
        }
    }
}