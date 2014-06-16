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
        const int WH_List_Manage_COL_NUM = 2;
        ExcelImportStruct[] WH_List_Manage_Col = new ExcelImportStruct[WH_List_Manage_COL_NUM];
        const int WH_List_Manage_WH_List_ID = 0;
        const int WH_List_Manage_WH_List_Name = 1;

        public DataTable WH_List_Manage_Tbl;
        public DataSet WH_List_Manage_Tbl_ds = new DataSet();
        public SqlDataAdapter WH_List_Manage_Tbl_da;

        private void WH_List_Manage_InitExcelCol_Infor()
        {
            WH_List_Manage_Col[0] = new ExcelImportStruct("WareHouse_ID", "WareHouse_ID", Excel_Col_Type.COL_STRING, 30);
            WH_List_Manage_Col[1] = new ExcelImportStruct("WareHouse_Name", "WareHouse_Name", Excel_Col_Type.COL_STRING, 50);
        }

        private bool WH_List_Manage_Get_Col_info(Excel.Workbook cur_wbook)
        {
            int i, col, row;
            string cell_val;
            string error_log = "";
            bool error = false;

            for (i = 0; i < WH_List_Manage_COL_NUM; i++)
            {
                WH_List_Manage_Col[i].Col = 0;
            }

            row = 2;
            for (col = 1; col < 20; col++)
            {
                cell_val = Get_Excel_Line((Excel.Worksheet)cur_wbook.Sheets[1], row, col, 1);
                for (i = 0; i < WH_List_Manage_COL_NUM; i++)
                {
                    if (cell_val == WH_List_Manage_Col[i].Col_str)
                    {
                        WH_List_Manage_Col[i].Col = col;
                        break;
                    }
                }
            }

            for (i = 0; i < WH_List_Manage_COL_NUM; i++)
            {
                if (WH_List_Manage_Col[i].Col == 0)
                {
                    error_log += "Can not find Column:" + WH_List_Manage_Col[i].Col_str + "\n";
                    error = true;
                }
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

        private bool Import_WH_List_Manage_Table_in_file(string file_name)
        {
            int row;
            string wh_id;
            string cell_str;


            ProgressBar1.Visible = true;
            StatusLabel.Text = "Loading File";
            row = 1;
            OpenWB = Open_excel_file(file_name, "");
            Error_log = "";
            if (WH_List_Manage_Get_Col_info(OpenWB) == true)
            {
                Load_Form_WH_List_Manage_Line();
                row = 3;
                cell_str = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, 2, 20);
                while (cell_str != "")
                {
                    wh_id = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, WH_List_Manage_Col[WH_List_Manage_WH_List_ID].Col, WH_List_Manage_Col[WH_List_Manage_WH_List_ID].Data_Max_len);
                    // Kiem tra Line da co trong database chua
                    if (Is_exist_WH_List_Manage(wh_id) == true)
                    {
                        // Update for this row
                        Update_WH_List_Manage_Line(wh_id, (Excel.Worksheet)OpenWB.Sheets[1], row);
                    }
                    else
                    {
                        // Insert new row
                        Create_New_WH_List_Manage_Line(wh_id, (Excel.Worksheet)OpenWB.Sheets[1], row);
                    }
                    row++;
                    cell_str = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, 2, 20);
                    ProgressBar1.Value = row % 100;
                    StatusLabel.Text = "Loading File, Line: " + row.ToString();
                }

                Close_WorkBook(OpenWB);
                // Store data
                if (Update_SQL_Data(WH_ID_List_Table_Form.Data_da, WH_ID_List_Table_Form.Data_dtb) == true)
                {
                    ProgressBar1.Visible = false;
                    StatusLabel.Text = "DONE";
                    MessageBox.Show("Complete Import Data");
                }
                else
                {
                    ProgressBar1.Visible = false;
                    StatusLabel.Text = "Import Failed";
                    MessageBox.Show("Import Data Failed");
                    StatusLabel.Text = "DONE";
                }

            }
            else
            {
                Close_WorkBook(OpenWB);
                ProgressBar1.Visible = false;
                StatusLabel.Text = "Import Failed";
                MessageBox.Show(Error_log, "Error File");
                StatusLabel.Text = "DONE";
            }
            return true;
        }

        private void Load_Form_WH_List_Manage_Line()
        {
            if (WH_ID_List_Table_Form.Data_dtb != null)
            {
                WH_ID_List_Table_Form.Data_dtb.Clear();
            }

            //Load Data into Table and display in gridview
            WH_ID_List_Table_Form.Load_DataBase(Database_WHM_Stock_Con_Str, @"SELECT * FROM [WHM_STOCK_DB].[dbo].[Warehouse_List_tb]");
        }

        private bool Is_exist_WH_List_Manage(string wh_id)
        {
            string cur_wh_id;
            foreach (DataRow row in WH_ID_List_Table_Form.Data_dtb.Rows)
            {
                cur_wh_id = row[WH_List_Manage_Col[WH_List_Manage_WH_List_ID].DB_str].ToString().Trim();
                if (cur_wh_id == wh_id)
                {
                    return true;
                }
            }
            return false;
        }

        private bool Update_WH_List_Manage_Line(string wh_id, Excel.Worksheet xsheet, int row)
        {
            string cur_wh_id;
            Excel_Col_Type col_type;
            int i;

            foreach (DataRow dt_row in WH_ID_List_Table_Form.Data_dtb.Rows)
            {
                cur_wh_id = dt_row[WH_List_Manage_Col[WH_List_Manage_WH_List_ID].DB_str].ToString().Trim();
                if (cur_wh_id == wh_id)
                {
                    for (i = 0; i < WH_List_Manage_COL_NUM; i++)
                    {
                        if (i != WH_List_Manage_WH_List_ID)
                        {
                            col_type = WH_List_Manage_Col[i].Col_type;
                            switch (col_type)
                            {
                                case Excel_Col_Type.COL_STRING:
                                    dt_row[WH_List_Manage_Col[i].DB_str] = Get_Text_Cell(xsheet, row, WH_List_Manage_Col[i].Col, WH_List_Manage_Col[i].Data_Max_len);
                                    break;
                                case Excel_Col_Type.COL_INT:
                                    dt_row[WH_List_Manage_Col[i].DB_str] = Get_int_Cell(xsheet, row, WH_List_Manage_Col[i].Col);
                                    break;
                                case Excel_Col_Type.COL_FLOAT:
                                    dt_row[WH_List_Manage_Col[i].DB_str] = Get_float_Cell(xsheet, row, WH_List_Manage_Col[i].Col);
                                    break;
                                case Excel_Col_Type.COL_DATE:
                                    dt_row[WH_List_Manage_Col[i].DB_str] = Get_date_str_Cell(xsheet, row, WH_List_Manage_Col[i].Col);
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                    return true;
                }
            }
            return false;
        }

        private bool Create_New_WH_List_Manage_Line (string wh_id, Excel.Worksheet xsheet, int row)
        {
            DataRow new_row = WH_ID_List_Table_Form.Data_dtb.NewRow();
            Excel_Col_Type col_type;
            int i;

            new_row[WH_List_Manage_Col[WH_List_Manage_WH_List_ID].DB_str] = wh_id;
            for (i = 0; i < WH_List_Manage_COL_NUM; i++)
            {
                if (i != WH_List_Manage_WH_List_ID)
                {
                    col_type = WH_List_Manage_Col[i].Col_type;
                    switch (col_type)
                    {
                        case Excel_Col_Type.COL_STRING:
                            new_row[WH_List_Manage_Col[i].DB_str] = Get_Text_Cell(xsheet, row, WH_List_Manage_Col[i].Col, WH_List_Manage_Col[i].Data_Max_len);
                            break;
                        case Excel_Col_Type.COL_INT:
                            new_row[WH_List_Manage_Col[i].DB_str] = Get_int_Cell(xsheet, row, WH_List_Manage_Col[i].Col);
                            break;
                        case Excel_Col_Type.COL_FLOAT:
                            new_row[WH_List_Manage_Col[i].DB_str] = Get_float_Cell(xsheet, row, WH_List_Manage_Col[i].Col);
                            break;
                        case Excel_Col_Type.COL_DATE:
                            new_row[WH_List_Manage_Col[i].DB_str] = Get_date_str_Cell(xsheet, row, WH_List_Manage_Col[i].Col);
                            break;
                        default:
                            break;
                    }
                }
            }
            WH_ID_List_Table_Form.Data_dtb.Rows.Add(new_row);
            return true;
        }
    }
}