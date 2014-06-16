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
        const int Stock_Manage_COL_NUM = 6;
        ExcelImportStruct[] Stock_Manage_Col = new ExcelImportStruct[Stock_Manage_COL_NUM];
        const int Stock_Manage_WH_ID = 0;
        const int Stock_Manage_Part_Number = 1;
        const int Stock_Manage_Bin = 2;
        const int Stock_Manage_Plant = 3;

        public DataTable Stock_Manage_Tbl;
        public DataSet Stock_Manage_Tbl_ds = new DataSet();
        public SqlDataAdapter Stock_Manage_Tbl_da;

        private void Stock_Manage_InitExcelCol_Infor()
        {
            Stock_Manage_Col[0] = new ExcelImportStruct("WareHouse_ID", "WareHouse_ID", Excel_Col_Type.COL_STRING, 30);
            Stock_Manage_Col[1] = new ExcelImportStruct("Part_Number", "Part_Number", Excel_Col_Type.COL_STRING, 50);
            Stock_Manage_Col[2] = new ExcelImportStruct("Bin", "Bin", Excel_Col_Type.COL_STRING, 50);
            Stock_Manage_Col[3] = new ExcelImportStruct("Plant", "Plant", Excel_Col_Type.COL_STRING, 50);
            Stock_Manage_Col[4] = new ExcelImportStruct("Qty", "Qty", Excel_Col_Type.COL_FLOAT, 10);
            Stock_Manage_Col[5] = new ExcelImportStruct("Description", "Description", Excel_Col_Type.COL_STRING, 200);
        }

        private bool Stock_Manage_Get_Col_info(Excel.Workbook cur_wbook)
        {
            int i, col, row;
            string cell_val;
            string error_log = "";
            bool error = false;

            for (i = 0; i < Stock_Manage_COL_NUM; i++)
            {
                Stock_Manage_Col[i].Col = 0;
            }

            row = 2;
            for (col = 1; col < 20; col++)
            {
                cell_val = Get_Excel_Line((Excel.Worksheet)cur_wbook.Sheets[1], row, col, 1);
                for (i = 0; i < Stock_Manage_COL_NUM; i++)
                {
                    if (cell_val == Stock_Manage_Col[i].Col_str)
                    {
                        Stock_Manage_Col[i].Col = col;
                        break;
                    }
                }
            }

            for (i = 0; i < Stock_Manage_COL_NUM; i++)
            {
                if (Stock_Manage_Col[i].Col == 0)
                {
                    error_log += "Can not find Column:" + Stock_Manage_Col[i].Col_str + "\n";
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

        private bool Import_Stock_Manage_Table_in_file(string file_name)
        {
            int row;
            string part_number, wh_id, bin, plant;
            string cell_str;


            ProgressBar1.Visible = true;
            StatusLabel.Text = "Loading File";
            row = 1;
            OpenWB = Open_excel_file(file_name, "");
            Error_log = "";
            if (Stock_Manage_Get_Col_info(OpenWB) == true)
            {
                Load_Form_Stock_Manage_Line();
                row = 3;
                cell_str = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, 2, 20);
                while (cell_str != "")
                {
                    part_number = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, Stock_Manage_Col[Stock_Manage_Part_Number].Col, Stock_Manage_Col[Stock_Manage_Part_Number].Data_Max_len);
                    bin = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, Stock_Manage_Col[Stock_Manage_Bin].Col, Stock_Manage_Col[Stock_Manage_Bin].Data_Max_len);
                    plant = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, Stock_Manage_Col[Stock_Manage_Plant].Col, Stock_Manage_Col[Stock_Manage_Plant].Data_Max_len);
                    wh_id = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, Stock_Manage_Col[Stock_Manage_WH_ID].Col, Stock_Manage_Col[Stock_Manage_WH_ID].Data_Max_len);
                    // Kiem tra Line da co trong database chua
                    if (Is_exist_Stock_Manage(part_number, bin, plant, wh_id) == true)
                    {
                        // Update for this row
                        Update_Stock_Manage_Line(part_number, bin, plant, wh_id, (Excel.Worksheet)OpenWB.Sheets[1], row);
                    }
                    else
                    {
                        // Insert new row
                        Create_New_Stock_Manage_Line(part_number, bin, plant, wh_id, (Excel.Worksheet)OpenWB.Sheets[1], row);
                    }
                    row++;
                    cell_str = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, 2, 20);
                    ProgressBar1.Value = row % 100;
                    StatusLabel.Text = "Loading File, Line: " + row.ToString();
                }

                Close_WorkBook(OpenWB);
                // Store data
                if (Update_SQL_Data(Stock_Table_Form.Data_da, Stock_Table_Form.Data_dtb) == true)
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

        private void Load_Form_Stock_Manage_Line()
        {
            if (Stock_Table_Form.Data_dtb != null)
            {
                Stock_Table_Form.Data_dtb.Clear();
            }

            //Load Data into Table and display in gridview
            Stock_Table_Form.Load_DataBase(Database_WHM_Stock_Con_Str, @"SELECT * FROM [WHM_STOCK_DB].[dbo].[Material_Stock_tb]");
        }

        private bool Is_exist_Stock_Manage(string part_number, string bin, string plant, string wh_id)
        {
            string cur_part_number, cur_bin, cur_plant, cur_wh_id;
            foreach (DataRow row in Stock_Table_Form.Data_dtb.Rows)
            {
                cur_part_number = row[Stock_Manage_Col[Stock_Manage_Part_Number].DB_str].ToString().Trim();
                cur_bin = row[Stock_Manage_Col[Stock_Manage_Bin].DB_str].ToString().Trim();
                cur_plant = row[Stock_Manage_Col[Stock_Manage_Plant].DB_str].ToString().Trim();
                cur_wh_id = row[Stock_Manage_Col[Stock_Manage_WH_ID].DB_str].ToString().Trim();
                if ((cur_part_number == part_number) && (cur_bin == bin) && (cur_plant == plant) && (cur_wh_id == wh_id))
                {
                    return true;
                }
            }
            return false;
        }

        private bool Update_Stock_Manage_Line(string part_number, string bin, string plant, string wh_id, Excel.Worksheet xsheet, int row)
        {
            string cur_part_number, cur_bin, cur_plant, cur_wh_id;
            Excel_Col_Type col_type;
            int i;

            foreach (DataRow dt_row in Stock_Table_Form.Data_dtb.Rows)
            {
                cur_part_number = dt_row[Stock_Manage_Col[Stock_Manage_Part_Number].DB_str].ToString().Trim();
                cur_bin = dt_row[Stock_Manage_Col[Stock_Manage_Bin].DB_str].ToString().Trim();
                cur_plant = dt_row[Stock_Manage_Col[Stock_Manage_Plant].DB_str].ToString().Trim();
                cur_wh_id = dt_row[Stock_Manage_Col[Stock_Manage_WH_ID].DB_str].ToString().Trim();
                if ((cur_part_number == part_number) && (cur_bin == bin) && (cur_wh_id == wh_id))
                {
                    for (i = 0; i < Stock_Manage_COL_NUM; i++)
                    {
                        if ((i != Stock_Manage_Part_Number) && (i != Stock_Manage_Bin) && (i != Stock_Manage_Plant) && (i != Stock_Manage_WH_ID))
                        {
                            col_type = Stock_Manage_Col[i].Col_type;
                            switch (col_type)
                            {
                                case Excel_Col_Type.COL_STRING:
                                    dt_row[Stock_Manage_Col[i].DB_str] = Get_Text_Cell(xsheet, row, Stock_Manage_Col[i].Col, Stock_Manage_Col[i].Data_Max_len);
                                    break;
                                case Excel_Col_Type.COL_INT:
                                    dt_row[Stock_Manage_Col[i].DB_str] = Get_int_Cell(xsheet, row, Stock_Manage_Col[i].Col);
                                    break;
                                case Excel_Col_Type.COL_FLOAT:
                                    dt_row[Stock_Manage_Col[i].DB_str] = Get_float_Cell(xsheet, row, Stock_Manage_Col[i].Col);
                                    break;
                                case Excel_Col_Type.COL_DATE:
                                    dt_row[Stock_Manage_Col[i].DB_str] = Get_date_str_Cell(xsheet, row, Stock_Manage_Col[i].Col);
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

        private bool Create_New_Stock_Manage_Line(string part_number, string bin, string plant, string wh_id, Excel.Worksheet xsheet, int row)
        {
            DataRow new_row = Stock_Table_Form.Data_dtb.NewRow();
            Excel_Col_Type col_type;
            int i;

            new_row[Stock_Manage_Col[Stock_Manage_Part_Number].DB_str] = part_number;
            new_row[Stock_Manage_Col[Stock_Manage_Bin].DB_str] = bin;
            new_row[Stock_Manage_Col[Stock_Manage_Plant].DB_str] = plant;
            new_row[Stock_Manage_Col[Stock_Manage_WH_ID].DB_str] = wh_id;
            for (i = 0; i < Stock_Manage_COL_NUM; i++)
            {
                if ((i != Stock_Manage_Part_Number) && (i != Stock_Manage_Bin) && (i != Stock_Manage_Plant) && (i != Stock_Manage_WH_ID))
                {
                    col_type = Stock_Manage_Col[i].Col_type;
                    switch (col_type)
                    {
                        case Excel_Col_Type.COL_STRING:
                            new_row[Stock_Manage_Col[i].DB_str] = Get_Text_Cell(xsheet, row, Stock_Manage_Col[i].Col, Stock_Manage_Col[i].Data_Max_len);
                            break;
                        case Excel_Col_Type.COL_INT:
                            new_row[Stock_Manage_Col[i].DB_str] = Get_int_Cell(xsheet, row, Stock_Manage_Col[i].Col);
                            break;
                        case Excel_Col_Type.COL_FLOAT:
                            new_row[Stock_Manage_Col[i].DB_str] = Get_float_Cell(xsheet, row, Stock_Manage_Col[i].Col);
                            break;
                        case Excel_Col_Type.COL_DATE:
                            new_row[Stock_Manage_Col[i].DB_str] = Get_date_str_Cell(xsheet, row, Stock_Manage_Col[i].Col);
                            break;
                        default:
                            break;
                    }
                }
            }
            Stock_Table_Form.Data_dtb.Rows.Add(new_row);
            return true;
        }    
    }
}