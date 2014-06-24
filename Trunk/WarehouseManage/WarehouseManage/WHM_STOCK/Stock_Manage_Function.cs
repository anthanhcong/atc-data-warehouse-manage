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
        const int Stock_Manage_COL_NUM = 7;
        ExcelImportStruct[] Stock_Manage_Col = new ExcelImportStruct[Stock_Manage_COL_NUM];
        const int Stock_Manage_Ma_LH = 0;
        const int Stock_Manage_WH_ID = 1;
        const int Stock_Manage_Part_Number = 2;
        const int Stock_Manage_Bin = 3;
        const int Stock_Manage_Plant = 4;

        public DataTable Stock_Manage_Tbl;
        public DataSet Stock_Manage_Tbl_ds = new DataSet();
        public SqlDataAdapter Stock_Manage_Tbl_da;

        public string Stock_error_log;

        private void Stock_Manage_InitExcelCol_Infor()
        {
            Stock_Manage_Col[0] = new ExcelImportStruct("Ma_loai_hinh", "Ma_loai_hinh", Excel_Col_Type.COL_STRING, 20);
            Stock_Manage_Col[1] = new ExcelImportStruct("WareHouse_ID", "WareHouse_ID", Excel_Col_Type.COL_STRING, 30);
            Stock_Manage_Col[2] = new ExcelImportStruct("Part_Number", "Part_Number", Excel_Col_Type.COL_STRING, 50);
            Stock_Manage_Col[3] = new ExcelImportStruct("Bin", "Bin", Excel_Col_Type.COL_STRING, 50);
            Stock_Manage_Col[4] = new ExcelImportStruct("Plant", "Plant", Excel_Col_Type.COL_STRING, 50);
            Stock_Manage_Col[5] = new ExcelImportStruct("Qty", "Qty", Excel_Col_Type.COL_FLOAT, 10);
            Stock_Manage_Col[6] = new ExcelImportStruct("Description", "Description", Excel_Col_Type.COL_STRING, 200);
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
            string part_number, wh_id, bin, plant, ma_lh;
            bool exist_wh_id = false, exist_part = false, ret_save = true;
            string import_allow;
            string cell_str;


            ProgressBar1.Visible = true;
            StatusLabel.Text = "Loading File";
            row = 1;
            OpenWB = Open_excel_file(file_name, "");
            Error_log = "";
            if (Stock_Manage_Get_Col_info(OpenWB) == true)
            {
                Load_Form_Stock_Manage_Line();
                Load_Material_List();
                Load_WH_ID_List();
                row = 3;
                cell_str = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, 2, 20);
                while (cell_str != "")
                {
                    part_number = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, Stock_Manage_Col[Stock_Manage_Part_Number].Col, Stock_Manage_Col[Stock_Manage_Part_Number].Data_Max_len);
                    bin = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, Stock_Manage_Col[Stock_Manage_Bin].Col, Stock_Manage_Col[Stock_Manage_Bin].Data_Max_len);
                    plant = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, Stock_Manage_Col[Stock_Manage_Plant].Col, Stock_Manage_Col[Stock_Manage_Plant].Data_Max_len);
                    wh_id = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, Stock_Manage_Col[Stock_Manage_WH_ID].Col, Stock_Manage_Col[Stock_Manage_WH_ID].Data_Max_len);
                    ma_lh = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, Stock_Manage_Col[Stock_Manage_Ma_LH].Col, Stock_Manage_Col[Stock_Manage_Ma_LH].Data_Max_len);
                    // Kiem tra Line da co trong database chua
                    import_allow = Allow_import(ma_lh, wh_id);
                    if (import_allow == "Y")
                    {
                        exist_wh_id = Is_exist_WH_ID(wh_id);
                        exist_part = Is_exist_Part_number(part_number);

                        if (exist_part == false)
                        {
                            Error_log += "\nPart number: " + part_number + " is'nt exist in Material list Datatable.";
                        }
                        if (exist_wh_id == false)
                        {
                            Error_log += "\nWarehouse ID: " + wh_id + " is'nt exist in Warehouse List Datatable.";
                        }
                        if ((exist_part == true) && (exist_wh_id == true))
                        {
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
                        }
                        row++;
                        cell_str = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, 2, 20);
                        ProgressBar1.Value = row % 100;
                        StatusLabel.Text = "Loading File, Line: " + row.ToString();
                    }
                    else if ((import_allow == "N") || (import_allow == ""))
                    {
                        Error_log = "Please check 'Import_allow' of 'Ma_loai_hinh':" + ma_lh + " and 'WareHouse_ID': " + wh_id
                                                            + " if you want to import";
                        //MessageBox.Show("Please check 'Import_allow' of 'Ma_loai_hinh':" + ma_lh+ " and 'WareHouse_ID': " + wh_id 
                        //                                    + " if you want to import", "Error");
                        row++;
                        cell_str = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, 2, 20);
                        ProgressBar1.Value = row % 100;
                        StatusLabel.Text = "Loading File, Line: " + row.ToString();
                        ret_save = false;
                        //break;
                    }
                    else
                    {
                        Error_log = "Please import info 'Ma_loai_hinh':" + ma_lh + " and 'WareHouse_ID': " + wh_id
                                                            + " with WH_List_and_Material_List Tab";
                        //MessageBox.Show("Please import info 'Ma_loai_hinh':" + ma_lh+ " and 'WareHouse_ID': " + wh_id 
                        //                                    + " with WH_List_and_Material_List Tab", "Error");
                        row++;
                        cell_str = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, 2, 20);
                        ProgressBar1.Value = row % 100;
                        StatusLabel.Text = "Loading File, Line: " + row.ToString();
                        ret_save = false;
                        //break;
                    }
                }
                Close_WorkBook(OpenWB);
                // Store data
                if (ret_save == true)
                {
                    if (Error_log != "")
                    {
                        MessageBox.Show(Error_log, "Error");
                        ProgressBar1.Visible = false;
                        StatusLabel.Text = "DONE";
                    }
                    else
                    {
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
                }
                else
                {
                    ProgressBar1.Visible = false;
                    MessageBox.Show(Error_log, "Error");
                    StatusLabel.Text = "Import Failed";
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

        private string Allow_import(string ma_loaihinh, string warehouse_id)
        {
            string cur_ma_lh, cur_wh_id, cur_sub_wh_id, im_allow;
            string ret_val = "Nulldata";
            foreach (DataRow row in Load_WH_ID_List_Tbl.Rows)
            {
                cur_ma_lh = row["Ma_loai_hinh"].ToString().Trim();
                cur_wh_id = row["WareHouse_ID"].ToString().Trim();
                cur_sub_wh_id = row["Sub_wh_id"].ToString().Trim();
                im_allow = row["Import_allow"].ToString().Trim();
                if ((cur_ma_lh == ma_loaihinh) && (cur_wh_id == warehouse_id))
                {
                    ret_val = im_allow;
                }
            }
            return ret_val;
        }

        private void Load_Material_List()
        {
            if (Load_Ma_List_Tbl != null)
            {
                Load_Ma_List_Tbl.Clear();
            }
            Load_Ma_List_Tbl = Get_SQL_Data(Database_WHM_Stock_Con_Str, @"SELECT * FROM dbo.Material_List_tb", ref Load_Ma_List_Tbl_da, ref Load_Ma_List_Tbl_ds);
        }

        private void Load_WH_ID_List()
        {
            if (Load_WH_ID_List_Tbl != null)
            {
                Load_WH_ID_List_Tbl.Clear();
            }
            Load_WH_ID_List_Tbl = Get_SQL_Data(Database_WHM_Stock_Con_Str, @"SELECT * FROM dbo.Warehouse_List_tb", ref Load_WH_ID_List_Tbl_da, ref Load_WH_ID_List_Tbl_ds);
        }

        private bool Is_exist_WH_ID(string wh_id)
        {
            string filterExpression = "";

            filterExpression = "WareHouse_ID =" + "'" + wh_id + "'";
            DataRow[] rows = Load_WH_ID_List_Tbl.Select(filterExpression);
            if (rows.Length == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private bool Is_exist_Part_number(string part_number)
        {
            string filterExpression = "";

            filterExpression = "Part_Number =" + "'" + part_number + "'";
            DataRow[] rows = Load_Ma_List_Tbl.Select(filterExpression);
            if (rows.Length == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
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