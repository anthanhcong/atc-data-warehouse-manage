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
        //const int WH_List_Manage_COL_NUM = 5;
        //ExcelImportStruct[] WH_List_Manage_Col = new ExcelImportStruct[WH_List_Manage_COL_NUM];
        //const int WH_List_Manage_WH_Ma_Loaihinh= 0;
        //const int WH_List_Manage_WH_List_ID = 1;
        //const int WH_List_Manage_WH_List_Name = 2;
        //const int WH_List_Manage_WH_Fa_wh = 3;
        //const int WH_List_Manage_WH_Sub_wh = 4;

        //public DataTable WH_List_Manage_Tbl;
        //public DataSet WH_List_Manage_Tbl_ds = new DataSet();
        //public SqlDataAdapter WH_List_Manage_Tbl_da;

        //private void WH_List_Manage_InitExcelCol_Infor()
        //{
        //    WH_List_Manage_Col[0] = new ExcelImportStruct("Ma_loai_hinh", "Ma_loai_hinh", Excel_Col_Type.COL_STRING, 20);
        //    WH_List_Manage_Col[1] = new ExcelImportStruct("WareHouse_ID", "WareHouse_ID", Excel_Col_Type.COL_STRING, 30);
        //    WH_List_Manage_Col[2] = new ExcelImportStruct("WareHouse_Name", "WareHouse_Name", Excel_Col_Type.COL_STRING, 50);
        //    WH_List_Manage_Col[3] = new ExcelImportStruct("Fa_wh_id", "Fa_wh_id", Excel_Col_Type.COL_STRING, 30);
        //    WH_List_Manage_Col[4] = new ExcelImportStruct("Sub_wh_id", "Sub_wh_id", Excel_Col_Type.COL_STRING, 30);
        //}

        //private bool WH_List_Manage_Get_Col_info(Excel.Workbook cur_wbook)
        //{
        //    int i, col, row;
        //    string cell_val;
        //    string error_log = "";
        //    bool error = false;

        //    for (i = 0; i < WH_List_Manage_COL_NUM; i++)
        //    {
        //        WH_List_Manage_Col[i].Col = 0;
        //    }

        //    row = 2;
        //    for (col = 1; col < 20; col++)
        //    {
        //        cell_val = Get_Excel_Line((Excel.Worksheet)cur_wbook.Sheets[1], row, col, 1);
        //        for (i = 0; i < WH_List_Manage_COL_NUM; i++)
        //        {
        //            if (cell_val == WH_List_Manage_Col[i].Col_str)
        //            {
        //                WH_List_Manage_Col[i].Col = col;
        //                break;
        //            }
        //        }
        //    }

        //    for (i = 0; i < WH_List_Manage_COL_NUM; i++)
        //    {
        //        if (WH_List_Manage_Col[i].Col == 0)
        //        {
        //            error_log += "Can not find Column:" + WH_List_Manage_Col[i].Col_str + "\n";
        //            error = true;
        //        }
        //    }

        //    if (error == true)
        //    {

        //        Error_log = error_log;
        //        return false;
        //    }
        //    else
        //    {
        //        Error_log = "";
        //        return true;
        //    }

        //}

        //private bool Import_WH_List_Manage_Table_in_file(string file_name)
        //{
        //    int row;
        //    string ma_loaihinh, wh_id, sub_wh;
        //    string cell_str;


        //    ProgressBar1.Visible = true;
        //    StatusLabel.Text = "Loading File";
        //    row = 1;
        //    OpenWB = Open_excel_file(file_name, "");
        //    Error_log = "";
        //    if (WH_List_Manage_Get_Col_info(OpenWB) == true)
        //    {
        //        Load_Form_WH_List_Manage_Line();
        //        row = 3;
        //        cell_str = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, 2, 20);
        //        while (cell_str != "")
        //        {
        //            wh_id = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, WH_List_Manage_Col[WH_List_Manage_WH_List_ID].Col, WH_List_Manage_Col[WH_List_Manage_WH_List_ID].Data_Max_len);
        //            ma_loaihinh = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, WH_List_Manage_Col[WH_List_Manage_WH_Ma_Loaihinh].Col, WH_List_Manage_Col[WH_List_Manage_WH_Ma_Loaihinh].Data_Max_len);
        //            sub_wh = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, WH_List_Manage_Col[WH_List_Manage_WH_Sub_wh].Col, WH_List_Manage_Col[WH_List_Manage_WH_Sub_wh].Data_Max_len);
        //            if (sub_wh == "")
        //            {
        //                sub_wh = "Blank";
        //            }
        //            // Kiem tra Line da co trong database chua
        //            if (Is_exist_WH_List_Manage(wh_id, ma_loaihinh, sub_wh) == true)
        //            {
        //                // Update for this row
        //                if (WH_List_Import_Auto_Update.My_CheckBox.Checked == true)
        //                {
        //                    Update_WH_List_Manage_Line(wh_id, ma_loaihinh, sub_wh, (Excel.Worksheet)OpenWB.Sheets[1], row);
        //                }
        //                else
        //                {
        //                    DialogResult thongbao;

        //                    thongbao = (MessageBox.Show("Warehouse ID: " + wh_id +
        //                                                "\nMã loại hình: " + ma_loaihinh + 
        //                                                "\nSub WH: " + sub_wh + " was created.\n"
        //                                                 + "Do you want to update ?", " Attention ",
        //                                                            MessageBoxButtons.YesNo, MessageBoxIcon.Warning));
        //                    if (thongbao == DialogResult.Yes)
        //                    {
        //                        Update_WH_List_Manage_Line(wh_id, ma_loaihinh, sub_wh, (Excel.Worksheet)OpenWB.Sheets[1], row);
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                // Insert new row
        //                Create_New_WH_List_Manage_Line(wh_id, ma_loaihinh, sub_wh, (Excel.Worksheet)OpenWB.Sheets[1], row);
        //            }
        //            row++;
        //            cell_str = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, 2, 20);
        //            ProgressBar1.Value = row % 100;
        //            StatusLabel.Text = "Loading File, Line: " + row.ToString();
        //        }

        //        Close_WorkBook(OpenWB);
        //        // Store data
        //        if (Update_SQL_Data(WH_ID_List_Table_Form.Data_da, WH_ID_List_Table_Form.Data_dtb) == true)
        //        {
        //            ProgressBar1.Visible = false;
        //            StatusLabel.Text = "DONE";
        //            MessageBox.Show("Complete Import Data");
        //        }
        //        else
        //        {
        //            ProgressBar1.Visible = false;
        //            StatusLabel.Text = "Import Failed";
        //            MessageBox.Show("Import Data Failed");
        //            StatusLabel.Text = "DONE";
        //        }

        //    }
        //    else
        //    {
        //        Close_WorkBook(OpenWB);
        //        ProgressBar1.Visible = false;
        //        StatusLabel.Text = "Import Failed";
        //        MessageBox.Show(Error_log, "Error File");
        //        StatusLabel.Text = "DONE";
        //    }
        //    return true;
        //}

        //private void Load_Form_WH_List_Manage_Line()
        //{
        //    if (WH_ID_List_Table_Form.Data_dtb != null)
        //    {
        //        WH_ID_List_Table_Form.Data_dtb.Clear();
        //    }

        //    //Load Data into Table and display in gridview
        //    WH_ID_List_Table_Form.Load_DataBase(Database_WHM_Stock_Con_Str, @"SELECT * FROM [WHM_STOCK_DB].[dbo].[Warehouse_List_tb]");
        //}

        //private bool Is_exist_WH_List_Manage(string wh_id, string ma_loaihinh, string sub_wh)
        //{
        //    string cur_wh_id, cur_ma_lh, cur_sub_wh;
        //    foreach (DataRow row in WH_ID_List_Table_Form.Data_dtb.Rows)
        //    {
        //        cur_wh_id = row[WH_List_Manage_Col[WH_List_Manage_WH_List_ID].DB_str].ToString().Trim();
        //        cur_ma_lh = row[WH_List_Manage_Col[WH_List_Manage_WH_Ma_Loaihinh].DB_str].ToString().Trim();
        //        cur_sub_wh = row[WH_List_Manage_Col[WH_List_Manage_WH_Sub_wh].DB_str].ToString().Trim();
        //        if ((cur_wh_id == wh_id) && (cur_ma_lh == ma_loaihinh) && (cur_sub_wh == sub_wh))
        //        {
        //            return true;
        //        }
        //    }
        //    return false;
        //}

        //private bool Update_WH_List_Manage_Line(string wh_id, string ma_loaihinh, string sub_wh, Excel.Worksheet xsheet, int row)
        //{
        //    string cur_wh_id, cur_ma_lh, cur_sub_wh;
        //    Excel_Col_Type col_type;
        //    int i;

        //    foreach (DataRow dt_row in WH_ID_List_Table_Form.Data_dtb.Rows)
        //    {
        //        cur_wh_id = dt_row[WH_List_Manage_Col[WH_List_Manage_WH_List_ID].DB_str].ToString().Trim();
        //        cur_ma_lh = dt_row[WH_List_Manage_Col[WH_List_Manage_WH_Ma_Loaihinh].DB_str].ToString().Trim();
        //        cur_sub_wh = dt_row[WH_List_Manage_Col[WH_List_Manage_WH_Sub_wh].DB_str].ToString().Trim();

        //        if ((cur_wh_id == wh_id) && (cur_ma_lh == ma_loaihinh) && (cur_sub_wh == sub_wh))
        //        {
        //            for (i = 0; i < WH_List_Manage_COL_NUM; i++)
        //            {
        //                if ((i != WH_List_Manage_WH_List_ID) && (i != WH_List_Manage_WH_Ma_Loaihinh) && (i != WH_List_Manage_WH_Sub_wh))
        //                {
        //                    col_type = WH_List_Manage_Col[i].Col_type;
        //                    switch (col_type)
        //                    {
        //                        case Excel_Col_Type.COL_STRING:
        //                            dt_row[WH_List_Manage_Col[i].DB_str] = Get_Text_Cell(xsheet, row, WH_List_Manage_Col[i].Col, WH_List_Manage_Col[i].Data_Max_len);
        //                            break;
        //                        case Excel_Col_Type.COL_INT:
        //                            dt_row[WH_List_Manage_Col[i].DB_str] = Get_int_Cell(xsheet, row, WH_List_Manage_Col[i].Col);
        //                            break;
        //                        case Excel_Col_Type.COL_FLOAT:
        //                            dt_row[WH_List_Manage_Col[i].DB_str] = Get_float_Cell(xsheet, row, WH_List_Manage_Col[i].Col);
        //                            break;
        //                        case Excel_Col_Type.COL_DATE:
        //                            dt_row[WH_List_Manage_Col[i].DB_str] = Get_date_str_Cell(xsheet, row, WH_List_Manage_Col[i].Col);
        //                            break;
        //                        default:
        //                            break;
        //                    }
        //                }
        //            }
        //            return true;
        //        }
        //    }
        //    return false;
        //}

        //private bool Create_New_WH_List_Manage_Line (string wh_id, string ma_loaihinh, string sub_wh, Excel.Worksheet xsheet, int row)
        //{
        //    DataRow new_row = WH_ID_List_Table_Form.Data_dtb.NewRow();
        //    Excel_Col_Type col_type;
        //    int i;

        //    new_row[WH_List_Manage_Col[WH_List_Manage_WH_List_ID].DB_str] = wh_id;
        //    new_row[WH_List_Manage_Col[WH_List_Manage_WH_Ma_Loaihinh].DB_str] = ma_loaihinh;
        //    new_row[WH_List_Manage_Col[WH_List_Manage_WH_Sub_wh].DB_str] = sub_wh;
        //    new_row["Import_allow"] = "N";
        //    for (i = 0; i < WH_List_Manage_COL_NUM; i++)
        //    {
        //        if ((i != WH_List_Manage_WH_List_ID) && (i != WH_List_Manage_WH_Ma_Loaihinh) && (i != WH_List_Manage_WH_Sub_wh))
        //        {
        //            col_type = WH_List_Manage_Col[i].Col_type;
        //            switch (col_type)
        //            {
        //                case Excel_Col_Type.COL_STRING:
        //                    new_row[WH_List_Manage_Col[i].DB_str] = Get_Text_Cell(xsheet, row, WH_List_Manage_Col[i].Col, WH_List_Manage_Col[i].Data_Max_len);
        //                    break;
        //                case Excel_Col_Type.COL_INT:
        //                    new_row[WH_List_Manage_Col[i].DB_str] = Get_int_Cell(xsheet, row, WH_List_Manage_Col[i].Col);
        //                    break;
        //                case Excel_Col_Type.COL_FLOAT:
        //                    new_row[WH_List_Manage_Col[i].DB_str] = Get_float_Cell(xsheet, row, WH_List_Manage_Col[i].Col);
        //                    break;
        //                case Excel_Col_Type.COL_DATE:
        //                    new_row[WH_List_Manage_Col[i].DB_str] = Get_date_str_Cell(xsheet, row, WH_List_Manage_Col[i].Col);
        //                    break;
        //                default:
        //                    break;
        //            }
        //        }
        //    }
        //    WH_ID_List_Table_Form.Data_dtb.Rows.Add(new_row);
        //    return true;
        //}

        /***********************************/

        const int WH_List_Manage_COL_NUM = 4;
        ExcelImportStruct[] WH_List_Manage_Col = new ExcelImportStruct[WH_List_Manage_COL_NUM];
        const int WH_List_Manage_WH_Ma_Loaihinh = 0;
        const int WH_List_Manage_WH_List_ID = 1;
        const int WH_List_Manage_WH_List_Name = 2;
        const int WH_List_Manage_WH_Fa_wh = 3;
        //const int WH_List_Manage_WH_Sub_wh = 4;

        public DataTable WH_List_Manage_Tbl;
        public DataSet WH_List_Manage_Tbl_ds = new DataSet();
        public SqlDataAdapter WH_List_Manage_Tbl_da;

        private void WH_List_Manage_InitExcelCol_Infor()
        {
            WH_List_Manage_Col[0] = new ExcelImportStruct("Ma_loai_hinh", "Ma_loai_hinh", Excel_Col_Type.COL_STRING, 20);
            WH_List_Manage_Col[1] = new ExcelImportStruct("WareHouse_ID", "WareHouse_ID", Excel_Col_Type.COL_STRING, 30);
            WH_List_Manage_Col[2] = new ExcelImportStruct("WareHouse_Name", "WareHouse_Name", Excel_Col_Type.COL_STRING, 50);
            WH_List_Manage_Col[3] = new ExcelImportStruct("Fa_wh_id", "Fa_wh_id", Excel_Col_Type.COL_STRING, 30);
            //WH_List_Manage_Col[4] = new ExcelImportStruct("Sub_wh_id", "Sub_wh_id", Excel_Col_Type.COL_STRING, 30);
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
            string ma_loaihinh, wh_id, fa_wh;
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
                    ma_loaihinh = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, WH_List_Manage_Col[WH_List_Manage_WH_Ma_Loaihinh].Col, WH_List_Manage_Col[WH_List_Manage_WH_Ma_Loaihinh].Data_Max_len);
                    fa_wh = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, WH_List_Manage_Col[WH_List_Manage_WH_Fa_wh].Col, WH_List_Manage_Col[WH_List_Manage_WH_Fa_wh].Data_Max_len);
                    // Kiem tra Line da co trong database chua
                    if (Is_exist_WH_List_Manage(wh_id, ma_loaihinh) == true)
                    {
                        // Update for this row
                        if (WH_List_Import_Auto_Update.My_CheckBox.Checked == true)
                        {
                            Update_WH_List_Manage_Line(wh_id, ma_loaihinh, (Excel.Worksheet)OpenWB.Sheets[1], row);
                        }
                        else
                        {
                            DialogResult thongbao;

                            thongbao = (MessageBox.Show("Warehouse ID: " + wh_id +
                                                        "\nMã loại hình: " + ma_loaihinh +
                                                        "\nFa WH: " + fa_wh + " was created.\n"
                                                         + "Do you want to update ?", " Attention ",
                                                                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning));
                            if (thongbao == DialogResult.Yes)
                            {
                                Update_WH_List_Manage_Line(wh_id, ma_loaihinh, (Excel.Worksheet)OpenWB.Sheets[1], row);
                            }
                        }
                    }
                    else
                    {
                        // Insert new row
                        Create_New_WH_List_Manage_Line(wh_id, ma_loaihinh, (Excel.Worksheet)OpenWB.Sheets[1], row);
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

        private bool Is_exist_WH_List_Manage(string wh_id, string ma_loaihinh)
        {
            string cur_wh_id, cur_ma_lh;
            foreach (DataRow row in WH_ID_List_Table_Form.Data_dtb.Rows)
            {
                cur_wh_id = row[WH_List_Manage_Col[WH_List_Manage_WH_List_ID].DB_str].ToString().Trim();
                cur_ma_lh = row[WH_List_Manage_Col[WH_List_Manage_WH_Ma_Loaihinh].DB_str].ToString().Trim();
                if ((cur_wh_id == wh_id) && (cur_ma_lh == ma_loaihinh))
                {
                    return true;
                }
            }
            return false;
        }

        private bool Update_WH_List_Manage_Line(string wh_id, string ma_loaihinh, Excel.Worksheet xsheet, int row)
        {
            string cur_wh_id, cur_ma_lh;
            Excel_Col_Type col_type;
            int i;

            foreach (DataRow dt_row in WH_ID_List_Table_Form.Data_dtb.Rows)
            {
                cur_wh_id = dt_row[WH_List_Manage_Col[WH_List_Manage_WH_List_ID].DB_str].ToString().Trim();
                cur_ma_lh = dt_row[WH_List_Manage_Col[WH_List_Manage_WH_Ma_Loaihinh].DB_str].ToString().Trim();
                dt_row["Import_allow"] = "N";
                if ((cur_wh_id == wh_id) && (cur_ma_lh == ma_loaihinh))
                {
                    for (i = 0; i < WH_List_Manage_COL_NUM; i++)
                    {
                        if ((i != WH_List_Manage_WH_List_ID) && (i != WH_List_Manage_WH_Ma_Loaihinh))
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

        private bool Create_New_WH_List_Manage_Line(string wh_id, string ma_loaihinh, Excel.Worksheet xsheet, int row)
        {
            DataRow new_row = WH_ID_List_Table_Form.Data_dtb.NewRow();
            Excel_Col_Type col_type;
            int i;

            new_row[WH_List_Manage_Col[WH_List_Manage_WH_List_ID].DB_str] = wh_id;
            new_row[WH_List_Manage_Col[WH_List_Manage_WH_Ma_Loaihinh].DB_str] = ma_loaihinh;
            new_row["Import_allow"] = "N";
            for (i = 0; i < WH_List_Manage_COL_NUM; i++)
            {
                if ((i != WH_List_Manage_WH_List_ID) && (i != WH_List_Manage_WH_Ma_Loaihinh))
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