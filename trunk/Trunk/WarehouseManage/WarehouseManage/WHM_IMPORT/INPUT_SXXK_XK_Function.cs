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
using ConvertDB;

namespace WarehouseManager
{
    public partial class Form1
    {
        const int INPUT_SXXK_XK_COL_NUM = 23;
        ExcelImportStruct[] INPUT_SXXK_XK_Col = new ExcelImportStruct[INPUT_SXXK_XK_COL_NUM];
        const int INPUT_SXXK_XK_So_TK = 0;
        const int INPUT_SXXK_XK_Ngay_DK = 1;
        const int INPUT_SXXK_XK_Ma_loai_hinh = 2;
        const int INPUT_SXXK_XK_Ma_HS = 18;

        public DataTable INPUT_SXXK_XK_Tbl;
        public DataSet INPUT_SXXK_XK_Tbl_ds = new DataSet();
        public SqlDataAdapter INPUT_SXXK_XK_Tbl_da;

        public DataTable INPUT_XK_List_TK_Tbl;
        public DataSet INPUT_XK_List_TK_ds = new DataSet();
        public SqlDataAdapter INPUT_XK_List_TK_da;


        private void INPUT_SXXK_XK_InitExcelCol_Infor()
        {
            INPUT_SXXK_XK_Col[0]  = new ExcelImportStruct("So_TK", "Số TK", Excel_Col_Type.COL_STRING, 20);
            INPUT_SXXK_XK_Col[1]  = new ExcelImportStruct("Ngay_DK", "Ngày ĐK", Excel_Col_Type.COL_DATE, 20);
            INPUT_SXXK_XK_Col[2]  = new ExcelImportStruct("Ma_loai_hinh", "Mã loại hình", Excel_Col_Type.COL_STRING, 30);
            INPUT_SXXK_XK_Col[3]  = new ExcelImportStruct("So_HD", "Số HĐ", Excel_Col_Type.COL_STRING, 100);
            INPUT_SXXK_XK_Col[4]  = new ExcelImportStruct("So_hoa_don_TM", "Số hóa đơn TM", Excel_Col_Type.COL_STRING, 100);
            INPUT_SXXK_XK_Col[5]  = new ExcelImportStruct("Nuoc_nhap", "Nước nhập", Excel_Col_Type.COL_STRING, 50);
            INPUT_SXXK_XK_Col[6]  = new ExcelImportStruct("Ma_cua_khau", "Mã cửa khẩu", Excel_Col_Type.COL_STRING, 100);
            INPUT_SXXK_XK_Col[7]  = new ExcelImportStruct("Ma_giao_hang", "Mã giao hàng", Excel_Col_Type.COL_STRING, 20);
            INPUT_SXXK_XK_Col[8]  = new ExcelImportStruct("Nguyen_te", "Nguyên tệ", Excel_Col_Type.COL_STRING, 20);
            INPUT_SXXK_XK_Col[9]  = new ExcelImportStruct("Ty_gia_VND", "Tỷ giá VNĐ", Excel_Col_Type.COL_FLOAT, 20);
            INPUT_SXXK_XK_Col[10] = new ExcelImportStruct("So_kien", "Số kiện", Excel_Col_Type.COL_INT, 20);
            INPUT_SXXK_XK_Col[11] = new ExcelImportStruct("Cont_20", "Cont 20", Excel_Col_Type.COL_INT, 10);
            INPUT_SXXK_XK_Col[12] = new ExcelImportStruct("Cont_40", "Cont 40", Excel_Col_Type.COL_INT, 10);
            INPUT_SXXK_XK_Col[13] = new ExcelImportStruct("Trong_luong", "Trọng lượng", Excel_Col_Type.COL_FLOAT, 20);
            INPUT_SXXK_XK_Col[14] = new ExcelImportStruct("Ma_HS", "Mã HS", Excel_Col_Type.COL_STRING, 20);
            INPUT_SXXK_XK_Col[15] = new ExcelImportStruct("Ma_hang", "Mã hàng", Excel_Col_Type.COL_STRING, 50);
            INPUT_SXXK_XK_Col[16] = new ExcelImportStruct("Ten_hang", "Tên hàng", Excel_Col_Type.COL_STRING, 100);
            INPUT_SXXK_XK_Col[17] = new ExcelImportStruct("Don_vi_tinh", "Đơn vị tính", Excel_Col_Type.COL_STRING, 30);
            INPUT_SXXK_XK_Col[18] = new ExcelImportStruct("So_luong", "Số lượng", Excel_Col_Type.COL_FLOAT, 20);
            INPUT_SXXK_XK_Col[19] = new ExcelImportStruct("Tri_gia_VND", "Trị giá VND", Excel_Col_Type.COL_FLOAT, 20);
            INPUT_SXXK_XK_Col[20] = new ExcelImportStruct("Don_gia", "Đơn giá", Excel_Col_Type.COL_FLOAT, 20);
            INPUT_SXXK_XK_Col[21] = new ExcelImportStruct("Tri_gia_NT", "Trị giá NT", Excel_Col_Type.COL_FLOAT, 20);
            INPUT_SXXK_XK_Col[22] = new ExcelImportStruct("Nuoc_xuat_xu", "Nước xuất xứ", Excel_Col_Type.COL_STRING, 50);
        }

        private bool INPUT_SXXK_XK_Get_Col_info(Excel.Workbook cur_wbook)
        {
            int i, col, row;
            string cell_val;
            string error_log = "";
            bool error = false;

            for (i = 0; i < INPUT_SXXK_XK_COL_NUM; i++)
            {
                INPUT_SXXK_XK_Col[i].Col = 0;
            }

            row = 1;
            for (col = 1; col < 100; col++)
            {
                cell_val = Get_Excel_Line((Excel.Worksheet)cur_wbook.Sheets[1], row, col, 1);
                for (i = 0; i < INPUT_SXXK_XK_COL_NUM; i++)
                {
                    if (cell_val == INPUT_SXXK_XK_Col[i].Col_str)
                    {
                        INPUT_SXXK_XK_Col[i].Col = col;
                        break;
                    }
                }
            }

            for (i = 0; i < INPUT_SXXK_XK_COL_NUM; i++)
            {
                if (INPUT_SXXK_XK_Col[i].Col == 0)
                {
                    error_log += "Can not find Column:" + INPUT_SXXK_XK_Col[i].Col_str + "\n";
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

        private bool Import_INPUT_SXXK_XK_Table_in_file(string file_name)
        {
            int row;
            string so_tk, ngay_dk, ma_loai_hinh;
            string cell_str;


            ProgressBar1.Visible = true;
            StatusLabel.Text = "Loading File";
            row = 1;
            OpenWB = Open_excel_file(file_name, "");
            Error_log = "";
            if (INPUT_SXXK_XK_Get_Col_info(OpenWB) == true)
            {
                Load_Form_XK_Line();
                row = 2;
                cell_str = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, 1, 20);
                while (cell_str != "")
                {
                    so_tk = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, INPUT_SXXK_XK_Col[INPUT_SXXK_XK_So_TK].Col, INPUT_SXXK_XK_Col[INPUT_SXXK_XK_So_TK].Data_Max_len);
                    ngay_dk = Get_date_str_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, INPUT_SXXK_XK_Col[INPUT_SXXK_XK_Ngay_DK].Col);
                    ma_loai_hinh = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, INPUT_SXXK_XK_Col[INPUT_SXXK_XK_Ma_loai_hinh].Col, INPUT_SXXK_XK_Col[INPUT_SXXK_XK_Ma_loai_hinh].Data_Max_len);
                    // Kiem tra Line da co trong database chua
                    if (Is_exist_SXXK_XK(so_tk, ngay_dk, ma_loai_hinh) == true)
                    {
                        // Update for this row
                        Update_SXXK_XK_Line(so_tk, ngay_dk, ma_loai_hinh, (Excel.Worksheet)OpenWB.Sheets[1], row);
                    }
                    else
                    {
                        // Insert new row
                        Add_New_SXXK_XK_Line(so_tk, ngay_dk, ma_loai_hinh, (Excel.Worksheet)OpenWB.Sheets[1], row);
                    }
                    row++;
                    cell_str = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, 1, 20);
                    ProgressBar1.Value = row % 100;
                    StatusLabel.Text = "Loading File, Line: " + row.ToString();
                }

                Close_WorkBook(OpenWB);
                // Store data
                if (Update_SQL_Data(INPUT_XK_Table_Form.Data_da, INPUT_XK_Table_Form.Data_dtb) == true)
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

        private void Load_So_TK_XK_Item(string so_tk, string ngay_dk, string ma_lh)
        {
            string sql_cmd = @"SELECT * FROM [WHM_INFOMATION_DB].[dbo].[INPUT_XK_tb]";
            sql_cmd += " WHERE So_TK = '" + so_tk + "'";
            sql_cmd += " and Ngay_DK = '" + Convert.ToDateTime(ngay_dk.Trim()).Date.ToShortDateString() + "'";
            sql_cmd += " and Ma_loai_hinh = '" + ma_lh + "'";
            INPUT_XK_Table_Form.Load_DataBase(Database_WHM_Info_Con_Str, sql_cmd);
        }

        private bool Import_INPUT_XK_Table_in_file(string file_name)
        {
            int row;
            //string so_tk, ngay_dk, ma_loai_hinh;
            string cell_str, cur_tk, last_tk = "", cur_ma_lh, last_ma_lh = "", cur_ngay_dk, last_ngay_dk = "";
            string message;
            bool tk_opened = false;
            bool start_file = true;


            ProgressBar1.Visible = true;
            StatusLabel.Text = "Loading File";
            row = 1;
            OpenWB = Open_excel_file(file_name, "");
            Error_log = "";
            if (INPUT_SXXK_XK_Get_Col_info(OpenWB) == true)
            {
                Load_Form_List_TK_XK_Line();
                row = 2;
                cell_str = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, 2, 20);
                while (cell_str != "")
                {
                    cur_tk = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, INPUT_SXXK_XK_Col[INPUT_SXXK_XK_So_TK].Col, INPUT_SXXK_XK_Col[INPUT_SXXK_XK_So_TK].Data_Max_len);
                    cur_ngay_dk = Get_date_str_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, INPUT_SXXK_XK_Col[INPUT_SXXK_XK_Ngay_DK].Col);
                    cur_ma_lh = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, INPUT_SXXK_XK_Col[INPUT_SXXK_XK_Ma_loai_hinh].Col, INPUT_SXXK_XK_Col[INPUT_SXXK_XK_Ma_loai_hinh].Data_Max_len);
                    if ((last_tk == "") && (cur_tk == ""))
                    {
                        MessageBox.Show("Can not Find Fisrt TK", "Error");
                        return false;
                    }

                    if (cur_tk != last_tk)
                    {
                        if (start_file == true)
                        {
                            start_file = false;
                        }
                        else
                        {
                            //Save Current TK
                            INPUT_XK_Table_Form.Save_Data();
                        }
                        Cur_Item_index = 1;

                        // New TK
                        if (Is_exist_So_TK_XK(cur_tk, cur_ngay_dk, cur_ma_lh, INPUT_XK_List_TK_Tbl) == true)
                        {
                            //NOTE : TK has already in database: Update new
                            message = "TK: " + cur_tk + " Ngày ĐK: " + cur_ngay_dk + " Mã loại hình:" + cur_ma_lh + " is has already.\nDo you want to update?";
                            if (MessageBox.Show(message, "Warning", MessageBoxButtons.YesNo) == DialogResult.Yes)
                            {
                                //NOTE : Delete All item in current TK
                                Load_So_TK_XK_Item(cur_tk, cur_ngay_dk, cur_ma_lh);
                                foreach (DataRow item_row in INPUT_XK_Table_Form.Data_dtb.Rows)
                                {
                                    item_row.Delete();
                                }
                                INPUT_XK_Table_Form.Save_Data();
                                //NOTE : add Current row into Items
                                tk_opened = true;
                                Add_New_SXXK_XK_Line(cur_tk, cur_ngay_dk, cur_ma_lh, (Excel.Worksheet)OpenWB.Sheets[1], row);
                            }
                            else
                            {
                                tk_opened = false;
                            }
                        }
                        else
                        {
                            // New So TK: Add New TK
                            Create_New_List_TK_XK_Line(cur_tk, cur_ngay_dk, cur_ma_lh, (Excel.Worksheet)OpenWB.Sheets[1], row);
                            Load_So_TK_XK_Item(cur_tk, cur_ngay_dk, cur_ma_lh);
                            Add_New_SXXK_XK_Line(cur_tk, cur_ngay_dk, cur_ma_lh, (Excel.Worksheet)OpenWB.Sheets[1], row);
                            tk_opened = true;
                        }

                        //End New TK
                    }
                    else
                    {
                        // Old TK
                        if ((cur_ngay_dk == last_ngay_dk) && (cur_ma_lh == last_ma_lh))
                        {
                            Cur_Item_index++;
                            if (tk_opened == true)
                            {
                                Add_New_SXXK_XK_Line(cur_tk, cur_ngay_dk, cur_ma_lh, (Excel.Worksheet)OpenWB.Sheets[1], row);
                            }
                        }
                        // end Old TK
                        else
                        {
                            Cur_Item_index = 1;
                            // New So TK: Add New TK
                            Create_New_List_TK_XK_Line(cur_tk, cur_ngay_dk, cur_ma_lh, (Excel.Worksheet)OpenWB.Sheets[1], row);
                            Load_So_TK_XK_Item(cur_tk, cur_ngay_dk, cur_ma_lh);
                            Add_New_SXXK_XK_Line(cur_tk, cur_ngay_dk, cur_ma_lh, (Excel.Worksheet)OpenWB.Sheets[1], row);
                            tk_opened = true;
                        }
                    }
                    last_tk = cur_tk;
                    last_ngay_dk = cur_ngay_dk;
                    last_ma_lh = cur_ma_lh;
                    row++;
                    cell_str = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, 2, 20);
                    ProgressBar1.Value = row % 100;
                    StatusLabel.Text = "Loading File, Line: " + row.ToString();
                }
                //INPUT_XK_Table_Form.Submit_BT_Click_event(null, null);
                Close_WorkBook(OpenWB);
                // Store data
                if (Update_SQL_Data(INPUT_XK_Table_Form.Data_da, INPUT_XK_Table_Form.Data_dtb) == true)
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

        private bool Load_List_TK_XK(string connection_str)
        {
            string sql_cmd = @"SELECT * FROM [WHM_INFOMATION_DB].[dbo].[List_TK_XK_tb]";
            if (XK_List_TK_TBL != null)
            {
                XK_List_TK_TBL.Clear();
            }
            XK_List_TK_TBL = Get_SQL_Data(connection_str, sql_cmd, ref XK_List_TK_da, ref XK_List_TK_ds);

            return true;
        }

        private bool Load_Ma_LH_XK(string connection_str)
        {
            string sql_cmd = @"SELECT [Ma_loai_hinh] FROM [WHM_INFOMATION_DB].[dbo].[List_TK_XK_tb]";
            string so_tk = INPUT_XK_So_TK_CbxL.My_Combo.Text.ToString().Trim();
            sql_cmd += " WHERE So_TK = '" + so_tk + "'";

            if (XK_Ma_LH_TBL != null)
            {
                XK_Ma_LH_TBL.Clear();
            }
            XK_Ma_LH_TBL = Get_SQL_Data(connection_str, sql_cmd, ref XK_Ma_LH_da, ref XK_Ma_LH_ds);

            return true;
        }

        private void Load_Form_List_TK_XK_Line()
        {
            string load_list_TK_str = "SELECT * FROM [WHM_INFOMATION_DB].[dbo].[List_TK_XK_tb]";

            // Clean old data
            if (INPUT_XK_List_TK_Tbl != null)
            {
                INPUT_XK_List_TK_Tbl.Clear();
            }

            //Load Data into Table
            INPUT_XK_List_TK_Tbl = Get_SQL_Data(Database_WHM_Info_Con_Str, load_list_TK_str, ref INPUT_XK_List_TK_da, ref INPUT_XK_List_TK_ds);
        }

        private void Load_Form_XK_Line()
        {
            if (INPUT_XK_Table_Form.Data_dtb != null)
            {
                INPUT_XK_Table_Form.Data_dtb.Clear();
            }

            //Load Data into Table and display in gridview
            INPUT_XK_Table_Form.Load_DataBase(Database_WHM_Info_Con_Str, @"SELECT * FROM [WHM_INFOMATION_DB].[dbo].[INPUT_XK_tb]");
        }

        private bool Is_exist_So_TK_XK(string sotk, string ngaydk, string ma_loaihinh, DataTable table)
        {
            string cur_so_tk, cur_ma_loaihinh, cur_ngay_dk; //cur_ngay_dk
            //DateTime cur_ngay_dk;
            foreach (DataRow row in table.Rows)
            {
                cur_so_tk = row[INPUT_SXXK_XK_Col[INPUT_SXXK_XK_So_TK].DB_str].ToString().Trim();
                cur_ngay_dk = Convert.ToDateTime(row[INPUT_SXXK_XK_Col[INPUT_SXXK_XK_Ngay_DK].DB_str].ToString().Trim()).Date.ToShortDateString();
                cur_ma_loaihinh = row[INPUT_SXXK_XK_Col[INPUT_SXXK_XK_Ma_loai_hinh].DB_str].ToString().Trim();
                if ((cur_so_tk == sotk) && (cur_ngay_dk == Convert.ToDateTime(ngaydk.Trim()).Date.ToShortDateString()) && (cur_ma_loaihinh == ma_loaihinh))
                {
                    return true;
                }
            }
            return false;
        }

        private bool Create_New_List_TK_XK_Line(string so_tk, string ngay_dk, string ma_loaihinh, Excel.Worksheet xsheet, int row)
        {
            DataRow new_row = INPUT_XK_List_TK_Tbl.NewRow();

            new_row["So_TK"] = so_tk;
            new_row["Ngay_DK"] = ngay_dk;
            new_row["Ma_loai_hinh"] = ma_loaihinh;
            //new_row["Ma_HS"] = Get_Text_Cell(xsheet, row, INPUT_SXXK_XK_Col[INPUT_SXXK_XK_Ma_HS].Col, INPUT_SXXK_XK_Col[INPUT_SXXK_XK_Ma_HS].Data_Max_len);
            //new_row["Ma_hang"] = Get_Text_Cell(xsheet, row, INPUT_SXXK_XK_Col[INPUT_SXXK_XK_Ma_hang].Col, INPUT_SXXK_XK_Col[INPUT_SXXK_XK_Ma_hang].Data_Max_len);
            //new_row["So_luong_TK"] = 1;
            new_row["KD_or_SX"] = Ma_KD_Or_SX;
            new_row["Ngay_import"] = DateTime.Now.ToShortDateString();

            INPUT_XK_List_TK_Tbl.Rows.Add(new_row);
            Update_SQL_Data(INPUT_XK_List_TK_da, INPUT_XK_List_TK_Tbl);
            return true;

        }

        private bool Is_exist_SXXK_XK(string sotk, string ngaydk, string ma_loaihinh)
        {
            string cur_so_tk, cur_ma_loaihinh, cur_ngay_dk; //cur_ngay_dk
            //DateTime cur_ngay_dk;
            foreach (DataRow row in INPUT_XK_Table_Form.Data_dtb.Rows)
            {
                cur_so_tk = row[INPUT_SXXK_XK_Col[INPUT_SXXK_XK_So_TK].DB_str].ToString().Trim();
                cur_ngay_dk = Convert.ToDateTime(row[INPUT_SXXK_XK_Col[INPUT_SXXK_XK_Ngay_DK].DB_str].ToString().Trim()).Date.ToShortDateString();
                cur_ma_loaihinh = row[INPUT_SXXK_XK_Col[INPUT_SXXK_XK_Ma_loai_hinh].DB_str].ToString().Trim();
                if ((cur_so_tk == sotk) && (cur_ngay_dk == Convert.ToDateTime(ngaydk.Trim()).Date.ToShortDateString()) && (cur_ma_loaihinh == ma_loaihinh))
                {
                    return true;
                }
            }
            return false;
        }

        private bool Update_SXXK_XK_Line(string so_tk, string ngay_dk, string ma_loaihinh, Excel.Worksheet xsheet, int row)
        {
            string cur_so_tk, cur_ngay_dk, cur_ma_loaihinh;
            Excel_Col_Type col_type;
            int i;

            foreach (DataRow dt_row in INPUT_XK_Table_Form.Data_dtb.Rows)
            {
                cur_so_tk = dt_row[INPUT_SXXK_XK_Col[INPUT_SXXK_XK_So_TK].DB_str].ToString().Trim();
                cur_ngay_dk = dt_row[INPUT_SXXK_XK_Col[INPUT_SXXK_XK_Ngay_DK].DB_str].ToString().Trim();
                cur_ma_loaihinh = dt_row[INPUT_SXXK_XK_Col[INPUT_SXXK_XK_Ma_loai_hinh].DB_str].ToString().Trim();
                if ((cur_so_tk == so_tk) && (cur_ngay_dk == ngay_dk) && (cur_ma_loaihinh == ma_loaihinh))
                {
                    for (i = 0; i < INPUT_SXXK_XK_COL_NUM; i++)
                    {
                        if ((i != INPUT_SXXK_XK_So_TK) && (i != INPUT_SXXK_XK_Ngay_DK) && (i != INPUT_SXXK_XK_Ma_loai_hinh))
                        {
                            col_type = INPUT_SXXK_XK_Col[i].Col_type;
                            switch (col_type)
                            {
                                case Excel_Col_Type.COL_STRING:
                                    dt_row[INPUT_SXXK_XK_Col[i].DB_str] = Get_Text_Cell(xsheet, row, INPUT_SXXK_XK_Col[i].Col, INPUT_SXXK_XK_Col[i].Data_Max_len);
                                    break;
                                case Excel_Col_Type.COL_INT:
                                    dt_row[INPUT_SXXK_XK_Col[i].DB_str] = Get_int_Cell(xsheet, row, INPUT_SXXK_XK_Col[i].Col);
                                    break;
                                case Excel_Col_Type.COL_FLOAT:
                                    dt_row[INPUT_SXXK_XK_Col[i].DB_str] = Get_float_Cell(xsheet, row, INPUT_SXXK_XK_Col[i].Col);
                                    break;
                                case Excel_Col_Type.COL_DATE:
                                    dt_row[INPUT_SXXK_XK_Col[i].DB_str] = Get_date_str_Cell(xsheet, row, INPUT_SXXK_XK_Col[i].Col);
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

        private bool Add_New_SXXK_XK_Line(string so_tk, string ngay_dk, string ma_loaihinh, Excel.Worksheet xsheet, int row)
        {
            DataRow new_row = INPUT_XK_Table_Form.Data_dtb.NewRow();
            Excel_Col_Type col_type;
            int i;

            new_row["STT"] = so_tk + "." + Cur_Item_index.ToString().PadLeft(2, '0');
            new_row["KD_or_SX"] = Ma_KD_Or_SX;
            new_row[INPUT_SXXK_XK_Col[INPUT_SXXK_XK_So_TK].DB_str] = so_tk;
            new_row[INPUT_SXXK_XK_Col[INPUT_SXXK_XK_Ngay_DK].DB_str] = ngay_dk;
            new_row[INPUT_SXXK_XK_Col[INPUT_SXXK_XK_Ma_loai_hinh].DB_str] = ma_loaihinh;
            for (i = 0; i < INPUT_SXXK_XK_COL_NUM; i++)
            {
                if ((i != INPUT_SXXK_XK_So_TK) && (i != INPUT_SXXK_XK_Ngay_DK) && (i != INPUT_SXXK_XK_Ma_loai_hinh))
                {
                    col_type = INPUT_SXXK_XK_Col[i].Col_type;
                    switch (col_type)
                    {
                        case Excel_Col_Type.COL_STRING:
                            new_row[INPUT_SXXK_XK_Col[i].DB_str] = Get_Text_Cell(xsheet, row, INPUT_SXXK_XK_Col[i].Col, INPUT_SXXK_XK_Col[i].Data_Max_len);
                            break;
                        case Excel_Col_Type.COL_INT:
                            new_row[INPUT_SXXK_XK_Col[i].DB_str] = Get_int_Cell(xsheet, row, INPUT_SXXK_XK_Col[i].Col);
                            break;
                        case Excel_Col_Type.COL_FLOAT:
                            new_row[INPUT_SXXK_XK_Col[i].DB_str] = Get_float_Cell(xsheet, row, INPUT_SXXK_XK_Col[i].Col);
                            break;
                        case Excel_Col_Type.COL_DATE:
                            new_row[INPUT_SXXK_XK_Col[i].DB_str] = Get_date_str_Cell(xsheet, row, INPUT_SXXK_XK_Col[i].Col);
                            break;
                        default:
                            break;
                    }
                }
            }
            INPUT_XK_Table_Form.Data_dtb.Rows.Add(new_row);
            return true;
        }

        private void XK_Find_So_TK()
        {
            string sql_cmd = @"SELECT * FROM [WHM_INFOMATION_DB].[dbo].[INPUT_XK_tb]";
            string so_tk = INPUT_XK_So_TK_CbxL.My_Combo.Text.ToString().Trim();
            string ma_lh = INPUT_XK_Loai_hinh_CbxL.My_Combo.Text.ToString().Trim();
            sql_cmd += " WHERE So_TK = '" + so_tk + "'";
            sql_cmd += " AND Ma_loai_hinh = '" + ma_lh + "'";
            if (INPUT_XK_Table_Form.Data_dtb != null)
            {
                INPUT_XK_Table_Form.Data_dtb.Clear();
                INPUT_XK_Table_Form.Refresh_Form();
            }
            INPUT_XK_Table_Form.Load_DataBase(Database_WHM_Info_Con_Str, sql_cmd);
        }
    }
}