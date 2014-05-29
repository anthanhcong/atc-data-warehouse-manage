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
            INPUT_SXXK_XK_Col[18] = new ExcelImportStruct("So_luong", "Số lượng", Excel_Col_Type.COL_INT, 20);
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
                        Create_New_SXXK_XK_Line(so_tk, ngay_dk, ma_loai_hinh, (Excel.Worksheet)OpenWB.Sheets[1], row);
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

        private void Load_Form_XK_Line()
        {
            if (INPUT_XK_Table_Form.Data_dtb != null)
            {
                INPUT_XK_Table_Form.Data_dtb.Clear();
            }

            //Load Data into Table and display in gridview
            INPUT_XK_Table_Form.Load_DataBase(Database_WHM_Info_Con_Str, @"SELECT * FROM [WHM_INFOMATION_DB].[dbo].[INPUT_XK_tb]");
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

        private bool Create_New_SXXK_XK_Line(string so_tk, string ngay_dk, string ma_loaihinh, Excel.Worksheet xsheet, int row)
        {
            DataRow new_row = INPUT_XK_Table_Form.Data_dtb.NewRow();
            Excel_Col_Type col_type;
            int i;

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
    }
}