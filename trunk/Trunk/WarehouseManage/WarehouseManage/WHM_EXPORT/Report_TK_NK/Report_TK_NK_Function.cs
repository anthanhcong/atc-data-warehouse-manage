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
        public DataTable Output_NK_tbl;
        public DataTable Temp_List_TK_tbl;
        public DataSet Temp_List_TK_tbl_ds = new DataSet();
        public SqlDataAdapter Temp_List_TK_tbl_da;

        public DataTable Temp_TK_tbl;
        public DataTable Temp_Data_tbl;
        decimal Sum_NT;

        private bool Report_Load_Ma_LH_NK(string connection_str)
        {
            string sql_cmd = @"SELECT distinct[Ma_loai_hinh] FROM [WHM_INFOMATION_DB].[dbo].[List_TK_NK_tb]";

            if (Report_NK_Ma_LH_TBL != null)
            {
                Report_NK_Ma_LH_TBL.Clear();
            }
            Report_NK_Ma_LH_TBL = Get_SQL_Data(connection_str, sql_cmd, ref Report_NK_Ma_LH_da, ref Report_NK_Ma_LH_ds);

            return true;
        }

        public void Process_Export_Form_NK()
        {
            string so_tk, ngay_dk, ma_loaihinh;

            Report_TK_NK_Table_Form.Load_DataBase(Database_WHM_Info_Con_Str, @"SELECT * FROM dbo.INPUT_NK_tb where So_TK = ''");
            Output_NK_tbl = Report_TK_NK_Table_Form.Data_dtb.Clone();
            Temp_Data_tbl = Report_TK_NK_Table_Form.Data_dtb.Clone();
            Get_data_dif_time();

            if (Temp_TK_tbl != null)
            {
                Temp_TK_tbl.Clear();
            }
            Temp_TK_tbl = Temp_List_TK_tbl.Copy();
            foreach (DataRow row in Temp_TK_tbl.Rows)
            {
                so_tk = row["So_TK"].ToString().Trim();
                ngay_dk = Convert.ToDateTime(row["Ngay_DK"].ToString().Trim()).Date.ToShortDateString();
                ma_loaihinh = row["Ma_loai_hinh"].ToString().Trim();
                Get_data_TK(so_tk, ngay_dk, ma_loaihinh);
                Sum_Qty_and_Re_calc();
                Final_Calc_Export();
                Paste_to_Ouput_table();
                Temp_Data_tbl.Clear();
            }
            Report_TK_NK_Table_Form.dataGridView_View.DataSource = Output_NK_tbl;
        }

        public void Get_data_dif_time()
        {
            string start_day, end_day, ma_lh;
            string sql_cmd = @"SELECT * FROM [WHM_INFOMATION_DB].[dbo].[List_TK_NK_tb]";
            if (Temp_List_TK_tbl != null)
            {
                Temp_List_TK_tbl.Clear();
            }
            start_day = Report_TK_NK_Start_Date.My_picker.Value.Date.ToString("yyyy-MM-dd");
            end_day = Report_TK_NK_End_Date.My_picker.Value.Date.ToString("yyyy-MM-dd");
            ma_lh = Report_TK_NK_Loai_hinh_CbxL.My_Combo.SelectedText.ToString().Trim();

            sql_cmd += " where Ngay_DK >= " + "'" + start_day + "'";
            sql_cmd += " and Ngay_DK <= " + "'" + end_day + "'";

            if (ma_lh != "")
            {
                sql_cmd += " and Ma_loai_hinh = " + "'" + ma_lh + "'";
            }

            //Report_TK_NK_Table_Form.Load_DataBase(Database_WHM_Info_Con_Str, sql_cmd);
            Temp_List_TK_tbl = Get_SQL_Data(Database_WHM_Info_Con_Str, sql_cmd, ref Temp_List_TK_tbl_da, ref Temp_List_TK_tbl_ds);
        }

        private bool Get_data_TK(string so_tk, string ngay_dk, string ma_loaihinh)
        {
            string sql_cmd = @"SELECT * FROM [WHM_INFOMATION_DB].[dbo].[INPUT_NK_tb]";
            if (Report_TK_NK_Table_Form.Data_dtb != null)
            {
                Report_TK_NK_Table_Form.Data_dtb.Clear();
            }

            sql_cmd += " where So_TK = " + "'" + so_tk + "'";
            sql_cmd += " and Ngay_DK = " + "'" + Convert.ToDateTime(ngay_dk).Date.ToString("yyyy-MM-dd") + "'";
            sql_cmd += " and Ma_loai_hinh = " + "'" + ma_loaihinh + "'";

            Report_TK_NK_Table_Form.Load_DataBase(Database_WHM_Info_Con_Str, sql_cmd);
            return true;
        }

        public bool Sum_Qty_and_Re_calc()
        {
            string ma_hang;
            float so_luong, don_gia;

            for (int i = 0; i < Report_TK_NK_Table_Form.Data_dtb.Rows.Count; i++)
            {

                ma_hang = Report_TK_NK_Table_Form.Data_dtb.Rows[i]["Ma_hang"].ToString().Trim();
                don_gia = float.Parse(Report_TK_NK_Table_Form.Data_dtb.Rows[i]["Don_gia"].ToString().Trim());
                so_luong = float.Parse(Report_TK_NK_Table_Form.Data_dtb.Rows[i]["So_luong"].ToString().Trim());
                Plus_or_Add_New_Ma_hang_Line(ma_hang, don_gia, so_luong, Report_TK_NK_Table_Form.Data_dtb.Rows[i]);
            }
            return true;
        }

        public bool Plus_or_Add_New_Ma_hang_Line(string ma_hang, float don_gia, float so_luong, DataRow row_add)
        {
            string filterExpression = "";
            float cur_qty;
            int so_dong;
            filterExpression = "Ma_hang =" + "'" + ma_hang + "'";
            filterExpression += " and Don_gia =" + "'" + don_gia + "'";

            DataRow[] rows = Temp_Data_tbl.Select(filterExpression);
            if (rows.Length != 0)
            {
                cur_qty = float.Parse(rows[0]["So_luong"].ToString().Trim());
                rows[0]["So_luong"] = cur_qty + so_luong;
            }
            else
            {
                //DataRow newrow = Temp_Data_tbl.NewRow();
                Temp_Data_tbl.ImportRow(row_add);

            }
            so_dong = Temp_Data_tbl.Rows.Count;
            return true;
        }

        public void Final_Calc_Export()
        {
            decimal so_luong, don_gia, phi_BH, phi_VC, ty_gia_VND, ts_XNK, ts_TTDB, ts_VAT, thu_khac;
            decimal tri_gia_NT, tri_gia_VND, tt_XNK, tt_TTDB, tt_VAT, t_thu_khac, tong_tien_thue;
            //decimal temp_tri_gia;
            Sum_NT = 0;
            foreach (DataRow row in Temp_Data_tbl.Rows)
            {
                so_luong = decimal.Parse(row["So_luong"].ToString().Trim());
                don_gia = decimal.Parse(row["Don_gia"].ToString().Trim());
                tri_gia_NT = Math.Round(so_luong * don_gia, 6);
                row["Tri_gia_NT"] = tri_gia_NT;
                Sum_NT = Sum_NT + tri_gia_NT;

            }
            foreach (DataRow row in Temp_Data_tbl.Rows)
            {
                phi_BH = decimal.Parse(row["Phi_BH"].ToString().Trim());
                phi_VC = decimal.Parse(row["Phi_VC"].ToString().Trim());
                ts_XNK = decimal.Parse(row["Thue_suat_XNK"].ToString().Trim());
                ts_TTDB = decimal.Parse(row["Thue_suat_TTDB"].ToString().Trim());
                ts_VAT = decimal.Parse(row["Thue_suat_VAT"].ToString().Trim());
                thu_khac = decimal.Parse(row["Thu_khac"].ToString().Trim());
                tri_gia_NT = decimal.Parse(row["Tri_gia_NT"].ToString().Trim());
                ty_gia_VND = decimal.Parse(row["Ty_gia_VND"].ToString().Trim());
                //temp_tri_gia = (((phi_BH + phi_VC) * tri_gia_NT / Sum_NT) + tri_gia_NT) ;
                //temp_tri_gia = temp_tri_gia * ty_gia_VND;
                //tri_gia_VND = Math.Round(temp_tri_gia* ty_gia_VND, 6);//Math.Round((((phi_BH + phi_VC) * tri_gia_NT / Sum_NT) + tri_gia_NT) * ty_gia_VND, 4);
                tri_gia_VND = Math.Round((((phi_BH + phi_VC) * tri_gia_NT / Sum_NT) + tri_gia_NT) * ty_gia_VND, 6);
                tt_XNK = Math.Round(tri_gia_VND * ts_XNK / 100, 6);
                tt_TTDB = Math.Round((tri_gia_VND + tt_XNK) * ts_TTDB / 100, 6);
                tt_VAT = Math.Round((tri_gia_VND + tt_XNK + tt_TTDB) * ts_VAT / 100, 6);
                t_thu_khac = Math.Round((tri_gia_VND + tt_XNK + tt_TTDB + tt_VAT) * thu_khac / 100, 6);
                tong_tien_thue = tt_XNK + tt_TTDB + tt_VAT + t_thu_khac;
                row["Tri_gia_VND"] = tri_gia_VND;
                row["Tien_thue_XNK"] = tt_XNK;
                row["Tien_thue_TTDB"] = tt_TTDB;
                row["Tien_thue_VAT"] = tt_VAT;
                row["Tien_thu_khac"] = t_thu_khac;
                row["Tong_tien_thue"] = tong_tien_thue;
            }
        }
        public void Paste_to_Ouput_table()
        {
            foreach (DataRow row in Temp_Data_tbl.Rows)
            {
                Output_NK_tbl.ImportRow(row);
            }
        }
    }
}