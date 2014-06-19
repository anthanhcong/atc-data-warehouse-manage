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
        const int Ma_List_Manage_COL_NUM = 5;
        ExcelImportStruct[] Ma_List_Manage_Col = new ExcelImportStruct[Ma_List_Manage_COL_NUM];
        const int Ma_List_Manage_Ma_Part_number = 0;

        public DataTable Ma_List_Manage_Tbl;
        public DataSet Ma_List_Manage_Tbl_ds = new DataSet();
        public SqlDataAdapter Ma_List_Manage_Tbl_da;

        private void Ma_List_Manage_InitExcelCol_Infor()
        {
            Ma_List_Manage_Col[0] = new ExcelImportStruct("Part_Number", "Part_Number", Excel_Col_Type.COL_STRING, 30);
            Ma_List_Manage_Col[1] = new ExcelImportStruct("Revision_Level", "Revision_Level", Excel_Col_Type.COL_STRING, 50);
            Ma_List_Manage_Col[2] = new ExcelImportStruct("Policy", "Policy", Excel_Col_Type.COL_STRING, 30);
            Ma_List_Manage_Col[3] = new ExcelImportStruct("Material_Description", "Material_Description", Excel_Col_Type.COL_STRING, 50);
            Ma_List_Manage_Col[4] = new ExcelImportStruct("Component_Unit", "Component_Unit", Excel_Col_Type.COL_STRING, 50);
        }

        private bool Ma_List_Manage_Get_Col_info(Excel.Workbook cur_wbook)
        {
            int i, col, row;
            string cell_val;
            string error_log = "";
            bool error = false;

            for (i = 0; i < Ma_List_Manage_COL_NUM; i++)
            {
                Ma_List_Manage_Col[i].Col = 0;
            }

            row = 2;
            for (col = 1; col < 20; col++)
            {
                cell_val = Get_Excel_Line((Excel.Worksheet)cur_wbook.Sheets[1], row, col, 1);
                for (i = 0; i < Ma_List_Manage_COL_NUM; i++)
                {
                    if (cell_val == Ma_List_Manage_Col[i].Col_str)
                    {
                        Ma_List_Manage_Col[i].Col = col;
                        break;
                    }
                }
            }

            for (i = 0; i < Ma_List_Manage_COL_NUM; i++)
            {
                if (Ma_List_Manage_Col[i].Col == 0)
                {
                    error_log += "Can not find Column:" + Ma_List_Manage_Col[i].Col_str + "\n";
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

        private bool Import_Ma_List_Manage_Table_in_file(string file_name)
        {
            int row;
            string part_number;
            string cell_str;


            ProgressBar1.Visible = true;
            StatusLabel.Text = "Loading File";
            row = 1;
            OpenWB = Open_excel_file(file_name, "");
            Error_log = "";
            if (Ma_List_Manage_Get_Col_info(OpenWB) == true)
            {
                Load_Form_Ma_List_Manage_Line();
                row = 3;
                cell_str = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, 2, 20);
                while (cell_str != "")
                {
                    part_number = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, Ma_List_Manage_Col[Ma_List_Manage_Ma_Part_number].Col, Ma_List_Manage_Col[Ma_List_Manage_Ma_Part_number].Data_Max_len);
                    // Kiem tra Line da co trong database chua
                    if (Is_exist_Ma_List_Manage(part_number) == true)
                    {
                        // Update for this row
                        if (Ma_List_Import_Auto_Update.My_CheckBox.Checked == true)
                        {
                            Update_Ma_List_Manage_Line(part_number, (Excel.Worksheet)OpenWB.Sheets[1], row);
                        }
                        else
                        {
                            DialogResult thongbao;

                                thongbao = (MessageBox.Show("Part number: " + part_number + " was created.\n"
                                                             + "Do you want to update ?", " Attention ",
                                                                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning));
                                if (thongbao == DialogResult.Yes)
                                {
                                    Update_Ma_List_Manage_Line(part_number, (Excel.Worksheet)OpenWB.Sheets[1], row);
                                }
                        }
                    }
                    else
                    {
                        // Insert new row
                        Create_New_Ma_List_Manage_Line(part_number, (Excel.Worksheet)OpenWB.Sheets[1], row);
                    }
                    row++;
                    cell_str = Get_Text_Cell((Excel.Worksheet)OpenWB.Sheets[1], row, 2, 20);
                    ProgressBar1.Value = row % 100;
                    StatusLabel.Text = "Loading File, Line: " + row.ToString();
                }

                Close_WorkBook(OpenWB);
                // Store data
                if (Update_SQL_Data(Material_List_Table_Form.Data_da, Material_List_Table_Form.Data_dtb) == true)
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

        private void Load_Form_Ma_List_Manage_Line()
        {
            if (Material_List_Table_Form.Data_dtb != null)
            {
                Material_List_Table_Form.Data_dtb.Clear();
            }

            //Load Data into Table and display in gridview
            Material_List_Table_Form.Load_DataBase(Database_WHM_Stock_Con_Str, @"SELECT * FROM dbo.Material_List_tb");
        }

        private bool Is_exist_Ma_List_Manage(string part_number)
        {
            string cur_part_number;
            foreach (DataRow row in Material_List_Table_Form.Data_dtb.Rows)
            {
                cur_part_number = row[Ma_List_Manage_Col[Ma_List_Manage_Ma_Part_number].DB_str].ToString().Trim();
                if (cur_part_number == part_number)
                {
                    return true;
                }
            }
            return false;
        }

        private bool Update_Ma_List_Manage_Line(string part_number, Excel.Worksheet xsheet, int row)
        {
            string cur_part_number;
            Excel_Col_Type col_type;
            int i;

            foreach (DataRow dt_row in Material_List_Table_Form.Data_dtb.Rows)
            {
                cur_part_number = dt_row[Ma_List_Manage_Col[Ma_List_Manage_Ma_Part_number].DB_str].ToString().Trim();
                if (cur_part_number == part_number)
                {
                    for (i = 0; i < Ma_List_Manage_COL_NUM; i++)
                    {
                        if (i != Ma_List_Manage_Ma_Part_number)
                        {
                            col_type = Ma_List_Manage_Col[i].Col_type;
                            switch (col_type)
                            {
                                case Excel_Col_Type.COL_STRING:
                                    dt_row[Ma_List_Manage_Col[i].DB_str] = Get_Text_Cell(xsheet, row, Ma_List_Manage_Col[i].Col, Ma_List_Manage_Col[i].Data_Max_len);
                                    break;
                                case Excel_Col_Type.COL_INT:
                                    dt_row[Ma_List_Manage_Col[i].DB_str] = Get_int_Cell(xsheet, row, Ma_List_Manage_Col[i].Col);
                                    break;
                                case Excel_Col_Type.COL_FLOAT:
                                    dt_row[Ma_List_Manage_Col[i].DB_str] = Get_float_Cell(xsheet, row, Ma_List_Manage_Col[i].Col);
                                    break;
                                case Excel_Col_Type.COL_DATE:
                                    dt_row[Ma_List_Manage_Col[i].DB_str] = Get_date_str_Cell(xsheet, row, Ma_List_Manage_Col[i].Col);
                                    break;
                                default:
                                    break;
                            }
                        }
                        dt_row["Create_by"] = User_Name;
                    }
                    return true;
                }
            }
            return false;
        }

        private bool Create_New_Ma_List_Manage_Line(string part_number, Excel.Worksheet xsheet, int row)
        {
            DataRow new_row = Material_List_Table_Form.Data_dtb.NewRow();
            Excel_Col_Type col_type;
            int i;

            new_row[Ma_List_Manage_Col[Ma_List_Manage_Ma_Part_number].DB_str] = part_number;
            new_row["Create_by"] = User_Name;
            for (i = 0; i < Ma_List_Manage_COL_NUM; i++)
            {
                if (i != Ma_List_Manage_Ma_Part_number)
                {
                    col_type = Ma_List_Manage_Col[i].Col_type;
                    switch (col_type)
                    {
                        case Excel_Col_Type.COL_STRING:
                            new_row[Ma_List_Manage_Col[i].DB_str] = Get_Text_Cell(xsheet, row, Ma_List_Manage_Col[i].Col, Ma_List_Manage_Col[i].Data_Max_len);
                            break;
                        case Excel_Col_Type.COL_INT:
                            new_row[Ma_List_Manage_Col[i].DB_str] = Get_int_Cell(xsheet, row, Ma_List_Manage_Col[i].Col);
                            break;
                        case Excel_Col_Type.COL_FLOAT:
                            new_row[Ma_List_Manage_Col[i].DB_str] = Get_float_Cell(xsheet, row, Ma_List_Manage_Col[i].Col);
                            break;
                        case Excel_Col_Type.COL_DATE:
                            new_row[Ma_List_Manage_Col[i].DB_str] = Get_date_str_Cell(xsheet, row, Ma_List_Manage_Col[i].Col);
                            break;
                        default:
                            break;
                    }
                }
            }
            Material_List_Table_Form.Data_dtb.Rows.Add(new_row);
            return true;
        }
    }
}