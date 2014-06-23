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
        /**********************************************/
        //              Report Nhập Khẩu              //
        /**********************************************/


        private System.Windows.Forms.TabPage Report_TK_NK_Manage_Tab;
        private Gridview_Grp Report_TK_NK_Table_Form;
        //public Button Report_TK_NK_Store_BT;
        //public Button Report_TK_NK_Test_Export_BT;

        private Button_Lbl Report_TK_NK_Find_BtL;
        private Button_Lbl Report_TK_NK_Report_BtL;
        private ComboBox_Lbl Report_TK_NK_So_TK_CbxL;
        private ComboBox_Lbl Report_TK_NK_Loai_hinh_CbxL;
        private TextBox_Lbl Report_TK_NK_So_TK_TxbL;
        private Checkbox_Lbl Report_TK_NK_Check_Select_All;
        private DatePick_LBL Report_TK_NK_Start_Date;
        private DatePick_LBL Report_TK_NK_End_Date;
        private GroupBox Report_TK_NK_Search_gbx;
        private GroupBox Report_TK_Report_NK_List_TK_NK_gbx;

        public DataTable Load_Report_TK_NK_TBL;
        //DataSet Load_Report_TK_NK_ds = new DataSet();
        //SqlDataAdapter Load_Report_TK_NK_da;

        public DataTable Report_NK_List_TK_TBL;
        DataSet Report_NK_List_TK_ds = new DataSet();
        SqlDataAdapter Report_NK_List_TK_da;

        public DataTable Report_NK_Ma_LH_TBL;
        DataSet Report_NK_Ma_LH_ds = new DataSet();
        SqlDataAdapter Report_NK_Ma_LH_da;

        private void Create_Report_TK_NK_Manage_Tab()
        {
            PosSize possize = new PosSize();
            string tab_name = "Report TK Nhập Khẩu";

            Report_TK_NK_Manage_Tab = new System.Windows.Forms.TabPage();
            Report_TK_NK_Manage_Tab.Text = tab_name;
            Report_TK_NK_Manage_Tab.SuspendLayout();
            Report_TK_NK_Manage_Tab.Location = new System.Drawing.Point(4, 22);
            Report_TK_NK_Manage_Tab.Size = new System.Drawing.Size(900, 390);
            Report_TK_NK_Manage_Tab.Padding = new System.Windows.Forms.Padding(3);
            //Report_TK_NK_Manage_Tab.TabIndex = 1;
            Report_TK_NK_Manage_Tab.UseVisualStyleBackColor = true;
            Report_TK_NK_Manage_Tab.ResumeLayout(true);
            Report_TK_NK_Manage_Tab.PerformLayout();
            this.MainTabControl.Controls.Add(this.Report_TK_NK_Manage_Tab);

            // Init Card Table
            possize.pos_x = 6;
            possize.pos_y = 140;
            possize.width = Report_TK_NK_Manage_Tab.Size.Width - 15;
            possize.height = Report_TK_NK_Manage_Tab.Size.Height - 150;
            Report_TK_NK_Table_Form = new Gridview_Grp(Report_TK_NK_Manage_Tab, "Report TK Nhap Khau Table", possize, AUTO_RESIZE,
                                                Database_WHM_Info_Con_Str, @"SELECT * FROM dbo.Report_TK_NK_tb", AnchorType.LEFT);
            //Report_TK_NK_Table_Form.Load_DataBase(Database_WHM_Info_Con_Str, @"SELECT * FROM dbo.Report_TK_NK_tb");

            Report_TK_NK_Form_Init();
            //Report_TK_NK_Init_BT();
        }

        public bool Report_TK_NK_Form_Init()
        {
            PosSize possize = new PosSize();
            Report_TK_NK_Search_gbx = new GroupBox();
            Report_TK_Report_NK_List_TK_NK_gbx = new GroupBox();

            Report_TK_NK_Manage_Tab.Controls.Add(Report_TK_NK_Search_gbx);
            Report_TK_NK_Search_gbx.Location = new System.Drawing.Point(170, 6);
            Report_TK_NK_Search_gbx.Name = "Report_TK_NK_Search_gbx";
            Report_TK_NK_Search_gbx.Size = new System.Drawing.Size(250, 106);
            Report_TK_NK_Search_gbx.TabIndex = 1;
            Report_TK_NK_Search_gbx.TabStop = false;
            Report_TK_NK_Search_gbx.Text = "Search";

            Report_TK_NK_Manage_Tab.Controls.Add(Report_TK_Report_NK_List_TK_NK_gbx);
            Report_TK_Report_NK_List_TK_NK_gbx.Location = new System.Drawing.Point(10, 6);
            Report_TK_Report_NK_List_TK_NK_gbx.Name = "Report_TK_Report_NK_List_TK_NK_gbx";
            Report_TK_Report_NK_List_TK_NK_gbx.Size = new System.Drawing.Size(150, 106);
            Report_TK_Report_NK_List_TK_NK_gbx.TabIndex = 1;
            Report_TK_Report_NK_List_TK_NK_gbx.TabStop = false;
            Report_TK_Report_NK_List_TK_NK_gbx.Text = "List TK Nhap Khau";

            possize.pos_x = 185;
            possize.pos_y = 42;
            Report_TK_NK_Find_BtL = new Button_Lbl(Report_TK_NK_Manage_Tab, "Search", possize, AnchorType.LEFT);
            Report_TK_NK_Find_BtL.My_Button.Size = new Size(55, 20);
            //Report_TK_NK_Find_BtL.My_Button.Click += new System.EventHandler(Report_TK_NK_Find_BtL_Click);
            Report_TK_NK_Search_gbx.Controls.Add(Report_TK_NK_Find_BtL.My_Button);

            possize.pos_x = 185;
            possize.pos_y = 68;
            Report_TK_NK_Report_BtL = new Button_Lbl(Report_TK_NK_Manage_Tab, "Report", possize, AnchorType.LEFT);
            Report_TK_NK_Report_BtL.My_Button.Size = new Size(55, 20);
            Report_TK_NK_Report_BtL.My_Button.Click += new System.EventHandler(Report_TK_NK_Report_BtL_Click);
            Report_TK_NK_Search_gbx.Controls.Add(Report_TK_NK_Report_BtL.My_Button);

            possize.pos_x = 185;
            possize.pos_y = 16;
            Report_TK_NK_Check_Select_All = new Checkbox_Lbl(Report_TK_NK_Manage_Tab, "All", possize, AnchorType.LEFT);
            Report_TK_NK_Check_Select_All.My_CheckBox.Checked = false;
            //Report_TK_NK_Check_Select_All.My_CheckBox.Click += new System.EventHandler(Report_TK_NK_Check_Select_All_CheckedChanged);

            //Load_List_TK_NK(Database_WHM_Info_Con_Str);
            possize.pos_x = 6;
            possize.pos_y = 16;
            Report_TK_NK_So_TK_CbxL = new ComboBox_Lbl(Report_TK_NK_Manage_Tab, "Số TK", possize, Report_NK_List_TK_TBL, "So_TK", "So_TK", AnchorType.LEFT);
            Report_TK_NK_So_TK_CbxL.My_Combo.Location = new Point(55, 16);
            Report_TK_NK_So_TK_CbxL.My_Combo.Size = new Size(77, 20);
            //Report_TK_NK_So_TK_CbxL.My_Combo.SelectedIndexChanged += new System.EventHandler(Report_TK_NK_So_TK_CbxL_Text_Change);
            //Report_TK_NK_So_TK_CbxL.My_Combo.Click += new EventHandler(Report_TK_NK_So_TK_CbxL_Click);

            Report_Load_Ma_LH_NK(Database_WHM_Info_Con_Str);
            possize.pos_x = 6;//6;
            possize.pos_y = 16;//42;
            Report_TK_NK_Loai_hinh_CbxL = new ComboBox_Lbl(Report_TK_NK_Manage_Tab, "Mã LH", possize, Report_NK_Ma_LH_TBL, "Ma_loai_hinh", "Ma_loai_hinh", AnchorType.LEFT);
            Report_TK_NK_Loai_hinh_CbxL.My_Combo.Location = new Point(55, 16);
            Report_TK_NK_Loai_hinh_CbxL.My_Combo.Size = new Size(100, 20);

            Report_TK_Report_NK_List_TK_NK_gbx.Controls.Add(Report_TK_NK_So_TK_CbxL.My_Label);
            Report_TK_Report_NK_List_TK_NK_gbx.Controls.Add(Report_TK_NK_So_TK_CbxL.My_Combo);
            Report_TK_NK_Search_gbx.Controls.Add(Report_TK_NK_Loai_hinh_CbxL.My_Label);
            Report_TK_NK_Search_gbx.Controls.Add(Report_TK_NK_Loai_hinh_CbxL.My_Combo);

            possize.pos_x = 6;
            possize.pos_y = 42;//74;
            Report_TK_NK_Start_Date = new DatePick_LBL(Report_TK_NK_Manage_Tab, "To", possize, AnchorType.LEFT);
            Report_TK_NK_Start_Date.My_picker.Location = new Point(55, 42);
            //Report_TK_NK_Start_Date.My_picker.ValueChanged += new EventHandler(Report_TK_NK_Start_Date_ValueChanged);

            possize.pos_x = 6;// 180;
            possize.pos_y = 68;//74;
            Report_TK_NK_End_Date = new DatePick_LBL(Report_TK_NK_Manage_Tab, "From", possize, AnchorType.LEFT);
            Report_TK_NK_End_Date.My_picker.Location = new Point(55, 68);
            //Report_TK_NK_End_Date.My_picker.ValueChanged += new EventHandler(Report_TK_NK_End_Date_ValueChanged);

            Report_TK_NK_Search_gbx.Controls.Add(Report_TK_NK_Check_Select_All.My_CheckBox);
            Report_TK_NK_Search_gbx.Controls.Add(Report_TK_NK_Start_Date.My_Label);
            Report_TK_NK_Search_gbx.Controls.Add(Report_TK_NK_Start_Date.My_picker);
            Report_TK_NK_Search_gbx.Controls.Add(Report_TK_NK_End_Date.My_Label);
            Report_TK_NK_Search_gbx.Controls.Add(Report_TK_NK_End_Date.My_picker);

            return true;
        }

    }
}