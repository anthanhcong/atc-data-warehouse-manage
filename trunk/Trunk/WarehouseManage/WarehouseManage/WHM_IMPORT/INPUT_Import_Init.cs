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
        //              Input Nhập Khẩu               //
        /**********************************************/


        private System.Windows.Forms.TabPage INPUT_NK_Manage_Tab;
        private Gridview_Grp INPUT_NK_Table_Form;
        private Button INPUT_NK_KD_Import_BT;
        private Button INPUT_NK_SX_Import_BT;
        public Button INPUT_NK_Store_BT;
        //public Button INPUT_NK_Search_BT;
        //private TextBox_Lbl INPUT_NK_Search_Txt_Lb;
        public Button INPUT_NK_Test_Export_BT;

        private Button_Lbl INPUT_NK_Find_BtL;
        private ComboBox_Lbl INPUT_NK_So_TK_CbxL;
        private ComboBox_Lbl INPUT_NK_Loai_hinh_CbxL;
        private TextBox_Lbl INPUT_NK_So_TK_TxbL;
        //private TextBox_Lbl INPUT_NK_Ngay_DK;
        private TextBox_Lbl INPUT_NK_Ma_loai_hinh_TxbL;
        private TextBox_Lbl INPUT_NK_Ma_hang_TxbL;
        private TextBox_Lbl INPUT_NK_Ma_HS_TxbL;
        //private TextBox_Lbl INPUT_NK_
        private Checkbox_Lbl INPUT_NK_Check_So_TK;
        private Checkbox_Lbl INPUT_NK_Check_Ngay_DK;
        private Checkbox_Lbl INPUT_NK_Check_Ma_loai_hinh;
        private Checkbox_Lbl INPUT_NK_Check_Ma_hang;
        private Checkbox_Lbl INPUT_NK_Check_Ma_HS;
        private Checkbox_Lbl INPUT_NK_Check_Recent_Day;
        private DatePick_LBL INPUT_NK_Start_Date;
        private DatePick_LBL INPUT_NK_End_Date;
        private GroupBox INPUT_NK_Search_gbx;
        private GroupBox INPUT_NK_List_TK_NK_gbx;

        public DataTable Load_INPUT_NK_TBL;
        //DataSet Load_INPUT_NK_ds = new DataSet();
        //SqlDataAdapter Load_INPUT_NK_da;

        public DataTable NK_List_TK_TBL;
        DataSet NK_List_TK_ds = new DataSet();
        SqlDataAdapter NK_List_TK_da;

        public DataTable NK_Ma_LH_TBL;
        DataSet NK_Ma_LH_ds = new DataSet();
        SqlDataAdapter NK_Ma_LH_da;

        private void Create_INPUT_NK_Manage_Tab()
        {
            PosSize possize = new PosSize();
            string tab_name = "Input Nhập Khẩu";

            INPUT_NK_Manage_Tab = new System.Windows.Forms.TabPage();
            INPUT_NK_Manage_Tab.Text = tab_name;
            INPUT_NK_Manage_Tab.SuspendLayout();
            INPUT_NK_Manage_Tab.Location = new System.Drawing.Point(4, 22);
            INPUT_NK_Manage_Tab.Size = new System.Drawing.Size(900, 390);
            INPUT_NK_Manage_Tab.Padding = new System.Windows.Forms.Padding(3);
            //INPUT_NK_Manage_Tab.TabIndex = 1;
            INPUT_NK_Manage_Tab.UseVisualStyleBackColor = true;
            INPUT_NK_Manage_Tab.ResumeLayout(true);
            INPUT_NK_Manage_Tab.PerformLayout();
            this.MainTabControl.Controls.Add(this.INPUT_NK_Manage_Tab);

            // Init Card Table
            possize.pos_x = 6;
            possize.pos_y = 150;
            possize.width = INPUT_NK_Manage_Tab.Size.Width - 15;
            possize.height = INPUT_NK_Manage_Tab.Size.Height - 160;
            INPUT_NK_Table_Form = new Gridview_Grp(INPUT_NK_Manage_Tab, "Input Nhap Khau Table", possize, AUTO_RESIZE,
                                                Database_WHM_Info_Con_Str, @"SELECT * FROM dbo.INPUT_NK_tb", AnchorType.LEFT);
            //INPUT_NK_Table_Form.Load_DataBase(Database_WHM_Info_Con_Str, @"SELECT * FROM dbo.INPUT_NK_tb");

            INPUT_NK_Form_Init();
            INPUT_NK_Init_BT();
        }

        public bool INPUT_NK_Form_Init()
        {
            PosSize possize = new PosSize();
            INPUT_NK_Search_gbx = new GroupBox();
            INPUT_NK_List_TK_NK_gbx = new GroupBox();

            INPUT_NK_Manage_Tab.Controls.Add(INPUT_NK_Search_gbx);
            INPUT_NK_Search_gbx.Location = new System.Drawing.Point(170, 6);
            INPUT_NK_Search_gbx.Name = "INPUT_NK_Search_gbx";
            INPUT_NK_Search_gbx.Size = new System.Drawing.Size(550, 106);
            INPUT_NK_Search_gbx.TabIndex = 1;
            INPUT_NK_Search_gbx.TabStop = false;
            INPUT_NK_Search_gbx.Text = "Search";

            INPUT_NK_Manage_Tab.Controls.Add(INPUT_NK_List_TK_NK_gbx);
            INPUT_NK_List_TK_NK_gbx.Location = new System.Drawing.Point(10, 6);
            INPUT_NK_List_TK_NK_gbx.Name = "INPUT_NK_List_TK_NK_gbx";
            INPUT_NK_List_TK_NK_gbx.Size = new System.Drawing.Size(150, 106);
            INPUT_NK_List_TK_NK_gbx.TabIndex = 1;
            INPUT_NK_List_TK_NK_gbx.TabStop = false;
            INPUT_NK_List_TK_NK_gbx.Text = "List TK Nhap Khau";

            possize.pos_x = 350;
            possize.pos_y = 74;
            INPUT_NK_Find_BtL = new Button_Lbl(INPUT_NK_Manage_Tab, "Search", possize, AnchorType.LEFT);
            INPUT_NK_Find_BtL.My_Button.Size = new Size(55, 23);
            INPUT_NK_Find_BtL.My_Button.Click += new System.EventHandler(INPUT_NK_Find_BtL_Click);
            INPUT_NK_Search_gbx.Controls.Add(INPUT_NK_Find_BtL.My_Button);

            possize.pos_x = 350;
            possize.pos_y = 16;
            INPUT_NK_Check_So_TK = new Checkbox_Lbl(INPUT_NK_Manage_Tab, "TK", possize, AnchorType.LEFT);
            //INPUT_NK_Check_So_TK.My_CheckBox.Checked = true;
            //INPUT_NK_Check_So_TK.My_CheckBox.Click += new System.EventHandler(INPUT_NK_Check_So_TK_CheckedChanged);

            possize.pos_x = 350;
            possize.pos_y = 36;
            INPUT_NK_Check_Ngay_DK = new Checkbox_Lbl(INPUT_NK_Manage_Tab, "Ngày ĐK", possize, AnchorType.LEFT);
            INPUT_NK_Check_Ngay_DK.My_CheckBox.Checked = false;
            //INPUT_NK_Check_Ngay_DK.My_CheckBox.Click += new System.EventHandler(INPUT_NK_Check_Ngay_DK_CheckedChanged);

            possize.pos_x = 350;
            possize.pos_y = 56;
            INPUT_NK_Check_Ma_loai_hinh = new Checkbox_Lbl(INPUT_NK_Manage_Tab, "Mã loại hình", possize, AnchorType.LEFT);
            //INPUT_NK_Check_Ma_loai_hinh.My_CheckBox.Checked = true;
            //INPUT_NK_Check_Ma_loai_hinh.My_CheckBox.Click += new System.EventHandler(INPUT_NK_Check_Ma_loai_hinh_CheckedChanged);

            possize.pos_x = 450;
            possize.pos_y = 16;
            INPUT_NK_Check_Ma_hang = new Checkbox_Lbl(INPUT_NK_Manage_Tab, "Mã hàng", possize, AnchorType.LEFT);
            INPUT_NK_Check_Ma_hang.My_CheckBox.Checked = false;
            //INPUT_NK_Check_Ma_hang.My_CheckBox.Click += new System.EventHandler(INPUT_NK_Check_Ma_hang_CheckedChanged);

            possize.pos_x = 450;
            possize.pos_y = 36;
            INPUT_NK_Check_Ma_HS = new Checkbox_Lbl(INPUT_NK_Manage_Tab, "Mã HS", possize, AnchorType.LEFT);
            INPUT_NK_Check_Ma_HS.My_CheckBox.Checked = false;
            //INPUT_NK_Check_Ma_HS.My_CheckBox.Click += new System.EventHandler(INPUT_NK_Check_Ma_HS_CheckedChanged);

            possize.pos_x = 450;
            possize.pos_y = 56;
            INPUT_NK_Check_Recent_Day = new Checkbox_Lbl(INPUT_NK_Manage_Tab, "Ngày gần đây", possize, AnchorType.LEFT);
            INPUT_NK_Check_Recent_Day.My_CheckBox.Checked = false;
            //INPUT_NK_Check_Recent_Day.My_CheckBox.Click += new System.EventHandler(INPUT_NK_Check_Recent_Day_CheckedChanged);

            Load_List_TK_NK(Database_WHM_Info_Con_Str);
            possize.pos_x = 6;
            possize.pos_y = 16;
            INPUT_NK_So_TK_CbxL = new ComboBox_Lbl(INPUT_NK_Manage_Tab, "Số TK", possize, NK_List_TK_TBL, "So_TK", "So_TK", AnchorType.LEFT);
            INPUT_NK_So_TK_CbxL.My_Combo.Location = new Point(55, 16);
            INPUT_NK_So_TK_CbxL.My_Combo.Size = new Size(77, 20);
            INPUT_NK_So_TK_CbxL.My_Combo.SelectedIndexChanged += new System.EventHandler(INPUT_NK_So_TK_CbxL_Text_Change);
            INPUT_NK_So_TK_CbxL.My_Combo.Click += new EventHandler(INPUT_NK_So_TK_CbxL_Click);
            
            Load_Ma_LH_NK(Database_WHM_Info_Con_Str);
            possize.pos_x = 6;
            possize.pos_y = 42;
            INPUT_NK_Loai_hinh_CbxL = new ComboBox_Lbl(INPUT_NK_Manage_Tab, "Mã LH", possize, NK_Ma_LH_TBL, "Ma_loai_hinh", "Ma_loai_hinh", AnchorType.LEFT);
            INPUT_NK_Loai_hinh_CbxL.My_Combo.Location = new Point(55, 42);
            INPUT_NK_Loai_hinh_CbxL.My_Combo.Size = new Size(77, 20);

            INPUT_NK_List_TK_NK_gbx.Controls.Add(INPUT_NK_So_TK_CbxL.My_Label);
            INPUT_NK_List_TK_NK_gbx.Controls.Add(INPUT_NK_So_TK_CbxL.My_Combo);
            INPUT_NK_List_TK_NK_gbx.Controls.Add(INPUT_NK_Loai_hinh_CbxL.My_Label);
            INPUT_NK_List_TK_NK_gbx.Controls.Add(INPUT_NK_Loai_hinh_CbxL.My_Combo);

            possize.pos_x = 6;
            possize.pos_y = 16;
            INPUT_NK_So_TK_TxbL = new TextBox_Lbl(INPUT_NK_Manage_Tab, "Số TK", TextBox_Type.TEXT, possize, AnchorType.LEFT);
            INPUT_NK_So_TK_TxbL.My_TextBox.Location = new Point(80, 16);
            INPUT_NK_So_TK_TxbL.My_TextBox.Size = new Size(77, 20);

            possize.pos_x = 6;
            possize.pos_y = 42;
            INPUT_NK_Ma_loai_hinh_TxbL = new TextBox_Lbl(INPUT_NK_Manage_Tab, "Mã loại hình", TextBox_Type.TEXT, possize, AnchorType.LEFT);
            INPUT_NK_Ma_loai_hinh_TxbL.My_TextBox.Location = new Point(80, 42);
            INPUT_NK_Ma_loai_hinh_TxbL.My_TextBox.Size = new Size(77, 20);

            possize.pos_x = 180;
            possize.pos_y = 10;
            INPUT_NK_Ma_hang_TxbL = new TextBox_Lbl(INPUT_NK_Manage_Tab, "Mã hàng", TextBox_Type.TEXT, possize, AnchorType.LEFT);
            INPUT_NK_Ma_hang_TxbL.My_TextBox.Location = new Point(245, 16);
            INPUT_NK_Ma_hang_TxbL.My_TextBox.Size = new Size(77, 20);

            possize.pos_x = 180;
            possize.pos_y = 36;
            INPUT_NK_Ma_HS_TxbL = new TextBox_Lbl(INPUT_NK_Manage_Tab, "Mã HS", TextBox_Type.TEXT, possize, AnchorType.LEFT);
            INPUT_NK_Ma_HS_TxbL.My_TextBox.Location = new Point(245, 42);
            INPUT_NK_Ma_HS_TxbL.My_TextBox.Size = new Size(77, 20);

            possize.pos_x = 6;
            possize.pos_y = 74;
            INPUT_NK_Start_Date = new DatePick_LBL(INPUT_NK_Manage_Tab, "To", possize, AnchorType.LEFT);
            INPUT_NK_Start_Date.My_picker.Location = new Point(57, 74);
            //INPUT_NK_Start_Date.My_picker.ValueChanged += new EventHandler(INPUT_NK_Start_Date_ValueChanged);

            possize.pos_x = 180;
            possize.pos_y = 74;
            INPUT_NK_End_Date = new DatePick_LBL(INPUT_NK_Manage_Tab, "From", possize, AnchorType.LEFT);
            INPUT_NK_End_Date.My_picker.Location = new Point(222, 74);
            //INPUT_NK_End_Date.My_picker.ValueChanged += new EventHandler(INPUT_NK_End_Date_ValueChanged);

            INPUT_NK_Search_gbx.Controls.Add(INPUT_NK_So_TK_TxbL.My_Label);
            INPUT_NK_Search_gbx.Controls.Add(INPUT_NK_So_TK_TxbL.My_TextBox);
            INPUT_NK_Search_gbx.Controls.Add(INPUT_NK_Ma_loai_hinh_TxbL.My_Label);
            INPUT_NK_Search_gbx.Controls.Add(INPUT_NK_Ma_loai_hinh_TxbL.My_TextBox);
            INPUT_NK_Search_gbx.Controls.Add(INPUT_NK_Ma_hang_TxbL.My_Label);
            INPUT_NK_Search_gbx.Controls.Add(INPUT_NK_Ma_hang_TxbL.My_TextBox);
            INPUT_NK_Search_gbx.Controls.Add(INPUT_NK_Ma_HS_TxbL.My_Label);
            INPUT_NK_Search_gbx.Controls.Add(INPUT_NK_Ma_HS_TxbL.My_TextBox);
            INPUT_NK_Search_gbx.Controls.Add(INPUT_NK_Check_So_TK.My_CheckBox);
            INPUT_NK_Search_gbx.Controls.Add(INPUT_NK_Check_Ngay_DK.My_CheckBox);
            INPUT_NK_Search_gbx.Controls.Add(INPUT_NK_Check_Ma_loai_hinh.My_CheckBox);
            INPUT_NK_Search_gbx.Controls.Add(INPUT_NK_Check_Ma_hang.My_CheckBox);
            INPUT_NK_Search_gbx.Controls.Add(INPUT_NK_Check_Ma_HS.My_CheckBox);
            INPUT_NK_Search_gbx.Controls.Add(INPUT_NK_Check_Recent_Day.My_CheckBox);
            INPUT_NK_Search_gbx.Controls.Add(INPUT_NK_Start_Date.My_Label);
            INPUT_NK_Search_gbx.Controls.Add(INPUT_NK_Start_Date.My_picker);
            INPUT_NK_Search_gbx.Controls.Add(INPUT_NK_End_Date.My_Label);
            INPUT_NK_Search_gbx.Controls.Add(INPUT_NK_End_Date.My_picker);

            return true;
        }

        public bool INPUT_NK_Init_BT()
        {
            PosSize possize = new PosSize();
            INPUT_NK_KD_Import_BT = new Button();
            INPUT_NK_SX_Import_BT = new Button();
            INPUT_NK_Store_BT = new Button();
            //INPUT_NK_Search_BT = new Button();
            INPUT_NK_Test_Export_BT = new Button();


            INPUT_NK_KD_Import_BT.Name = "INPUT_NK_KD_Import_BT";
            INPUT_NK_KD_Import_BT.Text = "Import Nhap-KD";
            INPUT_NK_KD_Import_BT.Location = new System.Drawing.Point(320, 120);
            INPUT_NK_KD_Import_BT.Size = new System.Drawing.Size(95, 23);
            INPUT_NK_KD_Import_BT.UseVisualStyleBackColor = true;
            INPUT_NK_KD_Import_BT.Click += new System.EventHandler(INPUT_NK_KD_Import_BT_Click);
            INPUT_NK_Manage_Tab.Controls.Add(INPUT_NK_KD_Import_BT);

            INPUT_NK_SX_Import_BT.Name = "INPUT_NK_SX_Import_BT";
            INPUT_NK_SX_Import_BT.Text = "Import Nhap-SX";
            INPUT_NK_SX_Import_BT.Location = new System.Drawing.Point(450, 120);
            INPUT_NK_SX_Import_BT.Size = new System.Drawing.Size(90, 23);
            INPUT_NK_SX_Import_BT.UseVisualStyleBackColor = true;
            INPUT_NK_SX_Import_BT.Click += new System.EventHandler(INPUT_NK_SX_Import_BT_Click);
            INPUT_NK_Manage_Tab.Controls.Add(INPUT_NK_SX_Import_BT);

            INPUT_NK_Store_BT.Name = "INPUT_NK_Store_BT";
            INPUT_NK_Store_BT.Text = "Save";
            INPUT_NK_Store_BT.Location = new System.Drawing.Point(570, 120);
            INPUT_NK_Store_BT.Size = new System.Drawing.Size(50, 23);
            INPUT_NK_Store_BT.UseVisualStyleBackColor = true;
            INPUT_NK_Store_BT.Click += new System.EventHandler(INPUT_NK_Store_BT_Click);
            INPUT_NK_Manage_Tab.Controls.Add(INPUT_NK_Store_BT);

            INPUT_NK_Test_Export_BT.Name = "INPUT_NK_Test_Export_BT";
            INPUT_NK_Test_Export_BT.Text = "Export";
            INPUT_NK_Test_Export_BT.Location = new System.Drawing.Point(650, 120);
            INPUT_NK_Test_Export_BT.Size = new System.Drawing.Size(50, 23);
            INPUT_NK_Test_Export_BT.UseVisualStyleBackColor = true;
            INPUT_NK_Test_Export_BT.Click += new System.EventHandler(INPUT_NK_Test_Export_BT_Click);
            INPUT_NK_Manage_Tab.Controls.Add(INPUT_NK_Test_Export_BT);
            //INPUT_NK_Search_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            //INPUT_NK_Search_BT.Name = "INPUT_NK_Search_BT";
            //INPUT_NK_Search_BT.Text = "Search";
            //INPUT_NK_Search_BT.Location = new System.Drawing.Point(600, 35);
            //INPUT_NK_Search_BT.Size = new System.Drawing.Size(60, 23);
            //INPUT_NK_Search_BT.UseVisualStyleBackColor = true;
            //INPUT_NK_Search_BT.Click += new System.EventHandler(INPUT_NK_Search_BT_Click);
            //INPUT_NK_Manage_Tab.Controls.Add(INPUT_NK_Search_BT);

            //possize.pos_x = 500;
            //possize.pos_y = 6;
            //INPUT_NK_Search_Txt_Lb = new TextBox_Lbl(INPUT_NK_Manage_Tab, "              Search", TextBox_Type.TEXT, possize, AnchorType.RIGHT);

            return true;
        }

                /**********************************************/
                //              Input Xuất Khẩu               //
                /**********************************************/

        private System.Windows.Forms.TabPage INPUT_XK_Manage_Tab;
        private Gridview_Grp INPUT_XK_Table_Form;
        private Button INPUT_XK_KD_Import_BT;
        private Button INPUT_XK_SX_Import_BT;
        public Button INPUT_XK_Store_BT;
        //public Button INPUT_XK_Search_BT;
        //private TextBox_Lbl INPUT_XK_Search_Txt_Lb;

        private Button_Lbl INPUT_XK_Find_BtL;
        private ComboBox_Lbl INPUT_XK_So_TK_CbxL;
        private ComboBox_Lbl INPUT_XK_Loai_hinh_CbxL;
        private TextBox_Lbl INPUT_XK_So_TK_TxbL;
        //private TextBox_Lbl INPUT_XK_Ngay_DK;
        private TextBox_Lbl INPUT_XK_Ma_loai_hinh_TxbL;
        private TextBox_Lbl INPUT_XK_Ma_hang_TxbL;
        private TextBox_Lbl INPUT_XK_Ma_HS_TxbL;
        //private TextBox_Lbl INPUT_XK_
        private Checkbox_Lbl INPUT_XK_Check_So_TK;
        private Checkbox_Lbl INPUT_XK_Check_Ngay_DK;
        private Checkbox_Lbl INPUT_XK_Check_Ma_loai_hinh;
        private Checkbox_Lbl INPUT_XK_Check_Ma_hang;
        private Checkbox_Lbl INPUT_XK_Check_Ma_HS;
        private Checkbox_Lbl INPUT_XK_Check_Recent_Day;
        private DatePick_LBL INPUT_XK_Start_Date;
        private DatePick_LBL INPUT_XK_End_Date;
        private GroupBox INPUT_XK_Search_gbx;
        private GroupBox INPUT_XK_List_TK_NK_gbx;

        public DataTable Load_INPUT_XK_TBL;
        //DataSet Load_INPUT_XK_ds = new DataSet();
        //SqlDataAdapter Load_INPUT_XK_da;

        public DataTable XK_List_TK_TBL;
        DataSet XK_List_TK_ds = new DataSet();
        SqlDataAdapter XK_List_TK_da;


        public DataTable XK_Ma_LH_TBL;
        DataSet XK_Ma_LH_ds = new DataSet();
        SqlDataAdapter XK_Ma_LH_da;

        private void Create_INPUT_XK_Manage_Tab()
        {
            PosSize possize = new PosSize();
            string tab_name = "Input Xuất Khẩu";

            INPUT_XK_Manage_Tab = new System.Windows.Forms.TabPage();
            INPUT_XK_Manage_Tab.Text = tab_name;
            INPUT_XK_Manage_Tab.SuspendLayout();
            INPUT_XK_Manage_Tab.Location = new System.Drawing.Point(4, 22);
            INPUT_XK_Manage_Tab.Size = new System.Drawing.Size(900, 390);
            INPUT_XK_Manage_Tab.Padding = new System.Windows.Forms.Padding(3);
            //INPUT_XK_Manage_Tab.TabIndex = 1;
            INPUT_XK_Manage_Tab.UseVisualStyleBackColor = true;
            INPUT_XK_Manage_Tab.ResumeLayout(true);
            INPUT_XK_Manage_Tab.PerformLayout();
            this.MainTabControl.Controls.Add(this.INPUT_XK_Manage_Tab);

            // Init Card Table
            possize.pos_x = 6;
            possize.pos_y = 150;
            possize.width = INPUT_XK_Manage_Tab.Size.Width - 15;
            possize.height = INPUT_XK_Manage_Tab.Size.Height - 160;
            INPUT_XK_Table_Form = new Gridview_Grp(INPUT_XK_Manage_Tab, "Input Xuat Khau Table", possize, AUTO_RESIZE,
                                                Database_WHM_Info_Con_Str, @"SELECT * FROM dbo.INPUT_XK_tb", AnchorType.LEFT);
            //INPUT_XK_Table_Form.Load_DataBase(Database_WHM_Info_Con_Str, @"SELECT * FROM dbo.INPUT_XK_tb");

            INPUT_XK_Form_Init();
            INPUT_XK_Init_BT();
        }

        public bool INPUT_XK_Form_Init()
        {
            PosSize possize = new PosSize();
            INPUT_XK_Search_gbx = new GroupBox();
            INPUT_XK_List_TK_NK_gbx = new GroupBox();

            INPUT_XK_Manage_Tab.Controls.Add(INPUT_XK_Search_gbx);
            INPUT_XK_Search_gbx.Location = new System.Drawing.Point(170, 6);
            INPUT_XK_Search_gbx.Name = "INPUT_XK_Search_gbx";
            INPUT_XK_Search_gbx.Size = new System.Drawing.Size(550, 106);
            INPUT_XK_Search_gbx.TabIndex = 1;
            INPUT_XK_Search_gbx.TabStop = false;
            INPUT_XK_Search_gbx.Text = "Search";

            INPUT_XK_Manage_Tab.Controls.Add(INPUT_XK_List_TK_NK_gbx);
            INPUT_XK_List_TK_NK_gbx.Location = new System.Drawing.Point(10, 6);
            INPUT_XK_List_TK_NK_gbx.Name = "INPUT_NK_List_TK_NK_gbx";
            INPUT_XK_List_TK_NK_gbx.Size = new System.Drawing.Size(150, 106);
            INPUT_XK_List_TK_NK_gbx.TabIndex = 1;
            INPUT_XK_List_TK_NK_gbx.TabStop = false;
            INPUT_XK_List_TK_NK_gbx.Text = "List TK Xuat Khau";

            possize.pos_x = 350;
            possize.pos_y = 74;
            INPUT_XK_Find_BtL = new Button_Lbl(INPUT_XK_Manage_Tab, "Search", possize, AnchorType.LEFT);
            INPUT_XK_Find_BtL.My_Button.Size = new Size(55, 23);
            INPUT_XK_Find_BtL.My_Button.Click += new System.EventHandler(INPUT_XK_Find_BtL_Click);
            INPUT_XK_Search_gbx.Controls.Add(INPUT_XK_Find_BtL.My_Button);

            possize.pos_x = 350;
            possize.pos_y = 16;
            INPUT_XK_Check_So_TK = new Checkbox_Lbl(INPUT_XK_Manage_Tab, "TK", possize, AnchorType.LEFT);
            //INPUT_XK_Check_So_TK.My_CheckBox.Checked = true;
            //INPUT_XK_Check_So_TK.My_CheckBox.Click += new System.EventHandler(INPUT_XK_Check_So_TK_CheckedChanged);

            possize.pos_x = 350;
            possize.pos_y = 36;
            INPUT_XK_Check_Ngay_DK = new Checkbox_Lbl(INPUT_XK_Manage_Tab, "Ngày ĐK", possize, AnchorType.LEFT);
            INPUT_XK_Check_Ngay_DK.My_CheckBox.Checked = false;
            //INPUT_XK_Check_Ngay_DK.My_CheckBox.Click += new System.EventHandler(INPUT_XK_Check_Ngay_DK_CheckedChanged);

            possize.pos_x = 350;
            possize.pos_y = 56;
            INPUT_XK_Check_Ma_loai_hinh = new Checkbox_Lbl(INPUT_XK_Manage_Tab, "Mã loại hình", possize, AnchorType.LEFT);
            //INPUT_XK_Check_Ma_loai_hinh.My_CheckBox.Checked = true;
            //INPUT_XK_Check_Ma_loai_hinh.My_CheckBox.Click += new System.EventHandler(INPUT_XK_Check_Ma_loai_hinh_CheckedChanged);

            possize.pos_x = 450;
            possize.pos_y = 16;
            INPUT_XK_Check_Ma_hang = new Checkbox_Lbl(INPUT_XK_Manage_Tab, "Mã hàng", possize, AnchorType.LEFT);
            INPUT_XK_Check_Ma_hang.My_CheckBox.Checked = false;
            //INPUT_XK_Check_Ma_hang.My_CheckBox.Click += new System.EventHandler(INPUT_XK_Check_Ma_hang_CheckedChanged);

            possize.pos_x = 450;
            possize.pos_y = 36;
            INPUT_XK_Check_Ma_HS = new Checkbox_Lbl(INPUT_XK_Manage_Tab, "Mã HS", possize, AnchorType.LEFT);
            INPUT_XK_Check_Ma_HS.My_CheckBox.Checked = false;
            //INPUT_XK_Check_Ma_HS.My_CheckBox.Click += new System.EventHandler(INPUT_XK_Check_Ma_HS_CheckedChanged);

            possize.pos_x = 450;
            possize.pos_y = 56;
            INPUT_XK_Check_Recent_Day = new Checkbox_Lbl(INPUT_XK_Manage_Tab, "Ngày gần đây", possize, AnchorType.LEFT);
            INPUT_XK_Check_Recent_Day.My_CheckBox.Checked = false;
            //INPUT_XK_Check_Recent_Day.My_CheckBox.Click += new System.EventHandler(INPUT_XK_Check_Recent_Day_CheckedChanged);

            Load_List_TK_XK(Database_WHM_Info_Con_Str);
            possize.pos_x = 6;
            possize.pos_y = 16;
            INPUT_XK_So_TK_CbxL = new ComboBox_Lbl(INPUT_XK_Manage_Tab, "Số TK", possize, XK_List_TK_TBL, "So_TK", "So_TK", AnchorType.LEFT);
            INPUT_XK_So_TK_CbxL.My_Combo.Location = new Point(55, 16);
            INPUT_XK_So_TK_CbxL.My_Combo.Size = new Size(77, 20);
            INPUT_XK_So_TK_CbxL.My_Combo.SelectedIndexChanged += new System.EventHandler(INPUT_XK_So_TK_CbxL_Text_Change);
            INPUT_XK_So_TK_CbxL.My_Combo.Click += new EventHandler(INPUT_XK_So_TK_CbxL_Click);

            Load_Ma_LH_XK(Database_WHM_Info_Con_Str);
            possize.pos_x = 6;
            possize.pos_y = 42;
            INPUT_XK_Loai_hinh_CbxL = new ComboBox_Lbl(INPUT_XK_Manage_Tab, "Mã LH", possize, XK_Ma_LH_TBL, "Ma_loai_hinh", "Ma_loai_hinh", AnchorType.LEFT);
            INPUT_XK_Loai_hinh_CbxL.My_Combo.Location = new Point(55, 42);
            INPUT_XK_Loai_hinh_CbxL.My_Combo.Size = new Size(77, 20);

            INPUT_XK_List_TK_NK_gbx.Controls.Add(INPUT_XK_So_TK_CbxL.My_Label);
            INPUT_XK_List_TK_NK_gbx.Controls.Add(INPUT_XK_So_TK_CbxL.My_Combo);
            INPUT_XK_List_TK_NK_gbx.Controls.Add(INPUT_XK_Loai_hinh_CbxL.My_Label);
            INPUT_XK_List_TK_NK_gbx.Controls.Add(INPUT_XK_Loai_hinh_CbxL.My_Combo);

            possize.pos_x = 6;
            possize.pos_y = 16;
            INPUT_XK_So_TK_TxbL = new TextBox_Lbl(INPUT_XK_Manage_Tab, "Số TK", TextBox_Type.TEXT, possize, AnchorType.LEFT);
            INPUT_XK_So_TK_TxbL.My_TextBox.Location = new Point(80, 16);
            INPUT_XK_So_TK_TxbL.My_TextBox.Size = new Size(77, 20);

            possize.pos_x = 6;
            possize.pos_y = 42;
            INPUT_XK_Ma_loai_hinh_TxbL = new TextBox_Lbl(INPUT_XK_Manage_Tab, "Mã loại hình", TextBox_Type.TEXT, possize, AnchorType.LEFT);
            INPUT_XK_Ma_loai_hinh_TxbL.My_TextBox.Location = new Point(80, 42);
            INPUT_XK_Ma_loai_hinh_TxbL.My_TextBox.Size = new Size(77, 20);

            possize.pos_x = 180;
            possize.pos_y = 10;
            INPUT_XK_Ma_hang_TxbL = new TextBox_Lbl(INPUT_XK_Manage_Tab, "Mã hàng", TextBox_Type.TEXT, possize, AnchorType.LEFT);
            INPUT_XK_Ma_hang_TxbL.My_TextBox.Location = new Point(245, 16);
            INPUT_XK_Ma_hang_TxbL.My_TextBox.Size = new Size(77, 20);

            possize.pos_x = 180;
            possize.pos_y = 36;
            INPUT_XK_Ma_HS_TxbL = new TextBox_Lbl(INPUT_XK_Manage_Tab, "Mã HS", TextBox_Type.TEXT, possize, AnchorType.LEFT);
            INPUT_XK_Ma_HS_TxbL.My_TextBox.Location = new Point(245, 42);
            INPUT_XK_Ma_HS_TxbL.My_TextBox.Size = new Size(77, 20);

            possize.pos_x = 6;
            possize.pos_y = 74;
            INPUT_XK_Start_Date = new DatePick_LBL(INPUT_XK_Manage_Tab, "To", possize, AnchorType.LEFT);
            INPUT_XK_Start_Date.My_picker.Location = new Point(57, 74);
            //INPUT_XK_Start_Date.My_picker.ValueChanged += new EventHandler(INPUT_XK_Start_Date_ValueChanged);

            possize.pos_x = 180;
            possize.pos_y = 74;
            INPUT_XK_End_Date = new DatePick_LBL(INPUT_XK_Manage_Tab, "From", possize, AnchorType.LEFT);
            INPUT_XK_End_Date.My_picker.Location = new Point(222, 74);
            //INPUT_XK_End_Date.My_picker.ValueChanged += new EventHandler(INPUT_XK_End_Date_ValueChanged);

            INPUT_XK_Search_gbx.Controls.Add(INPUT_XK_So_TK_TxbL.My_Label);
            INPUT_XK_Search_gbx.Controls.Add(INPUT_XK_So_TK_TxbL.My_TextBox);
            INPUT_XK_Search_gbx.Controls.Add(INPUT_XK_Ma_loai_hinh_TxbL.My_Label);
            INPUT_XK_Search_gbx.Controls.Add(INPUT_XK_Ma_loai_hinh_TxbL.My_TextBox);
            INPUT_XK_Search_gbx.Controls.Add(INPUT_XK_Ma_hang_TxbL.My_Label);
            INPUT_XK_Search_gbx.Controls.Add(INPUT_XK_Ma_hang_TxbL.My_TextBox);
            INPUT_XK_Search_gbx.Controls.Add(INPUT_XK_Ma_HS_TxbL.My_Label);
            INPUT_XK_Search_gbx.Controls.Add(INPUT_XK_Ma_HS_TxbL.My_TextBox);
            INPUT_XK_Search_gbx.Controls.Add(INPUT_XK_Check_So_TK.My_CheckBox);
            INPUT_XK_Search_gbx.Controls.Add(INPUT_XK_Check_Ngay_DK.My_CheckBox);
            INPUT_XK_Search_gbx.Controls.Add(INPUT_XK_Check_Ma_loai_hinh.My_CheckBox);
            INPUT_XK_Search_gbx.Controls.Add(INPUT_XK_Check_Ma_hang.My_CheckBox);
            INPUT_XK_Search_gbx.Controls.Add(INPUT_XK_Check_Ma_HS.My_CheckBox);
            INPUT_XK_Search_gbx.Controls.Add(INPUT_XK_Check_Recent_Day.My_CheckBox);
            INPUT_XK_Search_gbx.Controls.Add(INPUT_XK_Start_Date.My_Label);
            INPUT_XK_Search_gbx.Controls.Add(INPUT_XK_Start_Date.My_picker);
            INPUT_XK_Search_gbx.Controls.Add(INPUT_XK_End_Date.My_Label);
            INPUT_XK_Search_gbx.Controls.Add(INPUT_XK_End_Date.My_picker);

            return true;
        }

        public bool INPUT_XK_Init_BT()
        {
            PosSize possize = new PosSize();
            INPUT_XK_KD_Import_BT = new Button();
            INPUT_XK_SX_Import_BT = new Button();
            INPUT_XK_Store_BT = new Button();
            //INPUT_XK_Search_BT = new Button();


            INPUT_XK_KD_Import_BT.Name = "INPUT_XK_KD_Import_BT";
            INPUT_XK_KD_Import_BT.Text = "Import Xuat-KD";
            INPUT_XK_KD_Import_BT.Location = new System.Drawing.Point(320, 120);
            INPUT_XK_KD_Import_BT.Size = new System.Drawing.Size(95, 23);
            INPUT_XK_KD_Import_BT.UseVisualStyleBackColor = true;
            INPUT_XK_KD_Import_BT.Click += new System.EventHandler(INPUT_XK_KD_Import_BT_Click);
            INPUT_XK_Manage_Tab.Controls.Add(INPUT_XK_KD_Import_BT);

            INPUT_XK_SX_Import_BT.Name = "INPUT_XK_SX_Import_BT";
            INPUT_XK_SX_Import_BT.Text = "Import Xuat-SX";
            INPUT_XK_SX_Import_BT.Location = new System.Drawing.Point(450, 120);
            INPUT_XK_SX_Import_BT.Size = new System.Drawing.Size(90, 23);
            INPUT_XK_SX_Import_BT.UseVisualStyleBackColor = true;
            INPUT_XK_SX_Import_BT.Click += new System.EventHandler(INPUT_XK_SX_Import_BT_Click);
            INPUT_XK_Manage_Tab.Controls.Add(INPUT_XK_SX_Import_BT);

            INPUT_XK_Store_BT.Name = "INPUT_XK_Store_BT";
            INPUT_XK_Store_BT.Text = "Save";
            INPUT_XK_Store_BT.Location = new System.Drawing.Point(570, 120);
            INPUT_XK_Store_BT.Size = new System.Drawing.Size(50, 23);
            INPUT_XK_Store_BT.UseVisualStyleBackColor = true;
            INPUT_XK_Store_BT.Click += new System.EventHandler(INPUT_XK_Store_BT_Click);
            INPUT_XK_Manage_Tab.Controls.Add(INPUT_XK_Store_BT);

            //INPUT_XK_Search_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            //INPUT_XK_Search_BT.Name = "INPUT_XK_Search_BT";
            //INPUT_XK_Search_BT.Text = "Search";
            //INPUT_XK_Search_BT.Location = new System.Drawing.Point(600, 35);
            //INPUT_XK_Search_BT.Size = new System.Drawing.Size(60, 23);
            //INPUT_XK_Search_BT.UseVisualStyleBackColor = true;
            //INPUT_XK_Search_BT.Click += new System.EventHandler(INPUT_XK_Search_BT_Click);
            //INPUT_XK_Manage_Tab.Controls.Add(INPUT_XK_Search_BT);

            //possize.pos_x = 500;
            //possize.pos_y = 6;
            //INPUT_XK_Search_Txt_Lb = new TextBox_Lbl(INPUT_XK_Manage_Tab, "              Search", TextBox_Type.TEXT, possize, AnchorType.RIGHT);

            return true;
        }
    }
}