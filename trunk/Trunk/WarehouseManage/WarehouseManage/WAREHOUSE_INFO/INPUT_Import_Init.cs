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
        public Button INPUT_NK_Search_BT;
        private TextBox_Lbl INPUT_NK_Search_Txt_Lb;

        public DataTable Load_INPUT_NK_TBL;
        //DataSet Load_INPUT_NK_ds = new DataSet();
        //SqlDataAdapter Load_INPUT_NK_da;

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
            possize.pos_y = 100;
            possize.width = INPUT_NK_Manage_Tab.Size.Width - 15;
            possize.height = INPUT_NK_Manage_Tab.Size.Height - 100;
            INPUT_NK_Table_Form = new Gridview_Grp(INPUT_NK_Manage_Tab, "Input KD-NK Table", possize, AUTO_RESIZE,
                                                Database_WHM_Info_Con_Str, @"SELECT * FROM dbo.INPUT_NK_tb", AnchorType.LEFT);
            INPUT_NK_Table_Form.Load_DataBase(Database_WHM_Info_Con_Str, @"SELECT * FROM dbo.INPUT_NK_tb");
            INPUT_NK_Init_BT();
        }

        public bool INPUT_NK_Init_BT()
        {
            PosSize possize = new PosSize();
            INPUT_NK_KD_Import_BT = new Button();
            INPUT_NK_SX_Import_BT = new Button();
            INPUT_NK_Store_BT = new Button();
            INPUT_NK_Search_BT = new Button();


            INPUT_NK_KD_Import_BT.Name = "INPUT_NK_KD_Import_BT";
            INPUT_NK_KD_Import_BT.Text = "Import Nhap-KD";
            INPUT_NK_KD_Import_BT.Location = new System.Drawing.Point(50, 69);
            INPUT_NK_KD_Import_BT.Size = new System.Drawing.Size(95, 23);
            INPUT_NK_KD_Import_BT.UseVisualStyleBackColor = true;
            INPUT_NK_KD_Import_BT.Click += new System.EventHandler(INPUT_NK_KD_Import_BT_Click);
            INPUT_NK_Manage_Tab.Controls.Add(INPUT_NK_KD_Import_BT);

            INPUT_NK_SX_Import_BT.Name = "INPUT_NK_SX_Import_BT";
            INPUT_NK_SX_Import_BT.Text = "Import Nhap-SX";
            INPUT_NK_SX_Import_BT.Location = new System.Drawing.Point(200, 69);
            INPUT_NK_SX_Import_BT.Size = new System.Drawing.Size(90, 23);
            INPUT_NK_SX_Import_BT.UseVisualStyleBackColor = true;
            INPUT_NK_SX_Import_BT.Click += new System.EventHandler(INPUT_NK_SX_Import_BT_Click);
            INPUT_NK_Manage_Tab.Controls.Add(INPUT_NK_SX_Import_BT);

            INPUT_NK_Store_BT.Name = "INPUT_NK_Store_BT";
            INPUT_NK_Store_BT.Text = "Save";
            INPUT_NK_Store_BT.Location = new System.Drawing.Point(350, 69);
            INPUT_NK_Store_BT.Size = new System.Drawing.Size(50, 23);
            INPUT_NK_Store_BT.UseVisualStyleBackColor = true;
            INPUT_NK_Store_BT.Click += new System.EventHandler(INPUT_NK_Store_BT_Click);
            INPUT_NK_Manage_Tab.Controls.Add(INPUT_NK_Store_BT);

            INPUT_NK_Search_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            INPUT_NK_Search_BT.Name = "INPUT_NK_Search_BT";
            INPUT_NK_Search_BT.Text = "Search";
            INPUT_NK_Search_BT.Location = new System.Drawing.Point(600, 35);
            INPUT_NK_Search_BT.Size = new System.Drawing.Size(60, 23);
            INPUT_NK_Search_BT.UseVisualStyleBackColor = true;
            INPUT_NK_Search_BT.Click += new System.EventHandler(INPUT_NK_Search_BT_Click);
            INPUT_NK_Manage_Tab.Controls.Add(INPUT_NK_Search_BT);

            possize.pos_x = 500;
            possize.pos_y = 6;
            INPUT_NK_Search_Txt_Lb = new TextBox_Lbl(INPUT_NK_Manage_Tab, "              Search", TextBox_Type.TEXT, possize, AnchorType.RIGHT);

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
        public Button INPUT_XK_Search_BT;
        private TextBox_Lbl INPUT_XK_Search_Txt_Lb;

        public DataTable Load_INPUT_XK_TBL;
        //DataSet Load_INPUT_XK_ds = new DataSet();
        //SqlDataAdapter Load_INPUT_XK_da;

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
            possize.pos_y = 100;
            possize.width = INPUT_XK_Manage_Tab.Size.Width - 15;
            possize.height = INPUT_XK_Manage_Tab.Size.Height - 100;
            INPUT_XK_Table_Form = new Gridview_Grp(INPUT_XK_Manage_Tab, "Input KD-NK Table", possize, AUTO_RESIZE,
                                                Database_WHM_Info_Con_Str, @"SELECT * FROM dbo.INPUT_XK_tb", AnchorType.LEFT);
            INPUT_XK_Table_Form.Load_DataBase(Database_WHM_Info_Con_Str, @"SELECT * FROM dbo.INPUT_XK_tb");
            INPUT_XK_Init_BT();
        }

        public bool INPUT_XK_Init_BT()
        {
            PosSize possize = new PosSize();
            INPUT_XK_KD_Import_BT = new Button();
            INPUT_XK_SX_Import_BT = new Button();
            INPUT_XK_Store_BT = new Button();
            INPUT_XK_Search_BT = new Button();


            INPUT_XK_KD_Import_BT.Name = "INPUT_XK_KD_Import_BT";
            INPUT_XK_KD_Import_BT.Text = "Import Xuat-KD";
            INPUT_XK_KD_Import_BT.Location = new System.Drawing.Point(50, 69);
            INPUT_XK_KD_Import_BT.Size = new System.Drawing.Size(90, 23);
            INPUT_XK_KD_Import_BT.UseVisualStyleBackColor = true;
            INPUT_XK_KD_Import_BT.Click += new System.EventHandler(INPUT_XK_KD_Import_BT_Click);
            INPUT_XK_Manage_Tab.Controls.Add(INPUT_XK_KD_Import_BT);

            INPUT_XK_SX_Import_BT.Name = "INPUT_XK_SX_Import_BT";
            INPUT_XK_SX_Import_BT.Text = "Import Xuat-SX";
            INPUT_XK_SX_Import_BT.Location = new System.Drawing.Point(200, 69);
            INPUT_XK_SX_Import_BT.Size = new System.Drawing.Size(90, 23);
            INPUT_XK_SX_Import_BT.UseVisualStyleBackColor = true;
            INPUT_XK_SX_Import_BT.Click += new System.EventHandler(INPUT_XK_SX_Import_BT_Click);
            INPUT_XK_Manage_Tab.Controls.Add(INPUT_XK_SX_Import_BT);

            INPUT_XK_Store_BT.Name = "INPUT_XK_Store_BT";
            INPUT_XK_Store_BT.Text = "Save";
            INPUT_XK_Store_BT.Location = new System.Drawing.Point(350, 69);
            INPUT_XK_Store_BT.Size = new System.Drawing.Size(50, 23);
            INPUT_XK_Store_BT.UseVisualStyleBackColor = true;
            INPUT_XK_Store_BT.Click += new System.EventHandler(INPUT_XK_Store_BT_Click);
            INPUT_XK_Manage_Tab.Controls.Add(INPUT_XK_Store_BT);

            INPUT_XK_Search_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            INPUT_XK_Search_BT.Name = "INPUT_XK_Search_BT";
            INPUT_XK_Search_BT.Text = "Search";
            INPUT_XK_Search_BT.Location = new System.Drawing.Point(600, 35);
            INPUT_XK_Search_BT.Size = new System.Drawing.Size(60, 23);
            INPUT_XK_Search_BT.UseVisualStyleBackColor = true;
            INPUT_XK_Search_BT.Click += new System.EventHandler(INPUT_XK_Search_BT_Click);
            INPUT_XK_Manage_Tab.Controls.Add(INPUT_XK_Search_BT);

            possize.pos_x = 500;
            possize.pos_y = 6;
            INPUT_XK_Search_Txt_Lb = new TextBox_Lbl(INPUT_XK_Manage_Tab, "              Search", TextBox_Type.TEXT, possize, AnchorType.RIGHT);

            return true;
        }
    }
}