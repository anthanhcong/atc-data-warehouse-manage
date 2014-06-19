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
        /**************************************/
        //            Stock Manage            //
        /*************************************/


        private System.Windows.Forms.TabPage Stock_Manage_Tab;
        private Gridview_Grp Stock_Table_Form;
        private Button Stock_Import_BT;
        public Button Stock_Store_BT;
        public Button Stock_Process_BT;
        public Button Stock_Search_BT;
        private TextBox_Lbl Stock_Search_Txt_Lb;
        private ComboBox_Lbl Stock_WH_ID_List_cbx;
        private ComboBox_Lbl Stock_Part_Number_cbx;
        private TextBox_Lbl Stock_Bin_Txt;
        private TextBox_Lbl Stock_Plant_Txt;

        public DataTable Load_Stock_TBL;
        //DataSet Load_Stock_ds = new DataSet();
        //SqlDataAdapter Load_Stock_da;

        public DataTable Load_Ma_List_Tbl;
        public DataSet Load_Ma_List_Tbl_ds = new DataSet();
        public SqlDataAdapter Load_Ma_List_Tbl_da;

        public DataTable Load_WH_ID_List_Tbl;
        public DataSet Load_WH_ID_List_Tbl_ds = new DataSet();
        public SqlDataAdapter Load_WH_ID_List_Tbl_da;

        private void Create_Stock_Manage_Tab()
        {
            PosSize possize = new PosSize();
            string tab_name = "Stock Manage Tab";

            Stock_Manage_Tab = new System.Windows.Forms.TabPage();
            Stock_Manage_Tab.Text = tab_name;
            Stock_Manage_Tab.SuspendLayout();
            Stock_Manage_Tab.Location = new System.Drawing.Point(4, 22);
            Stock_Manage_Tab.Size = new System.Drawing.Size(900, 390);
            Stock_Manage_Tab.Padding = new System.Windows.Forms.Padding(3);
            //Stock_Manage_Tab.TabIndex = 1;
            Stock_Manage_Tab.UseVisualStyleBackColor = true;
            Stock_Manage_Tab.ResumeLayout(true);
            Stock_Manage_Tab.PerformLayout();
            this.MainTabControl.Controls.Add(this.Stock_Manage_Tab);

            // Init Card Table
            possize.pos_x = 6;
            possize.pos_y = 150;
            possize.width = Stock_Manage_Tab.Size.Width - 15;
            possize.height = Stock_Manage_Tab.Size.Height - 160;
            Stock_Table_Form = new Gridview_Grp(Stock_Manage_Tab, "Stock Manage Table", possize, AUTO_RESIZE,
                                                Database_WHM_Stock_Con_Str, @"SELECT * FROM dbo.Material_Stock_tb", AnchorType.LEFT);
            Stock_Table_Form.Load_DataBase(Database_WHM_Stock_Con_Str, @"SELECT * FROM dbo.Material_Stock_tb");
            Stock_Init_BT();
        }

        public bool Stock_Init_BT()
        {
            PosSize possize = new PosSize();
            Stock_Import_BT = new Button();
            Stock_Store_BT = new Button();
            Stock_Process_BT = new Button();
            Stock_Search_BT = new Button();


            Stock_Import_BT.Name = "Stock_Import_BT";
            Stock_Import_BT.Text = "Import";
            Stock_Import_BT.Location = new System.Drawing.Point(220, 76);
            Stock_Import_BT.Size = new System.Drawing.Size(50, 20);
            Stock_Import_BT.UseVisualStyleBackColor = true;
            Stock_Import_BT.Click += new System.EventHandler(Stock_Import_BT_Click);
            Stock_Manage_Tab.Controls.Add(Stock_Import_BT);

            Stock_Store_BT.Name = "Stock_Store_BT";
            Stock_Store_BT.Text = "Save";
            Stock_Store_BT.Location = new System.Drawing.Point(220, 106);
            Stock_Store_BT.Size = new System.Drawing.Size(50, 20);
            Stock_Store_BT.UseVisualStyleBackColor = true;
            Stock_Store_BT.Click += new System.EventHandler(Stock_Store_BT_Click);
            Stock_Manage_Tab.Controls.Add(Stock_Store_BT);

            Stock_Process_BT.Name = "Stock_Process_BT";
            Stock_Process_BT.Text = "Process";
            Stock_Process_BT.Location = new System.Drawing.Point(340, 109);
            Stock_Process_BT.Size = new System.Drawing.Size(60, 23);
            Stock_Process_BT.UseVisualStyleBackColor = true;
            Stock_Process_BT.Click += new System.EventHandler(Stock_Process_BT_Click);
            //Stock_Manage_Tab.Controls.Add(Stock_Process_BT);

            Stock_Search_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            Stock_Search_BT.Name = "Stock_Search_BT";
            Stock_Search_BT.Text = "Search";
            Stock_Search_BT.Location = new System.Drawing.Point(600, 35);
            Stock_Search_BT.Size = new System.Drawing.Size(60, 23);
            Stock_Search_BT.UseVisualStyleBackColor = true;
            Stock_Search_BT.Click += new System.EventHandler(Stock_Search_BT_Click);
            Stock_Manage_Tab.Controls.Add(Stock_Search_BT);

            Load_WH_ID_List(); 
            possize.pos_x = 10;
            possize.pos_y = 46;
            Stock_WH_ID_List_cbx = new ComboBox_Lbl(Stock_Manage_Tab, "WH ID List", possize, Load_WH_ID_List_Tbl, "WareHouse_ID", "WareHouse_ID", AnchorType.RIGHT);
            Stock_WH_ID_List_cbx.My_Combo.Size = new Size(100, 20);

            Load_Material_List();
            possize.pos_x = 10;
            possize.pos_y = 16;
            Stock_Part_Number_cbx = new ComboBox_Lbl(Stock_Manage_Tab, "Part Number", possize, Load_Ma_List_Tbl, "Part_Number", "Part_Number", AnchorType.RIGHT);
            Stock_Part_Number_cbx.My_Combo.Size = new Size(100, 20);

            possize.pos_x = 10;
            possize.pos_y = 76;
            Stock_Bin_Txt = new TextBox_Lbl(Stock_Manage_Tab, "Bin", TextBox_Type.TEXT, possize, AnchorType.RIGHT);
            Stock_Bin_Txt.My_TextBox.Size = new Size(100, 20);

            possize.pos_x = 10;
            possize.pos_y = 106;
            Stock_Plant_Txt = new TextBox_Lbl(Stock_Manage_Tab, "Plant", TextBox_Type.TEXT, possize, AnchorType.RIGHT);
            Stock_Plant_Txt.My_TextBox.Size = new Size(100, 20);

            possize.pos_x = 500;
            possize.pos_y = 6;
            Stock_Search_Txt_Lb = new TextBox_Lbl(Stock_Manage_Tab, "              Search", TextBox_Type.TEXT, possize, AnchorType.RIGHT);

            return true;
        }

                    /**************************************/
                    //          WH_ID_List Manage         //
                    /*************************************/


        private System.Windows.Forms.TabPage WH_List_and_M_List_Tab;
        private Gridview_Grp WH_ID_List_Table_Form;
        private Gridview_Grp Material_List_Table_Form;
        private Button WH_ID_List_Import_BT;
        public Button WH_ID_List_Store_BT;
        public Button WH_ID_List_Process_BT;
        public Button WH_ID_List_Search_BT;
        private TextBox_Lbl WH_ID_List_Search_Txt_Lb;
        private Button Ma_List_Import_BT;
        public Button Ma_List_Store_BT;
        private Checkbox_Lbl Ma_List_Import_Auto_Update;
        private Checkbox_Lbl WH_List_Import_Auto_Update;

        public DataTable Load_WH_ID_List_TBL;
        //DataSet Load_WH_ID_List_ds = new DataSet();
        //SqlDataAdapter Load_WH_ID_List_da;

        private void Create_WH_List_and_M_List_Tab()
        {
            PosSize possize = new PosSize();
            string tab_name = "WH_List and Material List Tab";

            WH_List_and_M_List_Tab = new System.Windows.Forms.TabPage();
            WH_List_and_M_List_Tab.Text = tab_name;
            WH_List_and_M_List_Tab.SuspendLayout();
            WH_List_and_M_List_Tab.Location = new System.Drawing.Point(4, 22);
            WH_List_and_M_List_Tab.Size = new System.Drawing.Size(900, 390);
            WH_List_and_M_List_Tab.Padding = new System.Windows.Forms.Padding(3);
            //WH_List_and_M_List_Tab.TabIndex = 1;
            WH_List_and_M_List_Tab.UseVisualStyleBackColor = true;
            WH_List_and_M_List_Tab.ResumeLayout(true);
            WH_List_and_M_List_Tab.PerformLayout();
            this.MainTabControl.Controls.Add(this.WH_List_and_M_List_Tab);

            // Init Card Table
            possize.pos_x = 6;
            possize.pos_y = 140;
            possize.width = 300;
            possize.height = WH_List_and_M_List_Tab.Size.Height - 150;
            WH_ID_List_Table_Form = new Gridview_Grp(WH_List_and_M_List_Tab, "WH_ID_List Table", possize, NO_AUTO_RESIZE,
                                                Database_WHM_Stock_Con_Str, @"SELECT * FROM dbo.Warehouse_List_tb", AnchorType.LEFT);
            WH_ID_List_Table_Form.Tab_Grp.Anchor = (AnchorStyles)(AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Top);
            WH_ID_List_Table_Form.Load_DataBase(Database_WHM_Stock_Con_Str, @"SELECT * FROM dbo.Warehouse_List_tb");

            possize.pos_x = 320;
            possize.pos_y = 140;
            possize.width = WH_List_and_M_List_Tab.Size.Width - 340;
            possize.height = WH_List_and_M_List_Tab.Size.Height - 150;
            Material_List_Table_Form = new Gridview_Grp(WH_List_and_M_List_Tab, "Material List Table", possize, AUTO_RESIZE,
                                                Database_WHM_Stock_Con_Str, @"SELECT * FROM dbo.Material_List_tb", AnchorType.NONE);
            Material_List_Table_Form.Load_DataBase(Database_WHM_Stock_Con_Str, @"SELECT * FROM dbo.Material_List_tb");
            WH_ID_List_Init_BT();
        }

        public bool WH_ID_List_Init_BT()
        {
            PosSize possize = new PosSize();
            WH_ID_List_Import_BT = new Button();
            WH_ID_List_Store_BT = new Button();
            WH_ID_List_Process_BT = new Button();
            WH_ID_List_Search_BT = new Button();
            Ma_List_Import_BT = new Button();
            Ma_List_Store_BT = new Button();

            Ma_List_Import_BT.Name = "Ma_List_Import_BT";
            Ma_List_Import_BT.Text = "Ma Import";
            Ma_List_Import_BT.Location = new System.Drawing.Point(340, 69);
            Ma_List_Import_BT.Size = new System.Drawing.Size(65, 23);
            Ma_List_Import_BT.UseVisualStyleBackColor = true;
            Ma_List_Import_BT.Click += new System.EventHandler(Ma_List_Import_BT_Click);
            WH_List_and_M_List_Tab.Controls.Add(Ma_List_Import_BT);

            Ma_List_Store_BT.Name = "Ma_List_Store_BT";
            Ma_List_Store_BT.Text = "Ma Save";
            Ma_List_Store_BT.Location = new System.Drawing.Point(425, 69);
            Ma_List_Store_BT.Size = new System.Drawing.Size(60, 23);
            Ma_List_Store_BT.UseVisualStyleBackColor = true;
            Ma_List_Store_BT.Click += new System.EventHandler(Ma_List_Store_BT_Click);
            WH_List_and_M_List_Tab.Controls.Add(Ma_List_Store_BT);

            WH_ID_List_Import_BT.Name = "WH_ID_List_Import_BT";
            WH_ID_List_Import_BT.Text = "WH Import";
            WH_ID_List_Import_BT.Location = new System.Drawing.Point(20, 69);
            WH_ID_List_Import_BT.Size = new System.Drawing.Size(70, 23);
            WH_ID_List_Import_BT.UseVisualStyleBackColor = true;
            WH_ID_List_Import_BT.Click += new System.EventHandler(WH_ID_List_Import_BT_Click);
            WH_List_and_M_List_Tab.Controls.Add(WH_ID_List_Import_BT);

            WH_ID_List_Store_BT.Name = "WH_ID_List_Store_BT";
            WH_ID_List_Store_BT.Text = "WH Save";
            WH_ID_List_Store_BT.Location = new System.Drawing.Point(110, 69);
            WH_ID_List_Store_BT.Size = new System.Drawing.Size(65, 23);
            WH_ID_List_Store_BT.UseVisualStyleBackColor = true;
            WH_ID_List_Store_BT.Click += new System.EventHandler(WH_ID_List_Store_BT_Click);
            WH_List_and_M_List_Tab.Controls.Add(WH_ID_List_Store_BT);

            possize.pos_x = 20;
            possize.pos_y = 100;
            WH_List_Import_Auto_Update = new Checkbox_Lbl(WH_List_and_M_List_Tab, "Auto overwrite when import", possize, AnchorType.LEFT);
            WH_List_Import_Auto_Update.My_CheckBox.Checked = false;

            possize.pos_x = 340;
            possize.pos_y = 100;
            Ma_List_Import_Auto_Update = new Checkbox_Lbl(WH_List_and_M_List_Tab, "Auto overwrite when import", possize, AnchorType.LEFT);
            Ma_List_Import_Auto_Update.My_CheckBox.Checked = false;

            WH_ID_List_Search_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            WH_ID_List_Search_BT.Name = "WH_ID_List_Search_BT";
            WH_ID_List_Search_BT.Text = "Search";
            WH_ID_List_Search_BT.Location = new System.Drawing.Point(600, 35);
            WH_ID_List_Search_BT.Size = new System.Drawing.Size(60, 23);
            WH_ID_List_Search_BT.UseVisualStyleBackColor = true;
            WH_ID_List_Search_BT.Click += new System.EventHandler(WH_ID_List_Search_BT_Click);
            WH_List_and_M_List_Tab.Controls.Add(WH_ID_List_Search_BT);

            possize.pos_x = 500;
            possize.pos_y = 6;
            WH_ID_List_Search_Txt_Lb = new TextBox_Lbl(WH_List_and_M_List_Tab, "              Search", TextBox_Type.TEXT, possize, AnchorType.RIGHT);

            return true;
        }
    }
}