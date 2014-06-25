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
        private System.Windows.Forms.TabPage WH_Daskboard_Tab;
        private Gridview_Grp WH_Daskboard_Table_Form;
        private Gridview_Grp WH_ID_with_MaLH_Table_Form;
        private Button_Lbl WH_Daskboard_Create_BT;
        private Button_Lbl WH_Daskboard_Store_BT;
        //public Button_Lbl WH_Daskboard_Search_BT;
        private TextBox_Lbl WH_Daskboard_WH_ID_TbxL;
        private Label WH_Daskboard_WH_Name_Lb;
        private TextBox_Lbl WH_Daskboard_WH_Name_TbxL;
        private ComboBox_Lbl WH_Daskboard_Mother_WH_CbxL;
        private Checkbox_Lbl WH_Daskboard_Check_Import;
        private TextBox_Lbl WH_Daskboard_Search_Txt_Lb;
        private TextBox_Lbl WH_Daskboard_Note_Txt_Lb;
        private TextBox_Lbl WH_ID_with_MaLH_ID_TbxL;
        private TextBox_Lbl WH_ID_with_MaLH_Ma_LH_TbxL;
        private ComboBox_Lbl WH_ID_with_MaLH_WH_ID_CbxL;
        private TextBox_Lbl WH_ID_with_MaLH_Im_or_Ex_TbxL;
        private TextBox_Lbl WH_ID_with_MaLH_Ty_le_TbxL;
        private Button_Lbl WH_ID_with_MLH_Create_BT;
        private Button_Lbl WH_ID_with_MLH_Store_BT;
        private GroupBox WH_Daskboard_Create_gbx;
        private GroupBox WH_ID_with_MaLH_Create_gbx;

        public DataTable Load_WH_Daskboard_TBL;
        //DataSet Load_WH_Daskboard_ds = new DataSet();
        //SqlDataAdapter Load_WH_Daskboard_da;

        private void Create_WH_Daskboard_Tab()
        {
            PosSize possize = new PosSize();
            string tab_name = "Warehouse Dashboard Tab";

            WH_Daskboard_Tab = new System.Windows.Forms.TabPage();
            WH_Daskboard_Tab.Text = tab_name;
            WH_Daskboard_Tab.SuspendLayout();
            WH_Daskboard_Tab.Location = new System.Drawing.Point(4, 22);
            WH_Daskboard_Tab.Size = new System.Drawing.Size(900, 390);
            WH_Daskboard_Tab.Padding = new System.Windows.Forms.Padding(3);
            //WH_Daskboard_Tab.TabIndex = 1;
            WH_Daskboard_Tab.UseVisualStyleBackColor = true;
            WH_Daskboard_Tab.ResumeLayout(true);
            WH_Daskboard_Tab.PerformLayout();
            this.MainTabControl.Controls.Add(this.WH_Daskboard_Tab);

            // Init Card Table
            possize.pos_x = 6;
            possize.pos_y = 200;
            possize.width = 300;
            possize.height = WH_Daskboard_Tab.Size.Height - 210;
            WH_Daskboard_Table_Form = new Gridview_Grp(WH_Daskboard_Tab, "WH_Daskboard Table", possize, NO_AUTO_RESIZE,
                                                Database_WHM_Stock_Con_Str, @"SELECT * FROM dbo.Warehouse_Dashboard_tb", AnchorType.LEFT);
            WH_Daskboard_Table_Form.Tab_Grp.Anchor = (AnchorStyles)(AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Top);
            WH_Daskboard_Table_Form.Load_DataBase(Database_WHM_Stock_Con_Str, @"SELECT * FROM dbo.Warehouse_Dashboard_tb");

            possize.pos_x = 320;
            possize.pos_y = 200;
            possize.width = WH_Daskboard_Tab.Size.Width - 340;
            possize.height = WH_Daskboard_Tab.Size.Height - 210;
            WH_ID_with_MaLH_Table_Form = new Gridview_Grp(WH_Daskboard_Tab, "WH_ID with MLH Table", possize, AUTO_RESIZE,
                                                Database_WHM_Stock_Con_Str, @"SELECT * FROM dbo.List_WH_ID_MaLH_tb", AnchorType.NONE);
            WH_ID_with_MaLH_Table_Form.Load_DataBase(Database_WHM_Stock_Con_Str, @"SELECT * FROM dbo.List_WH_ID_MaLH_tb");
            WH_Daskboard_Init();
        }

        public bool WH_Daskboard_Init()
        {
            PosSize possize = new PosSize();
            WH_Daskboard_Create_gbx = new GroupBox();
            WH_ID_with_MaLH_Create_gbx = new GroupBox();
            //WH_Daskboard_WH_Name_Lb = new Label();

            WH_Daskboard_Tab.Controls.Add(WH_Daskboard_Create_gbx);
            WH_Daskboard_Create_gbx.Location = new System.Drawing.Point(10, 6);
            WH_Daskboard_Create_gbx.Name = "WH_Daskboard_Create_gbx";
            WH_Daskboard_Create_gbx.Size = new System.Drawing.Size (324, 180);
            WH_Daskboard_Create_gbx.TabIndex = 1;
            WH_Daskboard_Create_gbx.TabStop = false;
            WH_Daskboard_Create_gbx.Text = "WH Dashboard Group";

            WH_Daskboard_Tab.Controls.Add(WH_ID_with_MaLH_Create_gbx);
            WH_ID_with_MaLH_Create_gbx.Location = new System.Drawing.Point(360, 6);
            WH_ID_with_MaLH_Create_gbx.Name = "WH_ID_with_MaLH_Create_gbx";
            WH_ID_with_MaLH_Create_gbx.Size = new System.Drawing.Size(324, 180);
            WH_ID_with_MaLH_Create_gbx.TabIndex = 1;
            WH_ID_with_MaLH_Create_gbx.TabStop = false;
            WH_ID_with_MaLH_Create_gbx.Text = "WH_ID with MaLH Group";

            possize.pos_x = 256;
            possize.pos_y = 20;// 51;
            WH_Daskboard_Create_BT = new Button_Lbl(WH_Daskboard_Tab, "Create", possize, AnchorType.LEFT);
            WH_Daskboard_Create_BT.My_Button.Size = new Size(55, 23);
            WH_Daskboard_Create_BT.My_Button.Click += new System.EventHandler(WH_Daskboard_Create_BT_Click);
            WH_Daskboard_Create_gbx.Controls.Add(WH_Daskboard_Create_BT.My_Button);

            possize.pos_x = 256;
            possize.pos_y = 81;
            WH_Daskboard_Store_BT = new Button_Lbl(WH_Daskboard_Tab, "Save", possize, AnchorType.LEFT);
            WH_Daskboard_Store_BT.My_Button.Size = new Size(55, 23);
            WH_Daskboard_Store_BT.My_Button.Click += new System.EventHandler(WH_Daskboard_Store_BT_Click);
            WH_Daskboard_Create_gbx.Controls.Add(WH_Daskboard_Store_BT.My_Button);

            possize.pos_x = 258;
            possize.pos_y = 80;
            WH_ID_with_MLH_Create_BT = new Button_Lbl(WH_Daskboard_Tab, "Create", possize, AnchorType.LEFT);
            WH_ID_with_MLH_Create_BT.My_Button.Size = new Size(55, 23);
            WH_ID_with_MLH_Create_BT.My_Button.Click += new System.EventHandler(WH_ID_with_MLH_Create_BT_Click);
            
            possize.pos_x = 258;
            possize.pos_y = 111;
            WH_ID_with_MLH_Store_BT = new Button_Lbl(WH_Daskboard_Tab, "Save", possize, AnchorType.LEFT);
            WH_ID_with_MLH_Store_BT.My_Button.Size = new Size(55, 23);
            WH_ID_with_MLH_Store_BT.My_Button.Click += new System.EventHandler(WH_ID_with_MLH_Store_BT_Click);
            

            possize.pos_x = 10;
            possize.pos_y = 22;
            WH_Daskboard_WH_ID_TbxL = new TextBox_Lbl(WH_Daskboard_Tab, "WH ID", TextBox_Type.TEXT, possize, AnchorType.LEFT);
            WH_Daskboard_WH_ID_TbxL.My_TextBox.Location = new Point(100, 22);
            WH_Daskboard_WH_ID_TbxL.My_TextBox.Size = new Size(100, 20);

            possize.pos_x = 10;
            possize.pos_y = 54;
            WH_Daskboard_WH_Name_TbxL = new TextBox_Lbl(WH_Daskboard_Tab, "WH Name", TextBox_Type.TEXT, possize, AnchorType.LEFT);
            WH_Daskboard_WH_Name_TbxL.My_TextBox.Location = new Point(100, 54);
            WH_Daskboard_WH_Name_TbxL.My_TextBox.Size = new Size(210, 20);

            //WH_Daskboard_WH_Name_Lb.AutoSize = true;
            //WH_Daskboard_WH_Name_Lb.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            //WH_Daskboard_WH_Name_Lb.ForeColor = System.Drawing.Color.Blue;
            //WH_Daskboard_WH_Name_Lb.Location = new System.Drawing.Point(10, 54);
            //WH_Daskboard_WH_Name_Lb.Name = "WH_Daskboard_WH_Name_Lb";
            //WH_Daskboard_WH_Name_Lb.Size = new System.Drawing.Size(35, 13);
            //WH_Daskboard_WH_Name_Lb.TabIndex = 1;
            //WH_Daskboard_WH_Name_Lb.Text = "WH Name";

            possize.pos_x = 100;
            possize.pos_y = 83;
            WH_Daskboard_Check_Import = new Checkbox_Lbl(WH_Daskboard_Tab, "Check Import", possize, AnchorType.LEFT);
            WH_Daskboard_Check_Import.My_CheckBox.Checked = false;

            possize.pos_x = 10;
            possize.pos_y = 113;
            WH_Daskboard_Mother_WH_CbxL = new ComboBox_Lbl(WH_Daskboard_Tab, "Mother WH", possize, NK_Ma_LH_TBL, "Mother_WHID", "Mother_WHID", AnchorType.LEFT);
            WH_Daskboard_Mother_WH_CbxL.My_Combo.Location = new Point(100, 113);
            WH_Daskboard_Mother_WH_CbxL.My_Combo.Size = new Size(210, 20);
            
            possize.pos_x = 10;
            possize.pos_y = 144;
            WH_Daskboard_Note_Txt_Lb = new TextBox_Lbl(WH_Daskboard_Tab, "Note", TextBox_Type.TEXT, possize, AnchorType.LEFT);
            WH_Daskboard_Note_Txt_Lb.My_TextBox.Location = new Point(100, 144);
            WH_Daskboard_Note_Txt_Lb.My_TextBox.Size = new Size(210, 20);

            possize.pos_x = 10;
            possize.pos_y = 22;
            WH_ID_with_MaLH_ID_TbxL = new TextBox_Lbl(WH_Daskboard_Tab, "ID", TextBox_Type.TEXT, possize, AnchorType.LEFT);
            WH_ID_with_MaLH_ID_TbxL.My_TextBox.Location = new Point(130, 22);
            WH_ID_with_MaLH_ID_TbxL.My_TextBox.Size = new Size(100, 20);

            possize.pos_x = 10;
            possize.pos_y = 51;
            WH_ID_with_MaLH_Ma_LH_TbxL = new TextBox_Lbl(WH_Daskboard_Tab, "Mã loại hình", TextBox_Type.TEXT, possize, AnchorType.LEFT);
            WH_ID_with_MaLH_Ma_LH_TbxL.My_TextBox.Location = new Point(130, 51);
            WH_ID_with_MaLH_Ma_LH_TbxL.My_TextBox.Size = new Size(100, 20);

            possize.pos_x = 10;
            possize.pos_y = 113;
            WH_ID_with_MaLH_Im_or_Ex_TbxL = new TextBox_Lbl(WH_Daskboard_Tab, "IMPORT/EXPORT", TextBox_Type.TEXT, possize, AnchorType.LEFT);
            WH_ID_with_MaLH_Im_or_Ex_TbxL.My_TextBox.Location = new Point(130, 113);
            WH_ID_with_MaLH_Im_or_Ex_TbxL.My_TextBox.Size = new Size(100, 20);

            possize.pos_x = 10;
            possize.pos_y = 144;
            WH_ID_with_MaLH_Ty_le_TbxL = new TextBox_Lbl(WH_Daskboard_Tab, "Tỷ lệ (%)", TextBox_Type.TEXT, possize, AnchorType.LEFT);
            WH_ID_with_MaLH_Ty_le_TbxL.My_TextBox.Location = new Point(130, 144);
            WH_ID_with_MaLH_Ty_le_TbxL.My_TextBox.Size = new Size(100, 20);

            possize.pos_x = 10;
            possize.pos_y = 81;
            WH_ID_with_MaLH_WH_ID_CbxL = new ComboBox_Lbl(WH_Daskboard_Tab, "WH ID", possize, NK_Ma_LH_TBL, "WareHouse_ID", "WareHouse_ID", AnchorType.LEFT);
            WH_ID_with_MaLH_WH_ID_CbxL.My_Combo.Location = new Point(130, 81);
            WH_ID_with_MaLH_WH_ID_CbxL.My_Combo.Size = new Size(100, 20);

            WH_Daskboard_Create_gbx.Controls.Add(WH_Daskboard_WH_ID_TbxL.My_Label);
            WH_Daskboard_Create_gbx.Controls.Add(WH_Daskboard_WH_ID_TbxL.My_TextBox);
            WH_Daskboard_Create_gbx.Controls.Add(WH_Daskboard_WH_Name_TbxL.My_Label);
            WH_Daskboard_Create_gbx.Controls.Add(WH_Daskboard_WH_Name_TbxL.My_TextBox);
            //WH_Daskboard_Create_gbx.Controls.Add(WH_Daskboard_WH_Name_Lb);
            WH_Daskboard_Create_gbx.Controls.Add(WH_Daskboard_Note_Txt_Lb.My_Label);
            WH_Daskboard_Create_gbx.Controls.Add(WH_Daskboard_Note_Txt_Lb.My_TextBox);
            WH_Daskboard_Create_gbx.Controls.Add(WH_Daskboard_Mother_WH_CbxL.My_Label);
            WH_Daskboard_Create_gbx.Controls.Add(WH_Daskboard_Mother_WH_CbxL.My_Combo);
            WH_Daskboard_Create_gbx.Controls.Add(WH_Daskboard_Check_Import.My_CheckBox);
            
            WH_ID_with_MaLH_Create_gbx.Controls.Add(WH_ID_with_MaLH_WH_ID_CbxL.My_Label);
            WH_ID_with_MaLH_Create_gbx.Controls.Add(WH_ID_with_MaLH_WH_ID_CbxL.My_Combo);
            WH_ID_with_MaLH_Create_gbx.Controls.Add(WH_ID_with_MaLH_Ty_le_TbxL.My_Label);
            WH_ID_with_MaLH_Create_gbx.Controls.Add(WH_ID_with_MaLH_Ty_le_TbxL.My_TextBox);
            WH_ID_with_MaLH_Create_gbx.Controls.Add(WH_ID_with_MaLH_Im_or_Ex_TbxL.My_Label);
            WH_ID_with_MaLH_Create_gbx.Controls.Add(WH_ID_with_MaLH_Im_or_Ex_TbxL.My_TextBox);
            WH_ID_with_MaLH_Create_gbx.Controls.Add(WH_ID_with_MaLH_Ma_LH_TbxL.My_Label);
            WH_ID_with_MaLH_Create_gbx.Controls.Add(WH_ID_with_MaLH_Ma_LH_TbxL.My_TextBox);
            WH_ID_with_MaLH_Create_gbx.Controls.Add(WH_ID_with_MaLH_ID_TbxL.My_Label);
            WH_ID_with_MaLH_Create_gbx.Controls.Add(WH_ID_with_MaLH_ID_TbxL.My_TextBox);
            WH_ID_with_MaLH_Create_gbx.Controls.Add(WH_ID_with_MLH_Create_BT.My_Button);
            WH_ID_with_MaLH_Create_gbx.Controls.Add(WH_ID_with_MLH_Store_BT.My_Button);
            return true;
        }
    }
}