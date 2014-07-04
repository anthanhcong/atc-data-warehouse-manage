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
        private System.Windows.Forms.TabPage Mother_Stock_Manage_Tab;
        private Gridview_Grp Mother_Stock_Table_Form;
        private Gridview_Grp Mother_Stock_Total_Qty_Form;
        public Button Mother_Stock_Store_BT;
        public Button Mother_Stock_Process_BT;
        public Button Mother_Stock_Search_BT;
        private ComboBox_Lbl Mother_Stock_WH_ID_List_cbx;
        private ComboBox_Lbl Mother_Stock_Part_Number_cbx;
        private ComboBox_Lbl Mother_Stock_WH_ID_cbx;
        private ComboBox_Lbl Mother_Stock_With_wh_id_cbx;
        private TextBox_Lbl Mother_Stock_Qty_Txt;
        private TextBox_Lbl Mother_Stock_Sum_Qty_Txt;
        private ComboBox Mother_Stock_Menu_Sort_View_cbx;

        private TextBox_Lbl Mother_Stock_Bin_Txt;
        private TextBox_Lbl Mother_Stock_Plant_Txt;
        private GroupBox Mother_WH_Stock_Manage_Group_gbx;
        private Checkbox_Lbl Mother_Stock_Manage_Single_Check_WH_ID;
        private Checkbox_Lbl Mother_Stock_Manage_Single_Check_Part_Number;
        private Checkbox_Lbl Mother_Stock_Manage_Single_Check_Plant;
        private Checkbox_Lbl Mother_Stock_Manage_Single_Check_Bin;
        private Checkbox_Lbl Mother_Stock_Manage_Single_View_All;

        public DataTable Mother_Stock_WH_ID_List_Tbl;
        public DataSet Mother_Stock_WH_ID_List_ds = new DataSet();
        public SqlDataAdapter Mother_Stock_WH_ID_List_da;

        public DataTable Load_Mother_Stock_TBL;
        public DataSet Load_Mother_Stock_ds = new DataSet();
        public SqlDataAdapter Load_Mother_Stock_da;

        private void Create_Mother_Stock_Manage_Tab()
        {
            PosSize possize = new PosSize();
            string tab_name = "Mother WH Stock Manage Tab";

            Mother_Stock_Manage_Tab = new System.Windows.Forms.TabPage();
            Mother_Stock_Manage_Tab.Text = tab_name;
            Mother_Stock_Manage_Tab.SuspendLayout();
            Mother_Stock_Manage_Tab.Location = new System.Drawing.Point(4, 22);
            Mother_Stock_Manage_Tab.Size = new System.Drawing.Size(900, 390);
            Mother_Stock_Manage_Tab.Padding = new System.Windows.Forms.Padding(3);
            //Mother_Stock_Manage_Tab.TabIndex = 1;
            Mother_Stock_Manage_Tab.UseVisualStyleBackColor = true;
            Mother_Stock_Manage_Tab.ResumeLayout(true);
            Mother_Stock_Manage_Tab.PerformLayout();
            this.MainTabControl.Controls.Add(this.Mother_Stock_Manage_Tab);

            // Init Card Table
            possize.pos_x = 6;
            possize.pos_y = 140;
            possize.width = Mother_Stock_Manage_Tab.Size.Width - 280;
            possize.height = Mother_Stock_Manage_Tab.Size.Height - 150;
            Mother_Stock_Table_Form = new Gridview_Grp(Mother_Stock_Manage_Tab, "Mother_Stock Manage Table", possize, AUTO_RESIZE,
                                                Database_WHM_Stock_Con_Str, @"SELECT * FROM dbo.Material_Stock_tb", AnchorType.NONE);
            Mother_Stock_Table_Form.Load_DataBase(Database_WHM_Stock_Con_Str, @"SELECT * FROM dbo.Material_Stock_tb where Part_Number = '0'");
            Mother_Stock_Table_Form.dataGridView_View.CellDoubleClick += new DataGridViewCellEventHandler(Mother_Stock_Table_Form_CellDoubleClick);
            //Mother_Stock_Table_Form.Tab_Grp.Anchor = (AnchorStyles)(AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Top);

            possize.pos_x = 480;
            possize.pos_y = 140;
            possize.width = 250;
            possize.height = Mother_Stock_Manage_Tab.Size.Height - 150;
            Mother_Stock_Total_Qty_Form = new Gridview_Grp(Mother_Stock_Manage_Tab, "Total Qty Table", possize, NO_AUTO_RESIZE,
                                                Database_WHM_Stock_Con_Str, @"SELECT * FROM dbo.Material_Stock_tb", AnchorType.NONE);
            Mother_Stock_Total_Qty_Form.Tab_Grp.Anchor = (AnchorStyles)(AnchorStyles.Bottom | AnchorStyles.Right | AnchorStyles.Top);
            //Mother_Stock_Total_Qty_Form.Load_DataBase(Database_WHM_Stock_Con_Str, @"SELECT * FROM dbo.Material_Stock_tb");
            Mother_Stock_Total_Qty_Form.Review_BT.Visible = false;
            Mother_Stock_Total_Qty_Form.Delete_All_BT.Visible = false;
            Mother_Stock_Total_Qty_Form.Delete_Rows_BT.Visible = false;
            Mother_Stock_Total_Qty_Form.Export_BT.Visible = false;

            Mother_Stock_Init_BT();
        }

        public bool Mother_Stock_Init_BT()
        {
            PosSize possize = new PosSize();
            //Mother_Stock_Import_BT = new Button();
            Mother_Stock_Store_BT = new Button();
            Mother_Stock_Process_BT = new Button();
            Mother_Stock_Search_BT = new Button();
            Mother_Stock_Menu_Sort_View_cbx = new ComboBox();
            Mother_WH_Stock_Manage_Group_gbx = new GroupBox();

            Mother_Stock_Manage_Tab.Controls.Add(Mother_WH_Stock_Manage_Group_gbx);
            Mother_WH_Stock_Manage_Group_gbx.Location = new System.Drawing.Point(6, 6);
            Mother_WH_Stock_Manage_Group_gbx.Name = "Mother_WH_Stock_Manage_Group_gbx";
            Mother_WH_Stock_Manage_Group_gbx.Size = new System.Drawing.Size(550, 122);
            Mother_WH_Stock_Manage_Group_gbx.TabIndex = 1;
            Mother_WH_Stock_Manage_Group_gbx.TabStop = false;
            Mother_WH_Stock_Manage_Group_gbx.Text = "Manage Mother_WH_Stock Group";

            Mother_Stock_Store_BT.Name = "Mother_Stock_Store_BT";
            Mother_Stock_Store_BT.Text = "Save";
            Mother_Stock_Store_BT.Location = new System.Drawing.Point(485, 90);
            Mother_Stock_Store_BT.Size = new System.Drawing.Size(55, 20);
            Mother_Stock_Store_BT.UseVisualStyleBackColor = true;
            Mother_Stock_Store_BT.Click += new System.EventHandler(Mother_Stock_Store_BT_Click);
            //Mother_Stock_Manage_Tab.Controls.Add(Mother_Stock_Store_BT);
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_Store_BT);

            Mother_Stock_Process_BT.Name = "Mother_Stock_Process_BT";
            Mother_Stock_Process_BT.Text = "Process";
            Mother_Stock_Process_BT.Location = new System.Drawing.Point(340, 109);
            Mother_Stock_Process_BT.Size = new System.Drawing.Size(60, 23);
            Mother_Stock_Process_BT.UseVisualStyleBackColor = true;
            //Mother_Stock_Process_BT.Click += new System.EventHandler(Mother_Stock_Process_BT_Click);
            //Mother_Stock_Manage_Tab.Controls.Add(Mother_Stock_Process_BT);

            Mother_Stock_Search_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            Mother_Stock_Search_BT.Name = "Mother_Stock_Search_BT";
            Mother_Stock_Search_BT.Text = "Search";
            Mother_Stock_Search_BT.Location = new System.Drawing.Point(485, 42);
            Mother_Stock_Search_BT.Size = new System.Drawing.Size(55, 20);
            Mother_Stock_Search_BT.UseVisualStyleBackColor = true;
            Mother_Stock_Search_BT.Click += new System.EventHandler(Mother_Stock_Search_BT_Click);
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_Search_BT);

            possize.pos_x = 485;
            possize.pos_y = 20;
            Mother_Stock_Manage_Single_View_All = new Checkbox_Lbl(Mother_Stock_Manage_Tab, "View All", possize, AnchorType.LEFT);
            Mother_Stock_Manage_Single_View_All.My_CheckBox.Checked = false;
            //Mother_Stock_Manage_Single_View_All.My_CheckBox.CheckedChanged += new EventHandler(Mother_Stock_Manage_Single_View_All_CheckedChanged);
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_Manage_Single_View_All.My_CheckBox);

            possize.pos_x = 380;
            possize.pos_y = 44;
            Mother_Stock_Manage_Single_Check_WH_ID = new Checkbox_Lbl(Mother_Stock_Manage_Tab, "Select WH_ID", possize, AnchorType.LEFT);
            Mother_Stock_Manage_Single_Check_WH_ID.My_CheckBox.Checked = false;
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_Manage_Single_Check_WH_ID.My_CheckBox);

            possize.pos_x = 380;
            possize.pos_y = 20;
            Mother_Stock_Manage_Single_Check_Part_Number = new Checkbox_Lbl(Mother_Stock_Manage_Tab, "Select Part", possize, AnchorType.LEFT);
            Mother_Stock_Manage_Single_Check_Part_Number.My_CheckBox.Checked = false;
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_Manage_Single_Check_Part_Number.My_CheckBox);

            possize.pos_x = 380;
            possize.pos_y = 68;
            Mother_Stock_Manage_Single_Check_Bin = new Checkbox_Lbl(Mother_Stock_Manage_Tab, "Select Bin", possize, AnchorType.LEFT);
            Mother_Stock_Manage_Single_Check_Bin.My_CheckBox.Checked = false;
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_Manage_Single_Check_Bin.My_CheckBox);

            possize.pos_x = 380;
            possize.pos_y = 92;
            Mother_Stock_Manage_Single_Check_Plant = new Checkbox_Lbl(Mother_Stock_Manage_Tab, "Selest Plant", possize, AnchorType.LEFT);
            Mother_Stock_Manage_Single_Check_Plant.My_CheckBox.Checked = false;
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_Manage_Single_Check_Plant.My_CheckBox);

            Mother_WH_Load_Part_Number();
            possize.pos_x = 10;
            possize.pos_y = 42;
            Mother_Stock_WH_ID_List_cbx = new ComboBox_Lbl(Mother_Stock_Manage_Tab, "List WH ID", possize, Mother_Stock_WH_ID_List_Tbl, "WareHouse_ID", "WareHouse_ID", AnchorType.LEFT);
            Mother_Stock_WH_ID_List_cbx.My_Combo.Size = new Size(100, 20);
            Mother_Stock_WH_ID_List_cbx.My_Combo.KeyDown += new KeyEventHandler(Mother_Stock_WH_ID_List_cbx_KeyDown);
            Mother_Stock_WH_ID_List_cbx.My_Combo.SelectedValueChanged += new EventHandler(Mother_Stock_WH_ID_List_cbx_SelectedValueChanged);
            //Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_WH_ID_List_cbx.My_Label);
            //Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_WH_ID_List_cbx.My_Combo);

            possize.pos_x = 10;
            possize.pos_y = 16;
            Mother_Stock_Part_Number_cbx = new ComboBox_Lbl(Mother_Stock_Manage_Tab, "Part Number", possize, Mother_WH_List_Part_Tbl, "Part_Number", "Part_Number", AnchorType.LEFT);
            Mother_Stock_Part_Number_cbx.My_Combo.Size = new Size(100, 20);
            Mother_Stock_Part_Number_cbx.My_Combo.KeyDown += new KeyEventHandler(Mother_Stock_Part_Number_cbx_KeyDown);
            Mother_Stock_Part_Number_cbx.My_Combo.SelectedValueChanged += new EventHandler(Mother_Stock_Part_Number_cbx_SelectedValueChanged);
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_Part_Number_cbx.My_Label);
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_Part_Number_cbx.My_Combo);

            Mother_WH_List_WH_with_part();
            possize.pos_x = 10;
            possize.pos_y = 42;
            Mother_Stock_With_wh_id_cbx = new ComboBox_Lbl(Mother_Stock_Manage_Tab, "With wh_id", possize, List_WH_wiht_part_TBL, "WareHouse_ID", "WareHouse_ID", AnchorType.LEFT);
            Mother_Stock_With_wh_id_cbx.My_Combo.Size = new Size(100, 20);
            Mother_Stock_With_wh_id_cbx.My_Combo.KeyDown += new KeyEventHandler(Mother_Stock_With_wh_id_cbx_KeyDown);
            Mother_Stock_With_wh_id_cbx.My_Combo.SelectedValueChanged += new EventHandler(Mother_Stock_With_wh_id_cbx_SelectedValueChanged);
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_With_wh_id_cbx.My_Label);
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_With_wh_id_cbx.My_Combo);

            possize.pos_x = 210;
            possize.pos_y = 90;
            Mother_Stock_WH_ID_cbx = new ComboBox_Lbl(Mother_Stock_Manage_Tab, "WH_ID", possize, Mother_Stock_WH_ID_List_Tbl, "WareHouse_ID", "WareHouse_ID", AnchorType.LEFT);
            Mother_Stock_WH_ID_cbx.My_Combo.Location = new Point(265, 90);
            Mother_Stock_WH_ID_cbx.My_Combo.Size = new Size(100, 20);
            //Mother_Stock_WH_ID_cbx.My_Combo.KeyDown += new KeyEventHandler(Mother_Stock_WH_ID_cbx_KeyDown);
            Mother_Stock_WH_ID_cbx.My_Combo.SelectedValueChanged += new EventHandler(Mother_Stock_WH_ID_List_cbx_SelectedValueChanged);
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_WH_ID_cbx.My_Label);
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_WH_ID_cbx.My_Combo);

            possize.pos_x = 10;
            possize.pos_y = 66;
            Mother_Stock_Qty_Txt = new TextBox_Lbl(Mother_Stock_Manage_Tab, "Qty", TextBox_Type.TEXT, possize, AnchorType.LEFT);
            Mother_Stock_Qty_Txt.My_TextBox.Size = new Size(100, 20);
            //Mother_Stock_Qty_Txt.My_TextBox.KeyDown += new KeyEventHandler(Mother_Stock_Qty_Txt_KeyDown);
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_Qty_Txt.My_Label);
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_Qty_Txt.My_TextBox);

            possize.pos_x = 10;
            possize.pos_y = 90;
            Mother_Stock_Sum_Qty_Txt = new TextBox_Lbl(Mother_Stock_Manage_Tab, "Total Qty", TextBox_Type.TEXT, possize, AnchorType.LEFT);
            Mother_Stock_Sum_Qty_Txt.My_TextBox.Size = new Size(100, 20);
            //Mother_Stock_Sum_Qty_Txt.My_TextBox.KeyDown += new KeyEventHandler(Mother_Stock_Sum_Qty_Txt_KeyDown);
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_Sum_Qty_Txt.My_Label);
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_Sum_Qty_Txt.My_TextBox);

            possize.pos_x = 210;
            possize.pos_y = 42;
            Mother_Stock_Bin_Txt = new TextBox_Lbl(Mother_Stock_Manage_Tab, "Bin", TextBox_Type.TEXT, possize, AnchorType.LEFT);
            Mother_Stock_Bin_Txt.My_TextBox.Location = new Point(265, 42);
            Mother_Stock_Bin_Txt.My_TextBox.Size = new Size(100, 20);
            Mother_Stock_Bin_Txt.My_TextBox.KeyDown += new KeyEventHandler(Mother_Stock_Bin_Txt_KeyDown);
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_Bin_Txt.My_Label);
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_Bin_Txt.My_TextBox);

            possize.pos_x = 210;
            possize.pos_y = 66;
            Mother_Stock_Plant_Txt = new TextBox_Lbl(Mother_Stock_Manage_Tab, "Plant", TextBox_Type.TEXT, possize, AnchorType.LEFT);
            Mother_Stock_Plant_Txt.My_TextBox.Location = new Point(265, 66);
            Mother_Stock_Plant_Txt.My_TextBox.Size = new Size(100, 20);
            Mother_Stock_Plant_Txt.My_TextBox.KeyDown += new KeyEventHandler(Mother_Stock_Plant_Txt_KeyDown);
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_Plant_Txt.My_Label);
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_Plant_Txt.My_TextBox);

            Mother_Stock_Menu_Sort_View_cbx.FormattingEnabled = true;
            Mother_Stock_Menu_Sort_View_cbx.Items.AddRange(new object[] {
            "Single WH",
            "Sub_WH",
            "Mother and Sub"});
            Mother_Stock_Menu_Sort_View_cbx.Location = new System.Drawing.Point(265, 16);
            Mother_Stock_Menu_Sort_View_cbx.Name = "Menu_Sort_View_cbx";
            Mother_Stock_Menu_Sort_View_cbx.Size = new System.Drawing.Size(100, 20);
            //Mother_Stock_Menu_Sort_View_cbx.TabIndex = 7;
            Mother_Stock_Menu_Sort_View_cbx.Text = "Select Sort WH";
            Mother_WH_Stock_Manage_Group_gbx.Controls.Add(Mother_Stock_Menu_Sort_View_cbx);
            return true;
        }
    }
}