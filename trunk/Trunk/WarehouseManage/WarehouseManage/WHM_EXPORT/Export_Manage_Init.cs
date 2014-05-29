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
        //            Export Manage            //
        /*************************************/


        private System.Windows.Forms.TabPage Export_Manage_Tab;
        private Gridview_Grp Export_Table_Form;
        private Button Export_Create_BT;
        public Button Export_Store_BT;
        public Button Export_Process_BT;
        public Button Export_Search_BT;
        private TextBox_Lbl Export_Search_Txt_Lb;

        public DataTable Load_Export_TBL;
        //DataSet Load_Export_ds = new DataSet();
        //SqlDataAdapter Load_Export_da;

        private void Create_Export_Manage_Tab()
        {
            PosSize possize = new PosSize();
            string tab_name = "Export Manage Tab";

            Export_Manage_Tab = new System.Windows.Forms.TabPage();
            Export_Manage_Tab.Text = tab_name;
            Export_Manage_Tab.SuspendLayout();
            Export_Manage_Tab.Location = new System.Drawing.Point(4, 22);
            Export_Manage_Tab.Size = new System.Drawing.Size(900, 390);
            Export_Manage_Tab.Padding = new System.Windows.Forms.Padding(3);
            //Export_Manage_Tab.TabIndex = 1;
            Export_Manage_Tab.UseVisualStyleBackColor = true;
            Export_Manage_Tab.ResumeLayout(true);
            Export_Manage_Tab.PerformLayout();
            this.MainTabControl.Controls.Add(this.Export_Manage_Tab);

            // Init Card Table
            possize.pos_x = 6;
            possize.pos_y = 100;
            possize.width = Export_Manage_Tab.Size.Width - 15;
            possize.height = Export_Manage_Tab.Size.Height - 100;
            Export_Table_Form = new Gridview_Grp(Export_Manage_Tab, "Export Manage Table", possize, AUTO_RESIZE,
                                                Database_WHM_Export_Con_Str, @"SELECT * FROM dbo.Production_Export_tb", AnchorType.LEFT);
            Export_Table_Form.Load_DataBase(Database_WHM_Export_Con_Str, @"SELECT * FROM dbo.Production_Export_tb");
            Export_Init_BT();
        }

        public bool Export_Init_BT()
        {
            PosSize possize = new PosSize();
            Export_Create_BT = new Button();
            Export_Store_BT = new Button();
            Export_Process_BT = new Button();
            Export_Search_BT = new Button();


            Export_Create_BT.Name = "Export_Create_BT";
            Export_Create_BT.Text = "Create";
            Export_Create_BT.Location = new System.Drawing.Point(55, 69);
            Export_Create_BT.Size = new System.Drawing.Size(50, 23);
            Export_Create_BT.UseVisualStyleBackColor = true;
            Export_Create_BT.Click += new System.EventHandler(Export_Create_BT_Click);
            Export_Manage_Tab.Controls.Add(Export_Create_BT);

            Export_Store_BT.Name = "Export_Store_BT";
            Export_Store_BT.Text = "Save";
            Export_Store_BT.Location = new System.Drawing.Point(260, 69);
            Export_Store_BT.Size = new System.Drawing.Size(50, 23);
            Export_Store_BT.UseVisualStyleBackColor = true;
            Export_Store_BT.Click += new System.EventHandler(Export_Store_BT_Click);
            Export_Manage_Tab.Controls.Add(Export_Store_BT);

            Export_Process_BT.Name = "Export_Process_BT";
            Export_Process_BT.Text = "Process";
            Export_Process_BT.Location = new System.Drawing.Point(150, 69);
            Export_Process_BT.Size = new System.Drawing.Size(60, 23);
            Export_Process_BT.UseVisualStyleBackColor = true;
            Export_Process_BT.Click += new System.EventHandler(Export_Process_BT_Click);
            Export_Manage_Tab.Controls.Add(Export_Process_BT);

            Export_Search_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            Export_Search_BT.Name = "Export_Search_BT";
            Export_Search_BT.Text = "Search";
            Export_Search_BT.Location = new System.Drawing.Point(600, 35);
            Export_Search_BT.Size = new System.Drawing.Size(60, 23);
            Export_Search_BT.UseVisualStyleBackColor = true;
            Export_Search_BT.Click += new System.EventHandler(Export_Search_BT_Click);
            Export_Manage_Tab.Controls.Add(Export_Search_BT);

            possize.pos_x = 500;
            possize.pos_y = 6;
            Export_Search_Txt_Lb = new TextBox_Lbl(Export_Manage_Tab, "              Search", TextBox_Type.TEXT, possize, AnchorType.RIGHT);

            return true;
        }
    }
}