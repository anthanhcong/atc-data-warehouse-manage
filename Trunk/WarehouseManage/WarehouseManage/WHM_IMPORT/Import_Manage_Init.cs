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
        //            Import Manage            //
        /*************************************/


        private System.Windows.Forms.TabPage Import_Manage_Tab;
        private Gridview_Grp Import_Table_Form;
        private Button Import_Create_BT;
        public Button Import_Store_BT;
        public Button Import_Process_BT;
        public Button Import_Search_BT;
        private TextBox_Lbl Import_Search_Txt_Lb;

        public DataTable Load_Import_TBL;
        //DataSet Load_Import_ds = new DataSet();
        //SqlDataAdapter Load_Import_da;

        private void Create_Import_Manage_Tab()
        {
            PosSize possize = new PosSize();
            string tab_name = "Import Manage Tab";

            Import_Manage_Tab = new System.Windows.Forms.TabPage();
            Import_Manage_Tab.Text = tab_name;
            Import_Manage_Tab.SuspendLayout();
            Import_Manage_Tab.Location = new System.Drawing.Point(4, 22);
            Import_Manage_Tab.Size = new System.Drawing.Size(900, 390);
            Import_Manage_Tab.Padding = new System.Windows.Forms.Padding(3);
            //Import_Manage_Tab.TabIndex = 1;
            Import_Manage_Tab.UseVisualStyleBackColor = true;
            Import_Manage_Tab.ResumeLayout(true);
            Import_Manage_Tab.PerformLayout();
            this.MainTabControl.Controls.Add(this.Import_Manage_Tab);

            // Init Card Table
            possize.pos_x = 6;
            possize.pos_y = 100;
            possize.width = Import_Manage_Tab.Size.Width - 15;
            possize.height = Import_Manage_Tab.Size.Height - 100;
            Import_Table_Form = new Gridview_Grp(Import_Manage_Tab, "Import Manage Table", possize, AUTO_RESIZE,
                                                Database_WHM_Import_Con_Str, @"SELECT * FROM dbo.Production_Import_tb", AnchorType.LEFT);
            Import_Table_Form.Load_DataBase(Database_WHM_Import_Con_Str, @"SELECT * FROM dbo.Production_Import_tb");
            Import_Init_BT();
        }

        public bool Import_Init_BT()
        {
            PosSize possize = new PosSize();
            Import_Create_BT = new Button();
            Import_Store_BT = new Button();
            Import_Process_BT = new Button();
            Import_Search_BT = new Button();


            Import_Create_BT.Name = "Import_Create_BT";
            Import_Create_BT.Text = "Create";
            Import_Create_BT.Location = new System.Drawing.Point(55, 69);
            Import_Create_BT.Size = new System.Drawing.Size(50, 23);
            Import_Create_BT.UseVisualStyleBackColor = true;
            Import_Create_BT.Click += new System.EventHandler(Import_Create_BT_Click);
            Import_Manage_Tab.Controls.Add(Import_Create_BT);

            Import_Store_BT.Name = "Import_Store_BT";
            Import_Store_BT.Text = "Save";
            Import_Store_BT.Location = new System.Drawing.Point(260, 69);
            Import_Store_BT.Size = new System.Drawing.Size(50, 23);
            Import_Store_BT.UseVisualStyleBackColor = true;
            Import_Store_BT.Click += new System.EventHandler(Import_Store_BT_Click);
            Import_Manage_Tab.Controls.Add(Import_Store_BT);

            Import_Process_BT.Name = "Import_Process_BT";
            Import_Process_BT.Text = "Process";
            Import_Process_BT.Location = new System.Drawing.Point(150, 69);
            Import_Process_BT.Size = new System.Drawing.Size(60, 23);
            Import_Process_BT.UseVisualStyleBackColor = true;
            Import_Process_BT.Click += new System.EventHandler(Import_Process_BT_Click);
            Import_Manage_Tab.Controls.Add(Import_Process_BT);

            Import_Search_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            Import_Search_BT.Name = "Import_Search_BT";
            Import_Search_BT.Text = "Search";
            Import_Search_BT.Location = new System.Drawing.Point(600, 35);
            Import_Search_BT.Size = new System.Drawing.Size(60, 23);
            Import_Search_BT.UseVisualStyleBackColor = true;
            Import_Search_BT.Click += new System.EventHandler(Import_Search_BT_Click);
            Import_Manage_Tab.Controls.Add(Import_Search_BT);

            possize.pos_x = 500;
            possize.pos_y = 6;
            Import_Search_Txt_Lb = new TextBox_Lbl(Import_Manage_Tab, "              Search", TextBox_Type.TEXT, possize, AnchorType.RIGHT);

            return true;
        }
    }
}