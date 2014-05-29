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
        private Button Stock_Create_BT;
        public Button Stock_Store_BT;
        public Button Stock_Process_BT;
        public Button Stock_Search_BT;
        private TextBox_Lbl Stock_Search_Txt_Lb;

        public DataTable Load_Stock_TBL;
        //DataSet Load_Stock_ds = new DataSet();
        //SqlDataAdapter Load_Stock_da;

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
            possize.pos_y = 100;
            possize.width = Stock_Manage_Tab.Size.Width - 15;
            possize.height = Stock_Manage_Tab.Size.Height - 100;
            Stock_Table_Form = new Gridview_Grp(Stock_Manage_Tab, "Stock Manage Table", possize, AUTO_RESIZE,
                                                Database_WHM_Stock_Con_Str, @"SELECT * FROM dbo.Material_Stock_tb", AnchorType.LEFT);
            Stock_Table_Form.Load_DataBase(Database_WHM_Stock_Con_Str, @"SELECT * FROM dbo.Material_Stock_tb");
            Stock_Init_BT();
        }

        public bool Stock_Init_BT()
        {
            PosSize possize = new PosSize();
            Stock_Create_BT = new Button();
            Stock_Store_BT = new Button();
            Stock_Process_BT = new Button();
            Stock_Search_BT = new Button();


            Stock_Create_BT.Name = "Stock_Create_BT";
            Stock_Create_BT.Text = "Create";
            Stock_Create_BT.Location = new System.Drawing.Point(55, 69);
            Stock_Create_BT.Size = new System.Drawing.Size(50, 23);
            Stock_Create_BT.UseVisualStyleBackColor = true;
            Stock_Create_BT.Click += new System.EventHandler(Stock_Create_BT_Click);
            Stock_Manage_Tab.Controls.Add(Stock_Create_BT);

            Stock_Store_BT.Name = "Stock_Store_BT";
            Stock_Store_BT.Text = "Save";
            Stock_Store_BT.Location = new System.Drawing.Point(260, 69);
            Stock_Store_BT.Size = new System.Drawing.Size(50, 23);
            Stock_Store_BT.UseVisualStyleBackColor = true;
            Stock_Store_BT.Click += new System.EventHandler(Stock_Store_BT_Click);
            Stock_Manage_Tab.Controls.Add(Stock_Store_BT);

            Stock_Process_BT.Name = "Stock_Process_BT";
            Stock_Process_BT.Text = "Process";
            Stock_Process_BT.Location = new System.Drawing.Point(150, 69);
            Stock_Process_BT.Size = new System.Drawing.Size(60, 23);
            Stock_Process_BT.UseVisualStyleBackColor = true;
            Stock_Process_BT.Click += new System.EventHandler(Stock_Process_BT_Click);
            Stock_Manage_Tab.Controls.Add(Stock_Process_BT);

            Stock_Search_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            Stock_Search_BT.Name = "Stock_Search_BT";
            Stock_Search_BT.Text = "Search";
            Stock_Search_BT.Location = new System.Drawing.Point(600, 35);
            Stock_Search_BT.Size = new System.Drawing.Size(60, 23);
            Stock_Search_BT.UseVisualStyleBackColor = true;
            Stock_Search_BT.Click += new System.EventHandler(Stock_Search_BT_Click);
            Stock_Manage_Tab.Controls.Add(Stock_Search_BT);

            possize.pos_x = 500;
            possize.pos_y = 6;
            Stock_Search_Txt_Lb = new TextBox_Lbl(Stock_Manage_Tab, "              Search", TextBox_Type.TEXT, possize, AnchorType.RIGHT);

            return true;
        }
    }
}