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
        /*************************************************
                           #  ##    #  # #######
                           #  # #   #  #    #   
                           #  #  #  #  #    #   
                           #  #   # #  #    #   
                           #  #    ##  #    #   
                ************************************************/


        /**********************************************/
        //               User Manage                  //
        /**********************************************/

        private System.Windows.Forms.TabPage User_Manage_Tab;
        private Gridview_Grp User_Table_Form;
        private Gridview_Grp Permission_Table_Form;
        private Button User_Manage_Import_BT;
        public Button User_Manage_Store_BT;
        public Button User_Manage_Insert_BOM_BT;
        public Button User_Manage_Search_Component_BT;
        public Button User_Manage_Add_Part_number_BT;
        private TextBox_Lbl_F2 User_Manage_Search_Component;
        private TextBox_Lbl_F2 User_Manage_UserName;
        private TextBox_Lbl_F2 User_Manage_Password;
        private TextBox_Lbl_F2 User_Manage_MSNV;
        private TextBox_Lbl_F2 User_Manage_AliasName;
        private TextBox_Lbl_F2 User_Manage_PermissionID;
        private TextBox_Lbl_F2 User_Manage_BomManage;
        private TextBox_Lbl_F2 User_Manage_ImportMaterial;
        //private TextBox_Lbl_F2 User_Manage_;
        //private TextBox_Lbl_F2 User_Manage_;
        //private TextBox_Lbl_F2 User_Manage_;
        private GroupBox User_Manage_Add_Part_gbx;

        //public event EventHandler SelectedIndexChanged;

        private void Create_User_Manage_Tab()
        {
            PosSize possize = new PosSize();
            string tab_name = "User_Manage_Tab";

            User_Manage_Tab = new System.Windows.Forms.TabPage();
            User_Manage_Tab.Text = tab_name;
            User_Manage_Tab.SuspendLayout();
            User_Manage_Tab.Location = new System.Drawing.Point(4, 22);
            User_Manage_Tab.Size = new System.Drawing.Size(900, 390);
            User_Manage_Tab.Padding = new System.Windows.Forms.Padding(3);
            //User_Manage_Tab.TabIndex = 1;
            User_Manage_Tab.UseVisualStyleBackColor = true;
            User_Manage_Tab.ResumeLayout(true);
            User_Manage_Tab.PerformLayout();
            this.MainTabControl.Controls.Add(this.User_Manage_Tab);

            // Init Card Table
            possize.pos_x = 6;
            possize.pos_y = 160;
            possize.width = User_Manage_Tab.Size.Width - 400; // 200;
            possize.height = User_Manage_Tab.Size.Height - 160;
            User_Table_Form = new Gridview_Grp(User_Manage_Tab, "User Table", possize, NO_AUTO_RESIZE,
                                                User_Conn, @"SELECT * FROM dbo.USER_DASHBOARD_tb", AnchorType.LEFT);
            User_Table_Form.Load_DataBase(User_Conn, @"SELECT * FROM dbo.USER_DASHBOARD_tb");
            User_Table_Form.dataGridView_View.CellDoubleClick += new DataGridViewCellEventHandler(User_Table_Form_CellDoubleClick);
            possize.pos_x = 380;
            possize.pos_y = 160;
            possize.width = User_Manage_Tab.Size.Width - 200;
            possize.height = User_Manage_Tab.Size.Height - 160;
            Permission_Table_Form = new Gridview_Grp(User_Manage_Tab, "Permission Table", possize, NO_AUTO_RESIZE,
                                                User_Conn, @"SELECT * FROM dbo.PERMISSION_tb", AnchorType.LEFT);

            User_Manage_Init_BT();
            User_Manage_Init_Search();
            User_Manage_Control_Group_Init();
        }

        void dataGridView_View_SelectionChanged(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private bool User_Manage_Control_Group_Init()
        {
            PosSize possize = new PosSize();
            User_Manage_Add_Part_gbx = new GroupBox();
            User_Manage_Add_Part_number_BT = new Button();

            User_Manage_Tab.Controls.Add(User_Manage_Add_Part_gbx);
            User_Manage_Add_Part_gbx.Location = new System.Drawing.Point(6, 6);
            User_Manage_Add_Part_gbx.Name = "User_Manage_Add_Part_gbx";
            User_Manage_Add_Part_gbx.Size = new System.Drawing.Size(580, 110);
            User_Manage_Add_Part_gbx.TabIndex = 1;
            User_Manage_Add_Part_gbx.TabStop = false;
            User_Manage_Add_Part_gbx.Text = "Add Part";

            possize.pos_x = 10;
            possize.pos_y = 6;
            User_Manage_UserName = new TextBox_Lbl_F2(User_Manage_Tab, "UserName", 20, 100, TextBox_Type.TEXT, possize, AnchorType.LEFT);

            possize.pos_x = 120;
            possize.pos_y = 6;
            User_Manage_Password = new TextBox_Lbl_F2(User_Manage_Tab, "Password", 20, 70, TextBox_Type.TEXT, possize, AnchorType.LEFT);


            possize.pos_x = 120;
            possize.pos_y = 55;
            User_Manage_MSNV = new TextBox_Lbl_F2(User_Manage_Tab, "MSNV", 20, 70, TextBox_Type.TEXT, possize, AnchorType.LEFT);

            possize.pos_x = 10;
            possize.pos_y = 55;
            User_Manage_AliasName = new TextBox_Lbl_F2(User_Manage_Tab, "AliasName", 20, 100, TextBox_Type.TEXT, possize, AnchorType.LEFT);

            possize.pos_x = 200;
            possize.pos_y = 6;
            User_Manage_PermissionID = new TextBox_Lbl_F2(User_Manage_Tab, "PermissionID", 20, 70, TextBox_Type.TEXT, possize, AnchorType.LEFT);

            possize.pos_x = 290;
            possize.pos_y = 6;
            User_Manage_BomManage = new TextBox_Lbl_F2(User_Manage_Tab, "Bom Manage( true or false )", 20, 70, TextBox_Type.TEXT, possize, AnchorType.LEFT);

            possize.pos_x = 290;
            possize.pos_y = 55;
            User_Manage_ImportMaterial = new TextBox_Lbl_F2(User_Manage_Tab, "Import Material( true or false )", 20, 70, TextBox_Type.TEXT, possize, AnchorType.LEFT);


            User_Manage_Add_Part_number_BT.Name = "User_Manage_Add_Part_number_BT";
            User_Manage_Add_Part_number_BT.Text = "Add \n Part ";
            User_Manage_Add_Part_number_BT.Location = new System.Drawing.Point(515, 30);
            User_Manage_Add_Part_number_BT.Size = new System.Drawing.Size(50, 40);
            User_Manage_Add_Part_number_BT.UseVisualStyleBackColor = true;
            User_Manage_Add_Part_number_BT.Click += new System.EventHandler(User_Manage_Add_Part_number_BT_Click);
            User_Manage_Add_Part_gbx.Controls.Add(User_Manage_Add_Part_number_BT);

            User_Manage_Add_Part_gbx.Controls.Add(User_Manage_UserName.My_Label);
            User_Manage_Add_Part_gbx.Controls.Add(User_Manage_UserName.My_TextBox);
            User_Manage_Add_Part_gbx.Controls.Add(User_Manage_Password.My_Label);
            User_Manage_Add_Part_gbx.Controls.Add(User_Manage_Password.My_TextBox);
            User_Manage_Add_Part_gbx.Controls.Add(User_Manage_MSNV.My_Label);
            User_Manage_Add_Part_gbx.Controls.Add(User_Manage_MSNV.My_TextBox);
            User_Manage_Add_Part_gbx.Controls.Add(User_Manage_AliasName.My_Label);
            User_Manage_Add_Part_gbx.Controls.Add(User_Manage_AliasName.My_TextBox);
            User_Manage_Add_Part_gbx.Controls.Add(User_Manage_PermissionID.My_Label);
            User_Manage_Add_Part_gbx.Controls.Add(User_Manage_PermissionID.My_TextBox);
            User_Manage_Add_Part_gbx.Controls.Add(User_Manage_BomManage.My_Label);
            User_Manage_Add_Part_gbx.Controls.Add(User_Manage_BomManage.My_TextBox);
            User_Manage_Add_Part_gbx.Controls.Add(User_Manage_ImportMaterial.My_Label);
            User_Manage_Add_Part_gbx.Controls.Add(User_Manage_ImportMaterial.My_TextBox);

            return true;
        }

        public bool User_Manage_Init_Search()
        {
            PosSize possize = new PosSize();
            User_Manage_Search_Component_BT = new Button();

            User_Manage_Search_Component_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            User_Manage_Search_Component_BT.Name = "User_Manage_Search_Component_BT";
            User_Manage_Search_Component_BT.Text = "Search User";
            User_Manage_Search_Component_BT.Location = new System.Drawing.Point(630, 65);
            User_Manage_Search_Component_BT.Size = new System.Drawing.Size(70, 38);
            User_Manage_Search_Component_BT.UseVisualStyleBackColor = true;
            User_Manage_Search_Component_BT.Click += new System.EventHandler(User_Manage_Search_Component_BT_Click);
            User_Manage_Tab.Controls.Add(User_Manage_Search_Component_BT);

            possize.pos_x = 630;
            possize.pos_y = 6;
            User_Manage_Search_Component = new TextBox_Lbl_F2(User_Manage_Tab, "UserName", 20, 70, TextBox_Type.TEXT, possize, AnchorType.RIGHT);

            return true;
        }

        public bool User_Manage_Init_BT()
        {

            User_Manage_Import_BT = new Button();
            User_Manage_Store_BT = new Button();
            User_Manage_Insert_BOM_BT = new Button();

            User_Manage_Import_BT.Name = "User_Manage_Import_BT";
            User_Manage_Import_BT.Text = "Import";
            User_Manage_Import_BT.Location = new System.Drawing.Point(20, 130);
            User_Manage_Import_BT.Size = new System.Drawing.Size(50, 23);
            User_Manage_Import_BT.UseVisualStyleBackColor = true;
            User_Manage_Import_BT.Click += new System.EventHandler(User_Manage_Import_BT_Click_event);
            User_Manage_Tab.Controls.Add(User_Manage_Import_BT);

            User_Manage_Store_BT.Name = "User_Manage_Store_BT";
            User_Manage_Store_BT.Text = "Save";
            User_Manage_Store_BT.Location = new System.Drawing.Point(230, 130);
            User_Manage_Store_BT.Size = new System.Drawing.Size(50, 23);
            User_Manage_Store_BT.UseVisualStyleBackColor = true;
            User_Manage_Store_BT.Click += new System.EventHandler(User_Manage_Store_BT_Click);
            User_Manage_Tab.Controls.Add(User_Manage_Store_BT);


            //User_Manage_Insert_BOM_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            User_Manage_Insert_BOM_BT.Name = "User_Manage_Insert_BOM_BT";
            User_Manage_Insert_BOM_BT.Text = "Insert BOM";
            User_Manage_Insert_BOM_BT.Location = new System.Drawing.Point(115, 130);
            User_Manage_Insert_BOM_BT.Size = new System.Drawing.Size(70, 23);
            User_Manage_Insert_BOM_BT.UseVisualStyleBackColor = true;
            //User_Manage_Insert_BOM_BT.Click += new System.EventHandler(User_Manage_Insert_BOM_BT_Click);
            User_Manage_Tab.Controls.Add(User_Manage_Insert_BOM_BT);
            User_Manage_Insert_BOM_BT.Visible = false;

            return true;
        }

    }
}