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

namespace WarehouseManager
{
    class Gridview_Timer_Group : SQL_APPL
    {
        public System.Windows.Forms.GroupBox Tab_Grp;
        private System.Windows.Forms.Label Date_to_Lbl;
        private System.Windows.Forms.Label Date_from_Lbl;
        private System.Windows.Forms.DateTimePicker dateTimePicker_To;
        private System.Windows.Forms.DateTimePicker dateTimePicker_From;
        public System.Windows.Forms.DataGridView dataGridView_View;
        private System.Windows.Forms.Button Review_BT;
        private System.Windows.Forms.Button Submit_BT;
        //private System.Windows.Forms.Button Export_BT;

        private string Database_Conn;
        private string SQL_Load_CMD;
        public DataTable Data_dtb = new DataTable();
        DataSet Data_ds = new DataSet();
        SqlDataAdapter Data_da;
        PosSize My_PosSize;
        bool My_autoResize;
        AnchorType My_anchor;

        public Gridview_Timer_Group(System.Windows.Forms.TabPage owner_tab, string group_name, PosSize possize,
                            bool autoresize, string connection_str, string sql_load_cmd, AnchorType anchor)
        {
            Database_Conn = connection_str;
            SQL_Load_CMD = sql_load_cmd;
            My_PosSize = possize;
            My_autoResize = autoresize;
            My_anchor = anchor;
            Init_GrpBox(owner_tab, group_name);
            Init_GridView(owner_tab, group_name);
            Init_DatePicker(owner_tab, group_name);
            Load_DataBase(Database_Conn);
        }

        private bool Init_GrpBox(System.Windows.Forms.TabPage owner_tab, string group_name)
        {
            int height, width;
            height = My_PosSize.height;
            width = My_PosSize.width;

            Tab_Grp = new System.Windows.Forms.GroupBox();
            owner_tab.Controls.Add(Tab_Grp);
            this.Tab_Grp.AutoSize = true;
            this.Tab_Grp.SuspendLayout();
            this.Tab_Grp.Location = new System.Drawing.Point(My_PosSize.pos_x, My_PosSize.pos_y);
            this.Tab_Grp.Name = group_name;
            this.Tab_Grp.Size = new System.Drawing.Size(width, height);
            this.Tab_Grp.TabIndex = 0;
            this.Tab_Grp.TabStop = false;
            this.Tab_Grp.Text = group_name;
            this.Tab_Grp.ResumeLayout(true);
            this.Tab_Grp.PerformLayout();
            this.Tab_Grp.AutoSize = false;
            if (My_autoResize == true)
            {
                this.Tab_Grp.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                                        | System.Windows.Forms.AnchorStyles.Bottom
                                        | System.Windows.Forms.AnchorStyles.Left
                                        | System.Windows.Forms.AnchorStyles.Right));
            }
            else
            {
                if (My_anchor == AnchorType.RIGHT)
                {
                    this.Tab_Grp.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                                                            | System.Windows.Forms.AnchorStyles.Right));
                }
                else if (My_anchor == AnchorType.LEFT)
                {
                    this.Tab_Grp.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                                                            | System.Windows.Forms.AnchorStyles.Left));
                }
            }
            return true;
        }

        private bool Init_GridView(System.Windows.Forms.TabPage owner_tab, string group_name)
        {
            dataGridView_View = new DataGridView();
            Review_BT = new Button();
            Submit_BT = new Button();

            Tab_Grp.Controls.Add(dataGridView_View);
            Tab_Grp.Controls.Add(Review_BT);
            Tab_Grp.Controls.Add(Submit_BT);

            dataGridView_View.Location = new System.Drawing.Point(10, 16);
            dataGridView_View.Size = new System.Drawing.Size(My_PosSize.width - 20, My_PosSize.height - 50);
            dataGridView_View.Anchor = ((System.Windows.Forms.AnchorStyles)(
                                System.Windows.Forms.AnchorStyles.Top
                                | System.Windows.Forms.AnchorStyles.Bottom
                                | System.Windows.Forms.AnchorStyles.Left));
            dataGridView_View.ScrollBars = ScrollBars.Both;


            Review_BT.Text = "Review";
            Review_BT.Location = new System.Drawing.Point(10, Tab_Grp.Size.Height - 30);
            Review_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((
                                System.Windows.Forms.AnchorStyles.Bottom)
                                | System.Windows.Forms.AnchorStyles.Left));
            Review_BT.Click += new System.EventHandler(Review_BT_Click_event);

            Submit_BT.Text = "Submit";
            Submit_BT.Location = new System.Drawing.Point(110, Tab_Grp.Size.Height - 30);
            Submit_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((
                                System.Windows.Forms.AnchorStyles.Bottom)
                                | System.Windows.Forms.AnchorStyles.Left));
            Submit_BT.Click += new System.EventHandler(Submit_BT_Click_event);


            return true;
        }

        private bool Init_DatePicker(System.Windows.Forms.TabPage owner_tab, string tab_name)
        {
            dateTimePicker_To = new DateTimePicker();
            dateTimePicker_From = new DateTimePicker();
            Date_to_Lbl = new Label();
            Date_from_Lbl = new Label();
            dataGridView_View = new DataGridView();
            Review_BT = new Button();
            Submit_BT = new Button();

            Tab_Grp.Controls.Add(dataGridView_View);
            Tab_Grp.Controls.Add(Review_BT);
            Tab_Grp.Controls.Add(Submit_BT);
            
            Tab_Grp.Controls.Add(dateTimePicker_To);
            Tab_Grp.Controls.Add(dateTimePicker_From);
            Tab_Grp.Controls.Add(Date_to_Lbl);
            Tab_Grp.Controls.Add(Date_from_Lbl);
            Tab_Grp.Controls.Add(dataGridView_View);
            Tab_Grp.Controls.Add(Review_BT);
            Tab_Grp.Controls.Add(Submit_BT);
            //owner_tab.Controls.Add(Review_BT);
                        
            Date_from_Lbl.Text = "From";
            Date_from_Lbl.Location = new System.Drawing.Point(10, 20);
            dateTimePicker_From.Location = new System.Drawing.Point(50, 15);

            Date_to_Lbl.Text = "To";
            Date_to_Lbl.Location = new System.Drawing.Point(255, 20);
            dateTimePicker_To.Location = new System.Drawing.Point(300, 15);

            dataGridView_View.Location = new System.Drawing.Point(10, 50);
            dataGridView_View.Size = new System.Drawing.Size(My_PosSize.width - 20, My_PosSize.height - 80);
            dataGridView_View.Anchor = ((System.Windows.Forms.AnchorStyles)(
                                System.Windows.Forms.AnchorStyles.Top
                                | System.Windows.Forms.AnchorStyles.Bottom
                                | System.Windows.Forms.AnchorStyles.Left));
            dataGridView_View.ScrollBars = ScrollBars.Both;

            Review_BT.Text = "Review";
            Review_BT.Location = new System.Drawing.Point(10, Tab_Grp.Size.Height - 30);
            Review_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((
                                System.Windows.Forms.AnchorStyles.Bottom)
                                | System.Windows.Forms.AnchorStyles.Left));
            Review_BT.Click += new System.EventHandler(Review_BT_Click_event);
            Submit_BT.Text = "Submit";
            Submit_BT.Location = new System.Drawing.Point(110, Tab_Grp.Size.Height - 30);
            Submit_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((
                                System.Windows.Forms.AnchorStyles.Bottom)
                                | System.Windows.Forms.AnchorStyles.Left));
            Submit_BT.Click += new System.EventHandler(Submit_BT_Click_event);
            return true;
        }

        private bool Load_DataBase(string connection_str)
        {
            string sql_cmd = SQL_Load_CMD;
            sql_cmd = sql_cmd + " WHERE [Date] BETWEEN '" + dateTimePicker_From.Value.Date.ToString("MM/dd/yyyy")
                      + "' AND '" + dateTimePicker_To.Value.Date.ToString("MM/dd/yyyy") + "'";

            Data_dtb.Clear();
            Data_dtb = Get_SQL_Data(connection_str, sql_cmd, ref Data_da, ref Data_ds);
            dataGridView_View.DataSource = Data_dtb;

            return true;
        }

        public void Update_Size(PosSize possize)
        {
            My_PosSize = possize;
            dataGridView_View.Size = new System.Drawing.Size(My_PosSize.width - 20, My_PosSize.height - 80);
        }

        public void Review_BT_Click_event(object sender, EventArgs e)
        {
            Load_DataBase(Database_Conn);
        }

        public void Submit_BT_Click_event(object sender, EventArgs e)
        {
            if (Update_SQL_Data(Data_da, Data_dtb) == true)
            {
                MessageBox.Show("Store Data Complete", "Successful");
            }
            else
            {
                MessageBox.Show("Store Data Fail", "Failed");
            }
        }
        public void Refresh_Form()
        {
            Load_DataBase(Database_Conn);
        }
    }
}
