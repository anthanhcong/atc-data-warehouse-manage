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
    class Gridview_Grp:SQL_APPL
    {
        public System.Windows.Forms.GroupBox Tab_Grp;
        public System.Windows.Forms.DataGridView dataGridView_View;
        public System.Windows.Forms.Button Review_BT;
        private System.Windows.Forms.Button Submit_BT;
        public System.Windows.Forms.Button Export_BT;
        public System.Windows.Forms.Button Delete_All_BT;
        public System.Windows.Forms.Button Delete_Rows_BT;


        private string Database_Conn;
        private string SQL_Load_CMD;
        public DataTable Data_dtb = new DataTable();
        public DataSet Data_ds = new DataSet();
        public SqlDataAdapter Data_da;
        private string Group_Name;
        PosSize My_PosSize;
        bool My_autoResize;
        AnchorType My_anchor;


        public Gridview_Grp(System.Windows.Forms.TabPage owner_tab, string group_name,PosSize possize, 
                            bool autoresize, string connection_str, string sql_load_cmd, AnchorType anchor)
        {
            Database_Conn = connection_str;
            SQL_Load_CMD = sql_load_cmd;
            My_PosSize = possize;
            My_autoResize = autoresize;
            My_anchor = anchor;
            Group_Name = group_name;
            Init_GrpBox(owner_tab, group_name);
            Init_GridView(owner_tab, group_name);
            //Load_DataBase(Database_Conn, sql_load_cmd);
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
                //this.Tab_Grp.AutoSize = true;
            }else {
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
            Export_BT = new Button();
            Delete_All_BT = new Button();
            Delete_Rows_BT = new Button();
            
            Tab_Grp.Controls.Add(dataGridView_View);
            Tab_Grp.Controls.Add(Review_BT);
            //Tab_Grp.Controls.Add(Submit_BT);
            Tab_Grp.Controls.Add(Export_BT);
            Tab_Grp.Controls.Add(Delete_All_BT);
            Tab_Grp.Controls.Add(Delete_Rows_BT);

            dataGridView_View.Location = new System.Drawing.Point(10, 16);
            dataGridView_View.Size = new System.Drawing.Size(My_PosSize.width - 20, My_PosSize.height - 50);
            dataGridView_View.Anchor = ((System.Windows.Forms.AnchorStyles)(
                                System.Windows.Forms.AnchorStyles.Top
                                | System.Windows.Forms.AnchorStyles.Bottom
                                | System.Windows.Forms.AnchorStyles.Left
                                | System.Windows.Forms.AnchorStyles.Right));
            dataGridView_View.ScrollBars = ScrollBars.Both;
            dataGridView_View.AllowUserToDeleteRows = false;
            //dataGridView_View.CellContentDoubleClick += new DataGridViewCellEventHandler(dataGridView_View_CellContentDoubleClick);
            //dataGridView_View.RowsDefaultCellStyle.BackColor = Color.White;
            //dataGridView_View.AlternatingRowsDefaultCellStyle.BackColor = Color.LightSkyBlue;
            //dataGridView_View.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //dataGridView_View.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            //dataGridView_View.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;


            Review_BT.Text = "Refresh";
            Review_BT.Location = new System.Drawing.Point(235, Tab_Grp.Size.Height - 30);
            Review_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((
                                System.Windows.Forms.AnchorStyles.Bottom)
                                | System.Windows.Forms.AnchorStyles.Left));
            Review_BT.Size = new System.Drawing.Size(60, 23);
            Review_BT.Click += new System.EventHandler(Review_BT_Click_event);
            //Review_BT.Visible = false;

            Delete_Rows_BT.Text = "Del Rows";
            Delete_Rows_BT.Location = new System.Drawing.Point(10, Tab_Grp.Size.Height - 30);
            Delete_Rows_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((
                                System.Windows.Forms.AnchorStyles.Bottom)
                                | System.Windows.Forms.AnchorStyles.Left));
            Delete_Rows_BT.Size = new System.Drawing.Size(65, 23);
            Delete_Rows_BT.Click += new System.EventHandler(Delete_Rows_BT_Click_event);


            Submit_BT.Text = "Submit";
            Submit_BT.Location = new System.Drawing.Point(300, Tab_Grp.Size.Height - 30);
            Submit_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((
                                System.Windows.Forms.AnchorStyles.Bottom)
                                | System.Windows.Forms.AnchorStyles.Left));
            Submit_BT.Click += new System.EventHandler(Submit_BT_Click_event);
            Submit_BT.Visible = false;

            Delete_All_BT.Text = "Del Data";
            Delete_All_BT.Location = new System.Drawing.Point(90, Tab_Grp.Size.Height - 30);
            Delete_All_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((
                                System.Windows.Forms.AnchorStyles.Bottom)
                                | System.Windows.Forms.AnchorStyles.Left));
            Delete_All_BT.Size = new System.Drawing.Size(60, 23);
            Delete_All_BT.Click += new System.EventHandler(Delete_All_BT_Click_event);
            //Delete_All_BT.Visible = false;

            Export_BT.Text = "Export";
            Export_BT.Location = new System.Drawing.Point(165, Tab_Grp.Size.Height - 30);
            Export_BT.Size = new System.Drawing.Size(50, 23);
            Export_BT.Anchor = ((System.Windows.Forms.AnchorStyles)((
                                System.Windows.Forms.AnchorStyles.Bottom)
                                | System.Windows.Forms.AnchorStyles.Left));
            Export_BT.Click += new System.EventHandler(Export_BT_Click_event);
            //Export_BT.Visible = false;

            
            return true;
        }

        public bool Load_DataBase(string connection_str, string sql_cmd)
        {
            SQL_Load_CMD = sql_cmd;
            if (Data_dtb != null)
            {
                Data_dtb.Clear();
            }
            Data_dtb = Get_SQL_Data(connection_str, sql_cmd, ref Data_da, ref Data_ds);
            dataGridView_View.DataSource = Data_dtb;
            if (Data_dtb == null)
            {
                return false;
            }
            else
            {
                return true;
            }
            // dataGridView_View.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
        }

        public void Update_Size(PosSize possize)
        {
            My_PosSize = possize;
            dataGridView_View.Size = new System.Drawing.Size(My_PosSize.width - 20, My_PosSize.height - 50);
        }

        public void Review_BT_Click_event(object sender, EventArgs e)
        {
            Load_DataBase(Database_Conn, SQL_Load_CMD);
        }

        public void Delete_Rows_BT_Click_event(object sender, EventArgs e)
        {
            int max_row;

            max_row = dataGridView_View.RowCount;

            for (int i = 0; i < max_row - 1; i++)
            {
                if (dataGridView_View.Rows[i].Selected)
                {
                    if (MessageBox.Show("Would you like to Delete row: " + (i + 1).ToString() + "?", "Attention", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        dataGridView_View.Rows.RemoveAt(i);
                        max_row--;
                        i--;
                    }
                }
            }
        }

        public void Submit_BT_Click_event(object sender, EventArgs e)
        {
            if ((Update_SQL_Data(Data_da, Data_dtb) == true))
            {
                MessageBox.Show("Store Data Complete", "Successful");
            }
            else
            {
                MessageBox.Show("Store Data Fail", "Failed");
            }
        }


        public void Delete_All_BT_Click_event(object sender, EventArgs e)
        {  
            if (MessageBox.Show("Would you like to Delete All Data " + "?", "Attention", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                int total_row;
                total_row = dataGridView_View.RowCount;
                for (int i = 0; i < total_row -1; i++)
                {
                    dataGridView_View.Rows.RemoveAt(i);
                    total_row--;
                    i--;
                }
            }
        }

        public void dataGridView_View_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (MessageBox.Show("Do you want to copy clipboard and edit " + "?", "Attention", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                if (dataGridView_View.CurrentCell.Value.ToString().Trim() != null)
                {
                    Clipboard.SetDataObject(dataGridView_View.CurrentCell.Value.ToString().Trim(), false);
                    //dataGridView_View.ClearSelection();
                }
            }
        }

        public void Export_BT_Click_event(object sender, EventArgs e)
        {
            SaveFileDialog save_diaglog = new SaveFileDialog();
            string file_name, fInfo;
            string temp;

            if (Update_SQL_Data(Data_da, Data_dtb) == true)
            {
                save_diaglog.Filter = "Excel file (*.xls)|*.xls|All files (*.*)|*.*"; ;
                if (save_diaglog.ShowDialog() == DialogResult.OK)
                {
                    file_name = save_diaglog.FileName;
                    fInfo = Path.GetExtension(save_diaglog.FileName);
                    temp = Export_BT.Text;
                    Export_BT.Text = "Exporting ...";
                    Export_BT.Enabled = false;
                    if ((fInfo == ".xlsx") || (fInfo == ".xls"))
                    {
                        ExportDataToExcel(file_name, Group_Name, Data_dtb);
                    }
                    Export_BT.Enabled = true;
                    Export_BT.Text = temp;
                    MessageBox.Show("Export File thành công", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show("Cập nhật thay đổi trước khi export file thất bại", "Thông báo");
            }
            
        }
        public void Refresh_Form()
        {
            Load_DataBase(Database_Conn, SQL_Load_CMD);
        }
    }
}
