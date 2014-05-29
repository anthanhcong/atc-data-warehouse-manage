using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.IO;
using System.IO.Ports;
using System.Globalization;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace WarehouseManager
{
    public partial class Form1 : SQL_APPL
    {

        private Point _imageLocation = new Point(15, 5);
        private Point _imgHitArea = new Point(13, 2);

        string Error_log;

        public string User_Conn = @"server=ADMIN\SQLEXPRESS;database=USER_DB;uid=sa;pwd=123456";
        public string Database_WHM_Info_Con_Str = @"server=ADMIN\SQLEXPRESS;database=WHM_INFOMATION_DB;uid=sa;pwd=123456";
        public string Database_WHM_Stock_Con_Str = @"server=ADMIN\SQLEXPRESS;database=WHM_STOCK_DB;uid=sa;pwd=123456";
        public string Database_WHM_Import_Con_Str = @"server=ADMIN\SQLEXPRESS;database=WHM_IMPORT_DB;uid=sa;pwd=123456";
        public string Database_WHM_Export_Con_Str = @"server=ADMIN\SQLEXPRESS;database=WHM_EXPORT_DB;uid=sa;pwd=123456";

        private const bool AUTO_RESIZE = true;
        private const bool NO_AUTO_RESIZE = false;
        public int RELOAD_DB = 0;

        public DataTable List_Item_FA_TBL;

        public Form1()
        {
            InitializeComponent();
            OpenXL = new Excel.Application();
            OpenXL.SheetsInNewWorkbook = 1;
            OpenXL.Visible = false;
            OpenXL.DisplayAlerts = false;

            INPUT_SXXK_NK_InitExcelCol_Infor();
            INPUT_SXXK_XK_InitExcelCol_Infor();
            INPUT_KD_NK_InitExcelCol_Infor();
            INPUT_KD_XK_InitExcelCol_Infor();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CultureInfo ci = new CultureInfo(Application.CurrentCulture.Name, true);
            DateTimeFormatInfo dfi = new DateTimeFormatInfo();
            dfi.ShortDatePattern = "dd-MMM-yyyy";
            ci.DateTimeFormat = dfi;
            Application.CurrentCulture = ci;

            Create_INPUT_NK_Manage_Tab();
            Create_INPUT_XK_Manage_Tab();
            Create_Import_Manage_Tab();
            Create_Export_Manage_Tab();
            Create_Stock_Manage_Tab();
            //Create_User_Manage_Tab();

        }

        private void Form1_Closed(object sender, FormClosedEventArgs e)
        {
            OpenXL.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(OpenXL);
            Thread.Sleep(1000);
            Application.DoEvents();
        }

        private void Close_Tab(object sender, DrawItemEventArgs e)
        {
            // Using
            if (e.Index == MainTabControl.SelectedIndex)
            {
                e.Graphics.FillRectangle(Brushes.SkyBlue, e.Bounds);
                e.Graphics.DrawString(MainTabControl.TabPages[e.Index].Text,
                    new Font(MainTabControl.Font, FontStyle.Regular),
                    Brushes.Blue,
                    new PointF(e.Bounds.X + 2, e.Bounds.Y + 2));

            }
            else
            {
                e.Graphics.DrawString(MainTabControl.TabPages[e.Index].Text,
                    new Font(MainTabControl.Font, FontStyle.Regular),
                    Brushes.Black,
                    new PointF(e.Bounds.X + 2, e.Bounds.Y + 2));

            }
    }
 
        private void MainTabControl_MouseClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            TabControl tc = (TabControl)sender;
            Point p = e.Location;
            int _tabWidth = 0;
            _tabWidth = this.MainTabControl.GetTabRect(tc.SelectedIndex).Width - (_imgHitArea.X);
            Rectangle r = this.MainTabControl.GetTabRect(tc.SelectedIndex);
            r.Offset(_tabWidth, _imgHitArea.Y);
            r.Width = 16;
            r.Height = 16;
            if ((r.Contains(p)) && ((tc.SelectedIndex != 0) && (tc.SelectedIndex != 1)
                                 && (tc.SelectedIndex != 2) && (tc.SelectedIndex != 3)))
            {
                if (MessageBox.Show("Would you like to Close " + tc.TabPages[tc.SelectedIndex].Text + " ?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    TabPage TabP = (TabPage)tc.TabPages[tc.SelectedIndex];
                    tc.TabPages.Remove(TabP);
                    //TabP.Dispose();
                }
            }
        }

        private void MainTabControl_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.F)
            {
                //Search_Form();
            }
        }

        private void Text_Search_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter)
            //{
            //}
        }

        private void Item_Exception_CLick(object sender, EventArgs e)
        {
            //Create_Item_Exception_Manage_Tab();
        }

        private void User_Manage_Open_click(object sender, EventArgs e)
        {

        }

    }
}

