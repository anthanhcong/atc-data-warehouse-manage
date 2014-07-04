using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WarehouseManager
{
    class Button_Lbl : SQL_APPL
    {
        public Button My_Button;
        public Button_Lbl(System.Windows.Forms.TabPage owner_tab, string label,
                          PosSize possize, AnchorType anchor_type)
        {
            My_Button = new System.Windows.Forms.Button();
            owner_tab.Controls.Add(this.My_Button);
            // button1
            My_Button.Location = new System.Drawing.Point(possize.pos_x, possize.pos_y);
            My_Button.Name = label;
            My_Button.Size = new System.Drawing.Size(75, 23);
            My_Button.TabIndex = 0;
            My_Button.Text = label;
            My_Button.UseVisualStyleBackColor = true;
        }
    }

    class Checkbox_Lbl : SQL_APPL
    {
        public CheckBox My_CheckBox;
        public Checkbox_Lbl(System.Windows.Forms.TabPage owner_tab, string label,
                          PosSize possize, AnchorType anchor_type)
        {
            My_CheckBox = new System.Windows.Forms.CheckBox();
            owner_tab.Controls.Add(this.My_CheckBox);
            // button1
            My_CheckBox.Location = new System.Drawing.Point(possize.pos_x, possize.pos_y);
            My_CheckBox.AutoSize = true;
            My_CheckBox.Size = new System.Drawing.Size(80, 17);
            My_CheckBox.TabIndex = 0;
            My_CheckBox.Text = label;
            My_CheckBox.UseVisualStyleBackColor = true;
            My_CheckBox.Checked = true;
        }
    }

    class TextBox_Lbl : SQL_APPL
    {
        public TextBox My_TextBox = new TextBox();
        public Label My_Label = new Label();
        private TextBox_Type CheckType;
        PosSize My_PosSize = new PosSize();

        public TextBox_Lbl(System.Windows.Forms.TabPage owner_tab, string label,
                            TextBox_Type type, PosSize possize, AnchorType anchor_type)
        {
            My_PosSize = possize;
            CheckType = type;

            My_Label.ForeColor = Color.Black;
            My_Label.AutoSize = true;   
            My_Label.Visible = true;
            My_Label.Text = label + ":";
            My_Label.Location = new System.Drawing.Point(My_PosSize.pos_x, My_PosSize.pos_y+4);
            owner_tab.Controls.Add(My_Label);

            My_TextBox.Location = new System.Drawing.Point(My_PosSize.pos_x + 80, My_PosSize.pos_y);
            My_TextBox.Size = new System.Drawing.Size(90, 20);

            if (anchor_type == AnchorType.RIGHT)
            {
                My_Label.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                    | System.Windows.Forms.AnchorStyles.Right));
                My_TextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                            | System.Windows.Forms.AnchorStyles.Right));
            }
            else if (anchor_type == AnchorType.RIGHT)
            {
                My_Label.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                                   | System.Windows.Forms.AnchorStyles.Left));
                My_TextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                            | System.Windows.Forms.AnchorStyles.Left));
            }

            owner_tab.Controls.Add(My_TextBox);
        }
    }

    class TextBox_Lbl_F2 : SQL_APPL
    {
        public TextBox My_TextBox = new TextBox();
        public Label My_Label = new Label();
        private TextBox_Type CheckType;
        PosSize My_PosSize = new PosSize();

        public TextBox_Lbl_F2(System.Windows.Forms.TabPage owner_tab, string label, int height_tbx, int width_tbx,
                            TextBox_Type type, PosSize possize, AnchorType anchor_type)
        {
            My_PosSize = possize;
            CheckType = type;

            My_Label.ForeColor = Color.Black;
            My_Label.AutoSize = true;
            My_Label.Visible = true;
            My_Label.Text = label ;
            My_Label.Location = new System.Drawing.Point(My_PosSize.pos_x, My_PosSize.pos_y + 6);
            owner_tab.Controls.Add(My_Label);

            My_TextBox.Location = new System.Drawing.Point(My_PosSize.pos_x , My_PosSize.pos_y + 24);
            My_TextBox.Size = new System.Drawing.Size(width_tbx, height_tbx); // (70 ,20)

            if (anchor_type == AnchorType.RIGHT)
            {
                My_Label.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                    | System.Windows.Forms.AnchorStyles.Right));
                My_TextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                            | System.Windows.Forms.AnchorStyles.Right));
            }
            else if (anchor_type == AnchorType.RIGHT)
            {
                My_Label.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                                   | System.Windows.Forms.AnchorStyles.Left));
                My_TextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                            | System.Windows.Forms.AnchorStyles.Left));
            }

            owner_tab.Controls.Add(My_TextBox);
        }
    }

    class ComboBox_Lbl : SQL_APPL
    {
        public ComboBox My_Combo; // = new ComboBox();
        public Label My_Label;//  = new Label();
        private DataTable My_Table;
        PosSize My_PosSize = new PosSize();
        string My_display_member, My_value_member;

        public ComboBox_Lbl(System.Windows.Forms.TabPage owner_tab, string label, PosSize possize,
                            DataTable table, string display_member, string value_member, AnchorType anchor_type)
        {
            My_Combo = new ComboBox();
            My_Label = new Label();
            My_PosSize = possize;
            My_Table = table;
            My_display_member = display_member;
            My_value_member = value_member;

            My_Label.ForeColor = Color.Black;
            My_Label.AutoSize = true;
            My_Label.Visible = true;
            My_Label.Text = label + ":";
            My_Label.Location = new System.Drawing.Point(My_PosSize.pos_x, My_PosSize.pos_y+4);
            owner_tab.Controls.Add(My_Label);
            My_Label.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                                        | System.Windows.Forms.AnchorStyles.Right));
 
            My_Combo.DataSource = My_Table;
            My_Combo.DisplayMember = display_member;
            My_Combo.ValueMember = value_member;
            My_Combo.Size = new System.Drawing.Size(140, 20);
            My_Combo.Location = new System.Drawing.Point(My_PosSize.pos_x + 80, My_PosSize.pos_y);
            //My_Combo.Leave += new System.EventHandler(CheckCorrectValue);

            if (anchor_type == AnchorType.RIGHT)
            {
                My_Label.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                    | System.Windows.Forms.AnchorStyles.Right));
                My_Combo.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                            | System.Windows.Forms.AnchorStyles.Right));
            }
            else if (anchor_type == AnchorType.LEFT)
            {
                My_Label.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                                   | System.Windows.Forms.AnchorStyles.Left));
                My_Combo.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                            | System.Windows.Forms.AnchorStyles.Left));
            }

            owner_tab.Controls.Add(My_Combo);
        }

        private void CheckCorrectValue(object sender, EventArgs e)
        {
            string card_no = My_Combo.Text.Trim();

            if (Check_ItemExistTable(card_no, My_Table, My_display_member) == false)
            {
                MessageBox.Show("Can't Enter New Value", "Warning");
                if (My_Table.Rows.Count != 0)
                {
                    My_Combo.SelectedIndex = 0;
                }
                else
                {
                    My_Combo.SelectedIndex = -1;
                }
            }
        }
    }

    class RichText_Lbl : SQL_APPL
    {
        public RichTextBox My_RichText = new RichTextBox();
        private Label My_Label = new Label();
        PosSize My_PosSize = new PosSize();

        public RichText_Lbl(System.Windows.Forms.TabPage owner_tab, string label,
                            TextBox_Type type, PosSize possize, AnchorType anchor_type)
        {
            My_PosSize = possize;

            My_Label.ForeColor = Color.Black;
            My_Label.AutoSize = true;
            My_Label.Visible = true;
            My_Label.Text = label + ":";
            My_Label.Location = new System.Drawing.Point(My_PosSize.pos_x, My_PosSize.pos_y + 4);
            owner_tab.Controls.Add(My_Label);
            My_Label.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                                        | System.Windows.Forms.AnchorStyles.Right));

            My_RichText.Location = new System.Drawing.Point(My_PosSize.pos_x, My_PosSize.pos_y + 20);
            My_RichText.Size = new System.Drawing.Size(My_PosSize.width, My_PosSize.height);
            My_RichText.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                                        | System.Windows.Forms.AnchorStyles.Right));

            if (anchor_type == AnchorType.RIGHT)
            {
                My_Label.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                    | System.Windows.Forms.AnchorStyles.Right));
                My_RichText.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                            | System.Windows.Forms.AnchorStyles.Right));
            }
            else if (anchor_type == AnchorType.RIGHT)
            {
                My_Label.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                                   | System.Windows.Forms.AnchorStyles.Left));
                My_RichText.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                            | System.Windows.Forms.AnchorStyles.Left));
            }

            owner_tab.Controls.Add(My_RichText);
        }
    }

    class List_Lbl : SQL_APPL
    {
        private ListBox My_List = new ListBox();
        private Label My_Label = new Label();
        DataTable My_Table;
        PosSize My_PosSize = new PosSize();

        public List_Lbl (System.Windows.Forms.TabPage owner_tab, string label, PosSize possize,
                            DataTable table, string display_member, string value_member, AnchorType anchor_type)
        {
            My_PosSize = possize;
            My_Table = table;

            My_Label.ForeColor = Color.Black;
            My_Label.AutoSize = true;
            My_Label.Visible = true;
            My_Label.Text = label + ":";
            My_Label.Location = new System.Drawing.Point(My_PosSize.pos_x, My_PosSize.pos_y+4);
            owner_tab.Controls.Add(My_Label);
            My_Label.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                                        | System.Windows.Forms.AnchorStyles.Right));

            My_List.Location = new System.Drawing.Point(My_PosSize.pos_x, My_PosSize.pos_y+10);
            My_List.DisplayMember = display_member;
            My_List.ValueMember = value_member;
            My_List.DataSource = My_Table;
            My_List.Size = new System.Drawing.Size(My_PosSize.width, My_PosSize.height);
            My_List.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                                        | System.Windows.Forms.AnchorStyles.Right));

            if (anchor_type == AnchorType.RIGHT)
            {
                My_Label.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                    | System.Windows.Forms.AnchorStyles.Right));
                My_List.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                            | System.Windows.Forms.AnchorStyles.Right));
            }
            else if (anchor_type == AnchorType.RIGHT)
            {
                My_Label.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                                   | System.Windows.Forms.AnchorStyles.Left));
                My_List.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                            | System.Windows.Forms.AnchorStyles.Left));
            }

            owner_tab.Controls.Add(My_List);
        }
    }

    class DatePick_LBL : SQL_APPL
    {
        public DateTimePicker My_picker = new DateTimePicker();
        public Label My_Label = new Label();
        PosSize My_PosSize = new PosSize();

        public DatePick_LBL(System.Windows.Forms.TabPage owner_tab, string label,
                            PosSize possize, AnchorType anchor_type)
        {
            My_PosSize = possize;

            My_Label.ForeColor = Color.Black;
            My_Label.AutoSize = true;
            My_Label.Visible = true;
            My_Label.Text = label + ":";
            My_Label.Location = new System.Drawing.Point(My_PosSize.pos_x, My_PosSize.pos_y + 4);
            owner_tab.Controls.Add(My_Label);

            My_picker.Location = new System.Drawing.Point(My_PosSize.pos_x + 70, My_PosSize.pos_y);
            My_picker.Size = new System.Drawing.Size(100, 20);
            My_picker.Format = DateTimePickerFormat.Custom;
            My_picker.CustomFormat = "dd-MMM-yyyy";

            if (anchor_type == AnchorType.RIGHT)
            {
                My_Label.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                    | System.Windows.Forms.AnchorStyles.Right));
                My_picker.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                            | System.Windows.Forms.AnchorStyles.Right));
            }
            else if (anchor_type == AnchorType.RIGHT)
            {
                My_Label.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                                   | System.Windows.Forms.AnchorStyles.Left));
                My_picker.Anchor = ((System.Windows.Forms.AnchorStyles)(System.Windows.Forms.AnchorStyles.Top
                            | System.Windows.Forms.AnchorStyles.Left));
            }
            owner_tab.Controls.Add(My_picker);
        }
    }
}
