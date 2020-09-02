using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Interop.Office.Core
{
    public partial class Setting_Excel : Form
    {
        public Setting_Excel()
        {
            InitializeComponent();
        }

        private void Setting_Excel_Load(object sender, EventArgs e)
        {
            string exten = Properties.Settings.Default.ExcelExtension.ToLower();
            if (exten == ".xls")
            {
                this.radioButton1.Checked = true;
            }
            else if(exten ==".xlsx")
            {
                this.radioButton2.Checked = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string exten = "";
            if (this.radioButton1.Checked) exten = ".xls";
            else if (this.radioButton2.Checked) exten = ".xlsx";

            Properties.Settings.Default.ExcelExtension = exten;
            Properties.Settings.Default.Save();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
