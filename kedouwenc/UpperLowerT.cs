using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace kedouwenc
{
    public partial class UpperLowerT : Form
    {
        public UpperLowerT()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
             Excel.Worksheet sht = Globals.ThisAddIn.Application.ActiveSheet;
             Excel.Range rng;
             rng = Globals.ThisAddIn.Application.Intersect(Globals.ThisAddIn.Application.Selection, sht.UsedRange);
              foreach (Excel.Range myrng in rng)
              {
                  if (myrng.Value2 != null)
                  {
                      if (this.radioButton1.Checked)
                      {
                          myrng.Value = ((string)myrng.Value2).ToUpper();
                      }
                      else if (this.radioButton2.Checked)
                      {
                          myrng.Value = ((string)myrng.Value2).ToLower();
                      }
                      else
                      {
                          myrng.Value = ((string)myrng.Value2).Substring(0, 1).ToUpper() + ((string)myrng.Value2).Substring(1).ToLower();
                      }
                  }

              }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void UpperLowerT_Load(object sender, EventArgs e)
        {
            this.radioButton1.Checked = true;
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
