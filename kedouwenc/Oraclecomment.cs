using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace kedouwenc
{
    public partial class Oraclecomment : UserControl
    {
        public Oraclecomment()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Range rng1;
            Excel.Worksheet sht = Globals.ThisAddIn.Application.ActiveSheet;
            textBox1.Text = "";
            object[,] arr;
            string commenttable;
            string tablename;
            string commentcolumn="";
            Regex regEnglish = new Regex("^[a-zA-Z]");
            Int16 columnEnglish;
            Int16 columnChinese;
            try
            {
                rng1 = (Excel.Range)Globals.ThisAddIn.Application.InputBox(Prompt: "请选择单元格区域", Type: 8);
            }
            catch
            {
                return;
            }


            try
            {
                rng1 = Globals.ThisAddIn.Application.Intersect(sht.UsedRange, rng1);
                if (rng1.Address == "$A$1" || rng1 == null)
                {
                    MessageBox.Show(text: "请不要选择空白区域", buttons: MessageBoxButtons.OK, caption: "");
                    return;
                }

                if (rng1.Count == 1)
                {
                    MessageBox.Show(text: "请选择多个单元格", buttons: MessageBoxButtons.OK, caption: "");
                    return;
                }

            }
            catch
            {
                MessageBox.Show(text: "请不要选择空白区域", buttons: MessageBoxButtons.OK, caption: "", icon: MessageBoxIcon.Warning);
                return;
            }


            arr = rng1.Value2;
            tablename = arr[2, 1].ToString().Split('.')[1];

            //Add comments to the table
            commenttable = "comment on table " + tablename + "  is '" + arr[1, 1] + "';";
            textBox1.Text = commenttable;


            if (regEnglish.IsMatch(arr[3, 1].ToString())) {
                columnEnglish = 1;
                columnChinese = 2;
            }
            else {
                columnEnglish = 2;
                columnChinese = 1;
            }
            


            for (int i = 3; i <= arr.GetUpperBound(0); i++)
            {               
                commentcolumn = commentcolumn + "comment on column " + tablename + "." + arr[i, columnEnglish] + " is '" + arr[i, columnChinese] + "';\r\n";
            }
            
            textBox1.Text = commenttable + "\r\n" + commentcolumn;


            if (textBox1.Text != "")
                Clipboard.SetDataObject(textBox1.Text);
        }
    }
}
