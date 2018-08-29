using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace kedouwenc
{
    public partial class jsontoarray : UserControl
    {
        public jsontoarray()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string str = textBox1.Text;
            string[] strarr = str.Split(Environment.NewLine.ToCharArray());
            //string resultarr = string.Empty; 
            int k = 0;
            string[,] resultarr = new string[strarr.Length, 8];

            for (int i = 0; i < strarr.Length; i++)
            {

                if (strarr[i].Contains("//"))
                {

                    //MessageBox.Show(strarr[i].ToString());
                    //MessageBox.Show(strarr[i].Split(':')[0]);
                    //MessageBox.Show(strarr[i].Split('/')[2]);
                    //resultarr = resultarr+strarr[i].Split(':')[0].Trim() + " " + strarr[i].Split('/')[2].Trim()+"\r\n";

                    resultarr[k, 0] = strarr[i].Split(':')[0].Trim().Replace("\"", "");
                    resultarr[k, 1] = strarr[i].Split('/')[2].Trim();
                    resultarr[k, 2] = "VARCHAR2";
                    resultarr[k, 7] = strarr[i].Split(':')[0].Trim().Replace("\"", "");
                    k = k + 1;

                }
            }
            Excel.Range rng = (Excel.Range)Globals.ThisAddIn.Application.InputBox(Prompt: "请选择存放区域，选择单个单元格即可", Type: 8);
            rng.Offset[0, 0].Resize[resultarr.GetUpperBound(0) + 1, 8].Value = resultarr;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = string.Empty;

        }
    }
}
