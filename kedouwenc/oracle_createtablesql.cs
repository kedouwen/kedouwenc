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
    public partial class oracle_createtablesql : UserControl
    {
        public oracle_createtablesql()
        {
            InitializeComponent();
        }

        //判断输入的字符是否符合格式
        private bool CheckInputText(string inputStr)
        {
            //字符串不能为空
            if (string.IsNullOrEmpty(inputStr))
                return false;
            //判断字符串是否完全是中文
            //匹配中文字符的正则表达式
            //还要匹配换行符
            //string patternCN = @"^[\u4e00-\u9fa5\r\n]+$";
            string patternCN = @"^[A-Za-z0-9]+$";

            //匹配表达式，如果匹配成功，说明格式符合要求
            if (Regex.IsMatch(inputStr, patternCN))
            {
                return true;
            }
            else
                return false;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            textBox1.Text = "";
            Excel.Range rng;
            object[,] arr;
            string tablename_chs = textBox3.Text;
            string tablename = tablename_chs;
            if (!CheckInputText(tablename_chs))
            {
                TextToPinyinForm tempform = new TextToPinyinForm();
                tablename = tempform.ConvertCnToPinyinABC(tablename_chs);
            }


            string tempcreatetable = "";
            string commentcolumn = "";
            string ceratetable;


            rng = Globals.ThisAddIn.Application.Selection;
            arr = rng.Value;

            for (int i = 2; i <= arr.GetUpperBound(0); i++)
            {
                tempcreatetable = tempcreatetable + arr[i, 1] + "  " + arr[i, 2] + ",\r\n";
                commentcolumn = commentcolumn + "comment on column " + tablename + "." + arr[i, 1] + " is '" + arr[i, 3] + "';\r\n";
            }
            
            ceratetable = "CREATE TABLE " + tablename + "("+ tempcreatetable.Substring(0, tempcreatetable.Length - 3)+ ");";
            string tempcommenttable = "comment on table " + tablename + " is '" + tablename_chs + "';";
            //string tempcreatepk = "alter table " + tablename + " add constraint PK_" + tablename + "_UUID primary key (DATA_UP_UUID);";

            textBox1.Text = ceratetable + "\r\n" + commentcolumn + "\r\n" + tempcommenttable;

            if (textBox1.Text != "")
                Clipboard.SetDataObject(textBox1.Text); 



        }

        private void button2_Click(object sender, EventArgs e)
        {
            Excel.Range rng1;
            Excel.Worksheet sht = Globals.ThisAddIn.Application.ActiveSheet;
                        
            try
            {
                rng1 = (Excel.Range)Globals.ThisAddIn.Application.InputBox(Prompt: "请选择单个单元格", Type: 8);
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

                if (rng1.Count != 1)
                {
                    MessageBox.Show(text: "请选择单个单元格", buttons: MessageBoxButtons.OK, caption: "");
                    return;
                }
               
            }
            catch
            {
                MessageBox.Show(text: "请不要选择空白区域", buttons: MessageBoxButtons.OK, caption: "", icon: MessageBoxIcon.Warning);
                return;
            }


            textBox3.Text =  Convert.ToString(rng1.Value) ;
          
        }

       
    }
}
