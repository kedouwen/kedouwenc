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
    public partial class oracle_createtablesql : UserControl
    {
        public oracle_createtablesql()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            textBox1.Text = "";
            Excel.Range rng;
            object[,] arr;


            string tablename_chs = textBox3.Text;
            TextToPinyinForm tempform = new TextToPinyinForm();

            string tablename = tempform.ConvertCnToPinyinABC(tablename_chs);

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

            string tempcreateuuid = "DATA_UP_UUID  VARCHAR2(100)  default sys_guid() NOT NULL,";
            string tempcreatedateup = ",DATA_UP_TIME  date,DATA_UP_STATUS  varchar2(100));";

            ceratetable = "CREATE TABLE " + tablename + "(" + tempcreateuuid + "\r\n" + tempcreatetable.Substring(0, tempcreatetable.Length - 3) + tempcreatedateup;


            commentcolumn = commentcolumn + "comment on column " + tablename + ".DATA_UP_UUID is '主键';" +
                "\r\n" + "comment on column " + tablename + ".DATA_UP_TIME is '数据入库时间';" +
                "\r\n" + "comment on column " + tablename + ".DATA_UP_STATUS is '数据状态(I新增数据、U更新数据、D删除数据)';";


            string tempcommenttable = "comment on table " + tablename + " is '" + tablename_chs + "';";
            string tempcreatepk = "alter table " + tablename + " add constraint PK_" + tablename + "_UUID primary key (DATA_UP_UUID);";

            textBox1.Text = ceratetable + "\r\n" + commentcolumn + "\r\n" + tempcommenttable + "\r\n" + tempcreatepk;

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
