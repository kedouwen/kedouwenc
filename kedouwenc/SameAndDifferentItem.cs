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
    public partial class SameAndDifferentItem : Form        
    {

        Excel.Worksheet sht = Globals.ThisAddIn.Application.ActiveSheet;
        Excel.Range rng1;        
        Excel.Range rng2;
        object[,] arr1;
        object[,] arr2;
        string[] brr1;
        string[] brr2;
        string[] brr3;


       

        public SameAndDifferentItem()
        {
            InitializeComponent();
        }
            
        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            try
            {
                rng1 = (Excel.Range)Globals.ThisAddIn.Application.InputBox(Prompt: "请选择数据区域一", Type: 8);               
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
            }
            catch
            {
                MessageBox.Show(text: "请不要选择空白区域", buttons: MessageBoxButtons.OK, caption: "");
                return;
            }  
            arr1 = rng1.Value;
            textBox1.Text = rng1.Address;
            this.Show();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            try
            {
                rng2 = (Excel.Range)Globals.ThisAddIn.Application.InputBox(Prompt: "请选择数据区域二", Type: 8);
            }
            catch
            {
                return;
            }
            try
            {
                rng2 = Globals.ThisAddIn.Application.Intersect(sht.UsedRange, rng2);
                if (rng2.Address == "$A$1" || rng1 == null)
                {
                    MessageBox.Show(text: "请不要选择空白区域", buttons: MessageBoxButtons.OK, caption: "");
                    return;
                }
            }
            catch
            {
                MessageBox.Show(text: "请不要选择空白区域", buttons: MessageBoxButtons.OK, caption: "");
                return;
            }
            arr2 = rng2.Value;
            textBox2.Text = rng2.Address;
            this.Show();

        }

        private void button3_Click(object sender, EventArgs e)
        {

            Dictionary<string, string> d1 = new Dictionary<string, string>();
            Dictionary<string, string> d2 = new Dictionary<string, string>();
            Dictionary<string, string> d3 = new Dictionary<string, string>();
            Dictionary<string, string> d4 = new Dictionary<string, string>();
            Dictionary<string, string> d5 = new Dictionary<string, string>();

            if (textBox1.Text.Length == 0 || textBox2.Text.Length == 0)
            {
                MessageBox.Show(text:"请先输入区域！");
            }

            foreach (object rng in arr1)
            {
                if(rng !=null )
                {
                    if (rng.ToString().Length != 0) 
                    {
                       // d1.Add(rng, "");  要判断是否有KEY键了
                        d1[rng.ToString()] = "";
                    }  
            }
            }

            foreach (object rng in arr2)
            {
                if (rng != null)
                {
                    if (rng.ToString().Length != 0)
                    {
                        //d2.Add(rng, "");
                        d2[rng.ToString()] = "";
                    }
                }                        
            }

            string[] crr1 = new string[d1.Keys.Count];
            d1.Keys.CopyTo(crr1, 0);

            string[] crr2 = new string[d2.Keys.Count];
            d2.Keys.CopyTo(crr2, 0);

            //找出相同项
            foreach (string str in crr2)
            {
                if (d1.ContainsKey(str))
                {
                    //d5.Add(str, "");                                   
                    d5[str] = "";
                }
            }

            //找出区域一独有
            foreach (string str in crr1)
            {
                if (!d2.ContainsKey(str))
                {
                    //d3.Add(str, "");
                    d3[str] = "";
                }
            }

            //找出区域二独有
            foreach (string str in crr2)
            {
                if (!d1.ContainsKey(str))
                {
                    //d3.Add(str, "");
                    d4[str] = "";
                }
            }
            
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            listBox3.Items.Clear();

            

            foreach (string str in d3.Keys)
            {
                listBox1.Items.Add(str);
            }

            foreach (string str in d4.Keys)
            {
                listBox2.Items.Add(str);
            }

            foreach (string str in d5.Keys)
            {
                listBox3.Items.Add(str);
            }

            for (int i = 5; i <= 9; i++)
            {
                if(i != 7)
                {
                    this.Controls["Button" + i].Enabled = true;
                }
            }

             brr1 = new string[listBox1.Items.Count];
            listBox1.Items.CopyTo(brr1, 0);

             brr2 = new string[listBox2.Items.Count];
            listBox2.Items.CopyTo(brr2, 0);

             brr3 = new string[listBox3.Items.Count];
            listBox3.Items.CopyTo(brr3, 0);

        }


        
        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();

        }

        private void button5_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.Union(rng1, rng2).Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
            foreach (Excel.Range rng in Globals.ThisAddIn.Application.Union(rng1, rng2))
            {

               // MessageBox.Show(Convert.ToString(rng.Value2));
                if (Array.IndexOf(brr3, Convert.ToString(rng.Value2)) != -1 )
                {
                    rng.Interior.ColorIndex = 15;
                }
            }

            this.button7.Enabled = true;
        }

        private void button6_Click(object sender, EventArgs e)
        {

            Globals.ThisAddIn.Application.Union(rng1, rng2).Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
            foreach (Excel.Range rng in Globals.ThisAddIn.Application.Union(rng1, rng2))
            {
               
                // MessageBox.Show(Convert.ToString(rng.Value2));
                if (Array.IndexOf(brr3, Convert.ToString(rng.Value2)) == -1 && rng.Value != null)               
                {
                    rng.Interior.ColorIndex = 15;
                }
            }

            this.button7.Enabled = true;



        }

        private void button7_Click(object sender, EventArgs e)
        {
            sht.Cells.Interior.Color = Excel.XlColorIndex.xlColorIndexNone;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            this.Hide();
            try
            {
                Excel.Range rng = (Excel.Range)Globals.ThisAddIn.Application.InputBox(Prompt: "请选择存放区域，选择单个单元格即可", Type: 8);
                rng.Offset[0, 0].Resize[1, 1].Value = "相同项";
                rng.Offset[1, 0].Resize[brr3.GetUpperBound(0) + 1, 1].Value = Globals.ThisAddIn.Application.WorksheetFunction.Transpose(brr3);
                this.Show();
            }
            catch
            {
                this.Show();
                return;               
            }
           
        }

        private void button9_Click(object sender, EventArgs e)
        {
            this.Hide();
            try
            {
                Excel.Range rng = (Excel.Range)Globals.ThisAddIn.Application.InputBox(Prompt: "请选择存放区域，选择单个单元格即可", Type: 8);
                rng.Cells[1,1] = "区域一独有";
                rng.Offset[0, 1].Resize[1, 1].Value = "区域二独有";
            

                if (brr1.GetUpperBound(0) != -1)
                {
                    rng.Offset[1, 0].Resize[brr1.GetUpperBound(0) + 1, 1].Value = Globals.ThisAddIn.Application.WorksheetFunction.Transpose(brr1);
                }

                 if (brr2.GetUpperBound(0)!= -1)
                 {
                     rng.Offset[1, 1].Resize[brr2.GetUpperBound(0) + 1, 1].Value = Globals.ThisAddIn.Application.WorksheetFunction.Transpose(brr2);
                 }
               

               
                this.Show();
            }
            catch
            {
                this.Show();                
                return;
            }
           
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

       

     
    }
}
