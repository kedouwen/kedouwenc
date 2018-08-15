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
    public partial class ForDisplay : Form
    {
        public ForDisplay()
        {
            InitializeComponent();
            checkBox1.Checked = Globals.ThisAddIn.Application.ActiveWindow.DisplayGridlines;
            checkBox2.Checked = Globals.ThisAddIn.Application.ActiveWindow.DisplayVerticalScrollBar;
            checkBox3.Checked = Globals.ThisAddIn.Application.ActiveWindow.DisplayHorizontalScrollBar;
            checkBox4.Checked = Globals.ThisAddIn.Application.ActiveWindow.DisplayHeadings;
            checkBox5.Checked = Globals.ThisAddIn.Application.ActiveWindow.DisplayRightToLeft;
            checkBox6.Checked = Globals.ThisAddIn.Application.ActiveWindow.DisplayFormulas;
            checkBox7.Checked = Globals.ThisAddIn.Application.ActiveWindow.DisplayZeros;
            checkBox8.Checked = Globals.ThisAddIn.Application.ActiveWindow.DisplayWorkbookTabs;
            checkBox9.Checked = Globals.ThisAddIn.Application.DisplayStatusBar;
            checkBox10.Checked = Globals.ThisAddIn.Application.DisplayFormulaBar;
            if(Globals.ThisAddIn.Application.ActiveWorkbook.DisplayDrawingObjects == Excel.XlDisplayDrawingObjects.xlHide)
            {
                checkBox9.Checked=true;
            }
            else
            {
                checkBox9.Checked=false;
            }
            checkBox12.Checked = Globals.ThisAddIn.Application.DisplayFullScreen;
            checkBox13.Checked = Globals.ThisAddIn.Application.ShowWindowsInTaskbar;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveWindow.DisplayGridlines = checkBox1.Checked;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveWindow.DisplayVerticalScrollBar = checkBox2.Checked;
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveWindow.DisplayHorizontalScrollBar = checkBox3.Checked;
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveWindow.DisplayHeadings = checkBox4.Checked;
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveWindow.DisplayRightToLeft = checkBox5.Checked;
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveWindow.DisplayFormulas=checkBox6.Checked;
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveWindow.DisplayZeros = checkBox7.Checked;
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveWindow.DisplayWorkbookTabs=checkBox8.Checked;
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.DisplayStatusBar=checkBox9.Checked;
        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.DisplayFormulaBar=checkBox10.Checked;
        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
             if(checkBox11.Checked == true)
            {
            Globals.ThisAddIn.Application.ActiveWorkbook.DisplayDrawingObjects = Excel.XlDisplayDrawingObjects.xlHide;
            }
            else
            {
                Globals.ThisAddIn.Application.ActiveWorkbook.DisplayDrawingObjects = Excel.XlDisplayDrawingObjects.xlDisplayShapes;
            }
        }

        private void checkBox12_CheckedChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.DisplayFullScreen = checkBox12.Checked;
        }

        private void checkBox13_CheckedChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.ShowWindowsInTaskbar = checkBox13.Checked;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
             if (comboBox1.Text == "显示批注")
            {
            Globals.ThisAddIn.Application.DisplayCommentIndicator = Excel.XlCommentDisplayMode.xlCommentAndIndicator;
            }
            else if(comboBox1.Text == "隐藏批注")
            {
            Globals.ThisAddIn.Application.DisplayCommentIndicator = Excel.XlCommentDisplayMode.xlNoIndicator;
            }
            else if(comboBox1.Text == "批注指示符")
            {
            Globals.ThisAddIn.Application.DisplayCommentIndicator = Excel.XlCommentDisplayMode.xlCommentIndicatorOnly;
            }            

       
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Worksheet myselsht = Globals.ThisAddIn.Application.ActiveSheet;
            string shtname = myselsht.Name;
            Globals.ThisAddIn.Application.Application.ScreenUpdating = false;
            foreach (Excel.Worksheet sht in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
            {
                ((Microsoft.Office.Interop.Excel._Worksheet)sht).Activate();
                Globals.ThisAddIn.Application.ActiveWindow.DisplayGridlines = checkBox1.Checked;
                Globals.ThisAddIn.Application.ActiveWindow.DisplayVerticalScrollBar = checkBox2.Checked;
                Globals.ThisAddIn.Application.ActiveWindow.DisplayHorizontalScrollBar = checkBox3.Checked;
                Globals.ThisAddIn.Application.ActiveWindow.DisplayHeadings = checkBox4.Checked;
                Globals.ThisAddIn.Application.ActiveWindow.DisplayRightToLeft = checkBox5.Checked;
                Globals.ThisAddIn.Application.ActiveWindow.DisplayFormulas = checkBox6.Checked;
                Globals.ThisAddIn.Application.ActiveWindow.DisplayZeros = checkBox7.Checked;
                Globals.ThisAddIn.Application.ActiveWindow.DisplayWorkbookTabs = checkBox8.Checked;
                Globals.ThisAddIn.Application.DisplayStatusBar = checkBox9.Checked;
                Globals.ThisAddIn.Application.DisplayFormulaBar = checkBox10.Checked;

                if (checkBox11.Checked == true)
                {
                    Globals.ThisAddIn.Application.ActiveWorkbook.DisplayDrawingObjects = Excel.XlDisplayDrawingObjects.xlHide;
                }
                else
                {
                    Globals.ThisAddIn.Application.ActiveWorkbook.DisplayDrawingObjects = Excel.XlDisplayDrawingObjects.xlDisplayShapes;
                }

                Globals.ThisAddIn.Application.DisplayFullScreen = checkBox12.Checked;
                Globals.ThisAddIn.Application.ShowWindowsInTaskbar = checkBox13.Checked;


                if (comboBox1.Text == "显示批注")
                {
                    Globals.ThisAddIn.Application.DisplayCommentIndicator = Excel.XlCommentDisplayMode.xlCommentAndIndicator;
                }
                else if (comboBox1.Text == "隐藏批注")
                {
                    Globals.ThisAddIn.Application.DisplayCommentIndicator = Excel.XlCommentDisplayMode.xlNoIndicator;
                }
                else if (comboBox1.Text == "批注指示符")
                {
                    Globals.ThisAddIn.Application.DisplayCommentIndicator = Excel.XlCommentDisplayMode.xlCommentIndicatorOnly;
                }


            }

            Globals.ThisAddIn.Application.Worksheets[shtname].select();
            Globals.ThisAddIn.Application.Application.ScreenUpdating = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            checkBox1.Checked = true;
            checkBox2.Checked = true;
            checkBox3.Checked = true;
            checkBox4.Checked = true;
            checkBox7.Checked = true;
            checkBox8.Checked = true;
            checkBox9.Checked = true;
            checkBox10.Checked = true;
            checkBox13.Checked = true;
            comboBox1.Text = "批注指示符";
            
            checkBox5.Checked = false;
            checkBox6.Checked = false;
            checkBox11.Checked = false;
            checkBox12.Checked = false;

        }

       

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
            Globals.ThisAddIn.Application.Application.ScreenUpdating = false;

            foreach (Excel.Workbook wkb in Globals.ThisAddIn.Application.Workbooks)
            {
                ((Microsoft.Office.Interop.Excel._Workbook)wkb).Activate();
                Excel.Worksheet myselsht = Globals.ThisAddIn.Application.ActiveSheet;
                string shtname = myselsht.Name;
                foreach (Excel.Worksheet sht in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
                {
                    ((Microsoft.Office.Interop.Excel._Worksheet)sht).Activate();
                    Globals.ThisAddIn.Application.ActiveWindow.DisplayGridlines = checkBox1.Checked;
                    Globals.ThisAddIn.Application.ActiveWindow.DisplayVerticalScrollBar = checkBox2.Checked;
                    Globals.ThisAddIn.Application.ActiveWindow.DisplayHorizontalScrollBar = checkBox3.Checked;
                    Globals.ThisAddIn.Application.ActiveWindow.DisplayHeadings = checkBox4.Checked;
                    Globals.ThisAddIn.Application.ActiveWindow.DisplayRightToLeft = checkBox5.Checked;
                    Globals.ThisAddIn.Application.ActiveWindow.DisplayFormulas = checkBox6.Checked;
                    Globals.ThisAddIn.Application.ActiveWindow.DisplayZeros = checkBox7.Checked;
                    Globals.ThisAddIn.Application.ActiveWindow.DisplayWorkbookTabs = checkBox8.Checked;
                    Globals.ThisAddIn.Application.DisplayStatusBar = checkBox9.Checked;
                    Globals.ThisAddIn.Application.DisplayFormulaBar = checkBox10.Checked;

                    if (checkBox11.Checked == true)
                    {
                        Globals.ThisAddIn.Application.ActiveWorkbook.DisplayDrawingObjects = Excel.XlDisplayDrawingObjects.xlHide;
                    }
                    else
                    {
                        Globals.ThisAddIn.Application.ActiveWorkbook.DisplayDrawingObjects = Excel.XlDisplayDrawingObjects.xlDisplayShapes;
                    }

                    Globals.ThisAddIn.Application.DisplayFullScreen = checkBox12.Checked;
                    Globals.ThisAddIn.Application.ShowWindowsInTaskbar = checkBox13.Checked;


                    if (comboBox1.Text == "显示批注")
                    {
                        Globals.ThisAddIn.Application.DisplayCommentIndicator = Excel.XlCommentDisplayMode.xlCommentAndIndicator;
                    }
                    else if (comboBox1.Text == "隐藏批注")
                    {
                        Globals.ThisAddIn.Application.DisplayCommentIndicator = Excel.XlCommentDisplayMode.xlNoIndicator;
                    }
                    else if (comboBox1.Text == "批注指示符")
                    {
                        Globals.ThisAddIn.Application.DisplayCommentIndicator = Excel.XlCommentDisplayMode.xlCommentIndicatorOnly;
                    }

                }
                Globals.ThisAddIn.Application.Worksheets[shtname].select();
            }

            
            Globals.ThisAddIn.Application.Application.ScreenUpdating = true;
        }
       

    }
}
