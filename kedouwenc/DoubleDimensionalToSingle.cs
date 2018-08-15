using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace kedouwenc
{
    public partial class DoubleDimensionalToSingle : Form
    {
        Excel.Range rng;
        Excel.Worksheet sht = Globals.ThisAddIn.Application.ActiveSheet;
        public DoubleDimensionalToSingle()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            object arr1; //声明变量
            //On Error Resume Next VBConversions Warning: On Error Resume Next not supported in C# //如果执行时出错，继续执行下一句代码
            this.Hide();
            //弹出输入框让用户选择待转换的目标区域           
            // rng = CType(xlapp.InputBox(Prompt:="Please select a range with your Mouse to be bolded.", Title:="SPECIFY RANGE", Type:=8), Excel.Range)
            rng = (Excel.Range)(Globals.ThisAddIn.Application.InputBox(Prompt: "请选择数据区域", Type: 8));
            if (rng == null)
            {
                return; //如果代码出错(按下了“取消”键)，那么结束过程
            }

            //将变量rng重置为rng与已用区域的交集，从而忽略空白区域，节约代码执行时间
            rng = Globals.ThisAddIn.Application.Intersect(rng, sht.UsedRange);
            //如果没有交集，那么结束过程
            if (rng == null)
            {
                MessageBox.Show(text: "请不要选择空白区域", caption: "友情提示", buttons: MessageBoxButtons.OK);
                return;
            }
            if (rng.Rows.Count < 3 || rng.Columns.Count < 3)
            {
                MessageBox.Show(text: "请选择不小于3行x3列的非空连续区域！", caption: "温馨提示");
                return;
            }
            arr1 = rng.Value; //将变量rng的值导入到变量arr1中，此时arr1成为数组变量
            textBox1.Text = rng.Address;
            checkedListBox1.CheckOnClick = true;
            checkedListBox1.MultiColumn = true;
            for (int i = 1; i <= rng.Columns.Count; i++)
            {
                checkedListBox1.Items.Add(rng.Cells[1, i].Value); //添加值
            }
            for (int j = 0; j <= checkedListBox1.Items.Count - 1; j++)
            {
                checkedListBox1.SetItemChecked(j, true); //set checkbox  select
            }
            checkedListBox1.SetItemChecked(0, false);
            this.Show();

        }
        private void button2_Click(object sender, EventArgs e)
        {
            int iSel = 0;
            int iCount;
            Excel.Range rngTarget; //待处理的单元格区域
            Excel.Range targetrng; //存放结果的单元格区域区域
            long lRow;
            long lCol;
            int iRows; //待处理单元格区域的总行数
            int iCols; //待处理单元格区域的总列数
            object[,] aTarget; //单元格区域的转化的二维数组
            string[,] aOut; //处理后的二维数组
            bool[] bSel; //用来判断列是否被checked的一维数组
            this.Hide();

            rngTarget = Globals.ThisAddIn.Application.Range[textBox1.Text];

            iRows = rngTarget.Rows.Count;
            iCols = rngTarget.Columns.Count;
            aTarget = rngTarget.Value;

            iCount = checkedListBox1.Items.Count - 1; //减去1 是因为bSel数组从0开始的。
            bSel = new bool[iCount + 1];

            for (int i = 0; i <= iCount; i++)
            {
                if (checkedListBox1.GetItemChecked(i) == true) //返回复选框是否被选中
                {
                    bSel[i] = true;
                    iSel++; //用来统计选中字段的个数
                }
            }

            //定义转换后的数组，留出标题一行，根据选择的列字段确定数组第二维，不选的字段数4加上序号列和合并项及值共3列
            //根据选择的列字段确定数组第一维，选择的字段数3乘以数据区域除去标题行的行数，再加一行标题
            aOut = new string[iSel * (iRows - 1) + 1, 2 + iCount - iSel + 1];
            aOut[0, 0] = "序号";
            aOut[0, 1] = "合并";
            aOut[0, 2] = "值";

            int col = 3 - 1; //1维表只有3列。序号||合并||值
            for (int i = 1; i <= iCount; i++)
            {
                if (checkedListBox1.GetItemChecked(i) == false)
                {
                    col++;
                    aOut[0, col] = checkedListBox1.Items[i].ToString(); //将字段未选择的加入到数组后面。
                }
            }

            lRow = 1; //初始化值1（留出标题一行）
            lCol = 3;
            // MessageBox.Show(Convert.ToString(aTarget.GetUpperBound(0) + "-" + aTarget.GetUpperBound(1)));
            for (int i = 2; i <= aTarget.GetUpperBound(0); i++)
            {
                for (int j = 1; j <= aTarget.GetUpperBound(1); j++)
                {
                    if (bSel[j - 1] == true) //如果该列需要转换
                    {
                        aOut[lRow, 0] = System.Convert.ToString(aTarget[i, 1]);
                        aOut[lRow, 1] = System.Convert.ToString(aTarget[1, j]); //选择的字段
                        aOut[lRow, 2] = System.Convert.ToString(aTarget[i, j]); //字段的值
                        //没有选择的字段的数据()
                        for (int k = 2; k <= aTarget.GetUpperBound(1); k++)
                        {
                            if (bSel[k - 1] == false)
                            {
                                aOut[lRow, lCol] = System.Convert.ToString(aTarget[i, k]);
                                lCol++;
                            }
                        }
                        lRow++;
                        lCol = 3;
                    }
                }
            }

            //弹出输入框，让用户指定转换后的数据的存放区域					
            targetrng = (Excel.Range)(Globals.ThisAddIn.Application.InputBox(Prompt: "请选择要二维表的存放区域，选择单个单元格即可", Title: "目标区域", Type: 8));
            dynamic with_1 = targetrng.Offset[0, 0].Resize[aOut.GetUpperBound(0) + 1, aOut.GetUpperBound(1) + 1]; //引用存放结果的目标区域]

            with_1.Value = aOut; //将数组arr2的值存入目标区域中
            with_1.EntireColumn.AutoFit(); //让目标区域自动调整列宽
            with_1.CurrentRegion.Borders.LineStyle = Excel.XlLineStyle.xlContinuous; //对目标区域添加边框
            //选中目标区域(采用Goto方法而不是Range.Select方法是为了避免目标区域与数据源不在同一个工作表时，无法选择目标区域)
            Globals.ThisAddIn.Application.Goto(with_1.Cells);
        }
        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void DoubleDimensionalToSingle_Load(object sender, EventArgs e)
        {

        }
    }
}


