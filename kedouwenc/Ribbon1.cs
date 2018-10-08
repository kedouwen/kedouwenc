using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms; //命名空间引入；
using Microsoft.VisualBasic;//可以引用VB.NET中的IsNumberic等函数


//return跳出方法
//break跳出循环


// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace kedouwenc
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        public static string gifid = "";
        public static Boolean ispressed;
        public string StrLabel;
        internal Microsoft.Office.Tools.CustomTaskPane OracleCreateTableSqlTaskPane;
        internal Microsoft.Office.Tools.CustomTaskPane CreateJsonTaskPane;
        internal Microsoft.Office.Tools.CustomTaskPane oraclecommentTaskPane;
        internal Microsoft.Office.Tools.CustomTaskPane jsontoarrayTaskPane;

        



        public Ribbon1()
        {
            //
        }

        /// <summary>
        /// 阅读模式
        /// </summary>
        /// <param name="control"></param>
        /// <param name="pressed"></param>
        public void ReadMode(Office.IRibbonControl control, Boolean pressed = true)
        {
            ispressed = pressed;
            //MessageBox.Show(Convert.ToString(Ribbon1.ispressed) + "666");
        }



        /// <summary>
        /// 显示设置
        /// </summary>
        /// <param name="control"></param>
        public void ForDisplay(Office.IRibbonControl control)
        {
            gifid = control.Id;
            //MessageBox.Show(gifid);
            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                Form Gif_Help = new Gif_Help();
                Gif_Help.TopMost = true;

                Gif_Help.Show();
            }
            else
            {
                Form frm = new ForDisplay();
                frm.TopMost = true;
                frm.Show();
            }
        }

        /// <summary>
        /// 从当前开始
        /// </summary>
        /// <param name="control"></param>
        public void FromHere(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.Goto(Globals.ThisAddIn.Application.ActiveWindow.RangeSelection.Offset, Scroll: true);
            // Excel.Range rng =Globals.ThisAddIn.Application.ActiveWindow.RangeSelection;
            // rng.EntireRow.Interior.Color = 255;
            // rng.EntireColumn.Interior.Color = 255;
        }


        /// <summary>
        /// 文本型数字转数值
        /// </summary>
        /// <param name="control"></param>
        public void StringToNumer(Office.IRibbonControl control)
        {
            Excel.Worksheet sht = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range rngs;

            Globals.ThisAddIn.Application.ScreenUpdating = false;
            sht = Globals.ThisAddIn.Application.ActiveSheet;
            rngs = Globals.ThisAddIn.Application.Intersect(Globals.ThisAddIn.Application.Selection, sht.UsedRange);

            //如果没有交集，那么结束过程
            if (rngs == null)
            {                
                RangeSelectHelper.blankrange();
                return;
            }

            try
            {
                //rngs = rngs.SpecialCells(XlCellType.xlCellTypeConstants, 23);               
                rngs.NumberFormatLocal = "G/通用格式";
                rngs.Value = rngs.Value;
            }
            catch
            {
                RangeSelectHelper.textrange();
            }
            //int icount = 0;
            //foreach (Excel.Range myrng in rngs)
            //{
            //    if (!rngs.HasFormula)
            //    {
            //        icount = icount + 1;
            //        break;
            //    }
            //}
            //if (icount == 0)
            //{
            //    MessageBox.Show("选择的单元格没有文本区域", "友情提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //}
            //else
            //{               
            //        if (rngs != null)
            //        {
            //            rngs.NumberFormatLocal = "G/通用格式";
            //            rngs.Value = rngs.Value;

            //         //  rng.Copy();  选择性粘贴的方法
            //         //  rng.PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationAdd, false, false);

            //     }
            //}

            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        /// <summary>
        /// 数值转文本型数值
        /// </summary>
        /// <param name="control"></param>
        public void NumerToString(Office.IRibbonControl control)
        {
            Excel.Worksheet sht = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range rng;
            object[,] arr; //单元格区域的转化的二维数组
            //myudf udf = new myudf();

            Globals.ThisAddIn.Application.ScreenUpdating = false;
            rng = Globals.ThisAddIn.Application.Intersect(Globals.ThisAddIn.Application.Selection, sht.UsedRange);
            arr = rng.Value;            

            //如果没有交集，那么结束过程
            if (rng == null)
            {                
                RangeSelectHelper.blankrange();
                return;
            }
            int icount = 0;
            foreach (Excel.Range myrng in rng)
            {
                if (Information.IsNumeric(myrng.Value) && !rng.HasFormula)
                {
                    icount = icount + 1;
                    break;
                }
            }
            if (icount == 0)
            {
                RangeSelectHelper.valuerange();
            }
            else
            {
                rng.NumberFormatLocal = "@";
                //EXCEL-range里面循环太慢 改成数组；
                for (int i = 1; i <= arr.GetUpperBound(0); i++)
                {
                    for (int j = 1; j <= arr.GetUpperBound(1); j++)
                    {
                        arr[i, j] = "'" + System.Convert.ToString(arr[i, j]);
                    }
                }
                rng.Value = arr;

            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        /// <summary>
        /// 公式转数值
        /// 20180329 如果公式区域不是连续的，比如合并单元格填充会存在问题
        /// </summary>
        /// <param name="control"></param>
        public void FormulaToNumber(Office.IRibbonControl control)
        {
            Excel.Worksheet sht = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range rng;

            Globals.ThisAddIn.Application.ScreenUpdating = false;
            rng = Globals.ThisAddIn.Application.Intersect(Globals.ThisAddIn.Application.Selection, sht.UsedRange);

            //循环比较慢；
            //MessageBox.Show(rng.Address);
            //int icount = 0;
            //foreach (Excel.Range myrng in rng)
            //{
            //    if (myrng.HasFormula)
            //    {
            //        icount = icount + 1;
            //        break;
            //    }
            //}

            //if (icount == 0)
            //{
            //    //Interaction.MsgBox("选择的单元格没有公式。", MsgBoxStyle.Information);
            //    MessageBox.Show(text: "选择的单元格没有公式。", caption: "提示", buttons: MessageBoxButtons.OK);
            //}
            //else
            //{
            //    rng.NumberFormatLocal = "G/通用格式";
            //    rng.Value = rng.Value;                
            //}

            try
            {
                rng = rng.SpecialCells(XlCellType.xlCellTypeFormulas, 23);
                //MessageBox.Show(rng.Address);
                foreach (Excel.Range therng in rng)
                {
                    therng.NumberFormatLocal = "G/通用格式";
                    therng.Value = therng.Value;
                }

            }
            catch
            {
                RangeSelectHelper.formularange();
            }

            Globals.ThisAddIn.Application.ScreenUpdating = true;

        }

        /// <summary>
        /// 大小写转换
        /// </summary>
        /// <param name="control"></param>
        public void UpperLowerT(Office.IRibbonControl control)
        {
            Form frm = new UpperLowerT();
            frm.TopMost = true;
            frm.Show();
        }


        /// <summary>
        /// 区域分列转换
        /// </summary>
        /// <param name="control"></param>
        public void TranRangeByColumn(Office.IRibbonControl control)
        {
            //On Error Resume Next
            object[,] arr;
            int LBd;
            //数组上标
            int ubd;
            //数组下标
            int vmod;
            int blbd;
            //新数组的上标
            int m = 0;
            Excel.Range sourerange;
            //转换的区域
            Excel.Range targetrange;
            //存放的区域
            int k, x, y, z;
            try
            {
                sourerange = Globals.ThisAddIn.Application.InputBox(Prompt: "请输入转换的区域", Type: 8);
            }
            catch
            {
                return;
            }

            if (sourerange == null)
            {
                return;
            }
            //区域的值赋给arr
            arr = ((Excel.Range)sourerange).Value2;
            //inputbox点击取消后，返回值 FALSE，将其转换成数值
            k = Convert.ToInt32(Globals.ThisAddIn.Application.InputBox(Prompt: "请输入转换后的列数：", Title: "转换后的列数", Type: 1));
            if (k == 0)
            {
                MessageBox.Show("请不要点击取消或者输入0");
                return;
            }
            //可以直接使用数组的Get方法
            LBd = arr.GetUpperBound(0);
            ubd = arr.GetUpperBound(1);
            vmod = ubd % k;

            if (vmod == 0)
            {
                blbd = (LBd * ubd) / k;
            }
            else
            {
                blbd = (LBd * (ubd - vmod)) / k + LBd;
            }

            //  定义新的数组brr 其中的blbd,k是对应的个数。
            string[,] brr = new string[blbd, k];
            for (y = 1; y <= ubd; y += k)
            {
                for (x = 1; x <= LBd; x++)
                {
                    for (z = 0; z <= k - 1; z++)
                    {
                        if (y + z <= ubd)
                            //object数组用TOSTRING()方法转换下
                            brr[m, z] = arr[x, y + z].ToString();
                    }
                    m = m + 1;
                }
            }
            targetrange = Globals.ThisAddIn.Application.InputBox(Prompt: "请输入：", Title: "存放的区域", Type: 8);
            if (targetrange == null)
            {
                return;
            }

            //也可以调用vb.net的Information类。
            targetrange.get_Resize(Information.UBound(brr) + 1, Information.UBound(brr, 2) + 1).Value2 = brr;
        }

        /// <summary>
        /// 二维表转一维表
        /// </summary>
        /// <param name="control"></param>
        public void DoubleDimensionalToSingle(Office.IRibbonControl control)
        {
            Form frm = new DoubleDimensionalToSingle();
            frm.TopMost = true;
            frm.Show();
        }



        /// <summary>
        /// 删除空行
        /// rng.MergeCell 说明：True if the range contains merged cells. 
        /// 1.选择单个合并单元格，会返回ture;
        /// 2.选择一个区域，里面有合并单元格，会返回DBnull
        /// 3.选择非合并单元格区域，返回true;
        /// 
        /// </summary>
        /// <param name="control"></param>
        public void DelBlankRow(Office.IRibbonControl control)
        {
            //有一个大坑，就是单元格区域复制给数组，在本地窗口里面看 数组是0维开始，其实从1开始的。
            //On Error Resume Next VBConversions Warning: On Error Resume Next not supported in C#
            Globals.ThisAddIn.Application.ScreenUpdating = false; //关闭屏幕刷新
            Excel.Range rng;
            object[,] arr2;
            long i;
            byte j;
            Excel.Worksheet sht = Globals.ThisAddIn.Application.ActiveSheet; ;
            //将活动工作表的已用区域与第一行到最后一个非空行之间的区域的交集赋与变量rng	
            //If Globals.ThisAddIn.Application.WorksheetFunction.CountA(sht.UsedRange.Cells) = 0 Then Exit Sub
            rng = Globals.ThisAddIn.Application.Intersect(sht.UsedRange, Globals.ThisAddIn.Application.Rows["1:" + Globals.ThisAddIn.Application.Cells.Find(What: "*", After: Globals.ThisAddIn.Application.Cells[1, 1], LookIn: XlFindLookIn.xlValues, LookAt: XlLookAt.xlWhole, SearchOrder: XlSearchOrder.xlByRows, SearchDirection: XlSearchDirection.xlPrevious).Row]);
            

            if (Convert.IsDBNull(rng.MergeCells) || (!Convert.IsDBNull(rng.MergeCells) && rng.MergeCells))
            {
                MessageBox.Show("存在合并单元格区域！！！", "友情提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Globals.ThisAddIn.Application.ScreenUpdating = true; //恢复屏幕刷新
                return;
            }



            arr2 = rng.Value; //将rng区域的值赋予变量arr2
            int[,] arr = new int[arr2.GetUpperBound(0), 1]; //重置数组变量Arr的维数与上、下标，其中第一维上标等于arr2的上标
            for (i = 1; i <= arr2.GetUpperBound(0); i++) //遍历数组arr2的每一行
            {
                for (j = 1; j <= arr2.GetUpperBound(1); j++) //遍历数组arr2的每一列d
                {
                    arr[i - 1, 0] = arr[i - 1, 0] + Strings.Len(arr2[i, j]); //计算数组arr2的第i行、第j列的字符数量，并累加到数组arr中
                    if (arr[i - 1, 0] > 0)
                    {
                        arr[i - 1, 0] = -1; //如果数组arr的第i行的值大于0，那么将它重置为-1
                    }
                }
            }
            //在数据区域的右边一列创建辅助列，在辅助列中存放数组arr的值
            //数组arr的值由-1和0组成，如果等于0，表示该行空白，如果等于-1，表示该行有数据。
            rng.Offset[0, rng.Columns.Count].Columns[1] = arr;
            dynamic with_1 = sht.Sort; //引用活动工作表的Sort对象
            with_1.SortFields.Clear(); //清除以前设置的排序字段
            //添加一个新的排序字段，Key来自辅助列。排序方式为按数值大小升序排列
            with_1.SortFields.Add(rng.Offset[0, rng.Columns.Count].Columns[1], Excel.XlSortOn.xlSortOnValues, Excel.XlOrder.xlDownThenOver);
            //指定参与排序的整个区域(即包含rng与它右边一列的区域)
            with_1.SetRange(Globals.ThisAddIn.Application.Union(rng, rng.Offset[0, 1]));
            with_1.Apply(); //执行排序
            rng.Offset[0, rng.Columns.Count].Columns[1].Clear(); //删除辅助列的值
            //删除空行的格式信息(空行是指最后一个非空行下方的所有行)
            Globals.ThisAddIn.Application.Rows[(Globals.ThisAddIn.Application.Cells.Find(What: "*", After: Globals.ThisAddIn.Application.Cells[1, 1], LookIn: XlFindLookIn.xlValues, LookAt: XlLookAt.xlWhole, SearchOrder: XlSearchOrder.xlByRows, SearchDirection: XlSearchDirection.xlPrevious).Row + 1) + ":" + Globals.ThisAddIn.Application.Rows.Count].Clear();
            Globals.ThisAddIn.Application.ScreenUpdating = true; //恢复屏幕刷新
        }

        /// <summary>
        /// 隐藏非选中区域
        /// </summary>
        /// <param name="control"></param>
        public void HideNoSelectRange(Office.IRibbonControl control)
        {
            Excel.Range rng;
            Excel.Range myrng;
            Excel.Range rngleft;
            Excel.Range rngright;
            Excel.Range rngup;
            Excel.Range rngdown;
            Excel.Worksheet sht = Globals.ThisAddIn.Application.ActiveSheet;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            myrng = Globals.ThisAddIn.Application.Selection;
            rng = sht.Cells[1, 1];

            //不包含左上角
            if (myrng.Row > 1)
            {
                rngup = sht.Range[sht.Cells[1, 1], rng.Offset[myrng.Row - 2, 0]];
                rngup.EntireRow.Hidden = true;
            }
            //不包含左上角
            if (myrng.Column > 1)
            {
                rngleft = sht.Range[sht.Cells[1, 1], rng.Offset[0, myrng.Column - 2]];
                rngleft.EntireColumn.Hidden = true;
            }
            //不包含左下角
            if (System.Convert.ToInt32(myrng.Row + myrng.Rows.Count) - 1 < 1048576)
            {
                rngdown = sht.Range[sht.Cells[1048576, 1], rng.Offset[myrng.Row + myrng.Rows.Count - 1, 0]];
                rngdown.EntireRow.Hidden = true;
            }
            //不包含右上角
            if (System.Convert.ToInt32(myrng.Column + myrng.Rows.Column) - 1 < 16384)
            {
                rngright = sht.Range[sht.Cells[1, 16384], rng.Offset[0, myrng.Column + myrng.Columns.Count - 1]];
                rngright.EntireColumn.Hidden = true;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }


        /// <summary>
        /// 隐藏选中区域
        /// </summary>
        /// <param name="control"></param>
        public void HideSelectRange(Office.IRibbonControl control)
        {
            Excel.Range rng;
            rng = Globals.ThisAddIn.Application.Selection;
            rng.EntireColumn.Hidden = true;
            rng.EntireRow.Hidden = true;
        }

        /// <summary>
        /// 取消隐藏所有单元格        
        /// </summary>
        /// <param name="control"></param>
        public void CancelHideCells(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.Cells.EntireColumn.Hidden = false;
            Globals.ThisAddIn.Application.Cells.EntireRow.Hidden = false;
        }

        /// <summary>
        /// 合并一列相同且相邻的单元格
        /// 代码思路分析:
        ///1.首先将操作对象限制为活动工作表的已用数据区域与选区的交集部分，从而避免不必要的循环
        ///2.然后从操作区域的第二个单元格开始与上一个单元格执行比较，如相同就执行下一轮比较
        ///3.如果单元格与其上方的单元格的值不同，那么将上面的未合并的区域(值相同的区域)复制到辅助区域中
        ///4.接着合并辅助区域，再将辅助区域的格式复制到待合并的区域中，从而实现只合并单元格不删除数据
        ///5.最后清除辅助区域，执行下一轮循环。
        ///6.本例通过三方面实现代码提速：重置操作区域（Intersect）、关闭提示（DisplayAlerts）、关闭屏幕刷新（ScreenUpdating）
        /// </summary>
        /// <param name="control"></param>
        public void MergeColumn(Office.IRibbonControl control)
        {
            Excel.Worksheet sht = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range rng;
            Excel.Range rg;
            Excel.Range rngs;
            Excel.Range endrng;
            //object Friendlyreminder=  Interaction.MsgBox("请先对需要合并的列区域排序！", MsgBoxStyle.OkCancel, "友情提示");
            //Interaction.MsgBox(Friendlyreminder);
            if (MessageBox.Show(text: "是否已对需要合并的列区域进行排序?", caption: "友情提示", buttons: MessageBoxButtons.YesNo) == DialogResult.No)
            {
                return;
            }
            //如果选区只有一个单元格，那么提示用户，然后退出程序
            if (Globals.ThisAddIn.Application.Selection.Count == 1)
            {
                //Interaction.MsgBox("请选择一个较大区域！", MsgBoxStyle.Information, "友情提示");
                MessageBox.Show(text: "请选择一个较大区域！", caption: "友情提示");
                return;
            }

            //如果选区超过1列则提示用户，然后退出程序
            if (Globals.ThisAddIn.Application.Selection.Columns.Count > 1)
            {
                //Interaction.MsgBox("只对单列数据进行操作！", MsgBoxStyle.OkOnly, "友情提示");
                MessageBox.Show(text: "只对单列数据进行操作！", caption: "友情提示", buttons: MessageBoxButtons.OK);
                return;
            }
            Globals.ThisAddIn.Application.DisplayAlerts = false; //禁止提示
            Globals.ThisAddIn.Application.ScreenUpdating = false; //禁止屏幕更新
            //提取选区与已用数据区域的交集并赋值给变量，从而避免执行不必要的循环，降低效率
            rngs = Globals.ThisAddIn.Application.Intersect(sht.UsedRange, Globals.ThisAddIn.Application.Selection);
            rg = rngs.Cells[1]; //将区域rngs的第一个单元格赋值给变量rg
            endrng = Globals.ThisAddIn.Application.Cells[1, Globals.ThisAddIn.Application.Cells.Columns.Count];
            //在选区向下偏移一行的区域中循环
            foreach (Excel.Range tempLoopVar_rng in rngs.Offset[1, 0].Resize[rngs.Count, 1])
            {
                rng = tempLoopVar_rng;
                if (rng.Value2 != rng.Offset[-1, 0].Value2) //如果单元格rng与其上一个单元不相等
                {
                    //在工作表最右一列创建一个辅助区，区域的高度等于需要合并的区域的高度，宽度为1
                    dynamic with_1 = endrng.Resize[Globals.ThisAddIn.Application.Range[rg, rng.Offset[-1, 0]].Rows.Count, 1];
                    Globals.ThisAddIn.Application.Range[rg, rng.Offset[-1, 0]].Copy(endrng); //将需要合并的区域复制到辅助区域中
                    with_1.Merge(); //合并辅助区域
                    with_1.Copy(); //
                    Globals.ThisAddIn.Application.Range[rg, rng.Offset[-1, 0]].PasteSpecial(Excel.XlPasteType.xlPasteFormats); //将辅助区域的格式粘贴到需要合并的区域
                    with_1.Clear(); //清除辅助区域
                    rg = rng; //然后重新指定对象变量(rg更新为rng所代表的单元格)
                }
            }
            Globals.ThisAddIn.Application.DisplayAlerts = true; //还原提示
            Globals.ThisAddIn.Application.ScreenUpdating = true; //还原屏幕更新
            rngs[1].Select();
        }

        /// <summary>
        /// 相同项与不同项
        /// </summary>
        /// <param name="control"></param>
        public void SameAndDifferentItem(Office.IRibbonControl control)
        {
            Form frm = new SameAndDifferentItem();
            frm.TopMost = true;
            frm.Show();
            //frm.ShowDialog();
            for (int i = 5; i <= 9; i++)
            {
                frm.Controls["Button" + i].Enabled = false;
            }

        }

        /// <summary>
        /// 反向选择
        /// </summary>
        /// <param name="control"></param>
        public void InvertSelect(Office.IRibbonControl control)
        {
            Excel.Worksheet sht = Globals.ThisAddIn.Application.ActiveSheet;
            //如果选择的对象不是单元格，那么提示用户，然后结束过程
            if (Information.TypeName(Globals.ThisAddIn.Application.Selection) != "Range")
            {
                //Interaction.MsgBox("请选择区域", MsgBoxStyle.OkOnly, "提示");
                MessageBox.Show(text: "请选择区域", caption: "提示", buttons: MessageBoxButtons.OK);
                return;
            }
            Excel.Range rng = Globals.ThisAddIn.Application.Intersect(Globals.ThisAddIn.Application.Selection, sht.UsedRange);
            // MessageBox.Show(rng.Address);
            // MessageBox.Show(sht.UsedRange.Address);

            //如果选区与当前表已用区域不存在交集，那么结束过程
            if (Globals.ThisAddIn.Application.Intersect(Globals.ThisAddIn.Application.Selection, sht.UsedRange).address == "$A$1")
            {
                //Interaction.MsgBox("请不要选择非数据区域", MsgBoxStyle.Information, "提示");
                MessageBox.Show(text: "请不要选择非数据区域", caption: "提示", buttons: MessageBoxButtons.OK);
                return;
            }

            //如果活动工作表的已用区域等于选区的地址，那么结束过程
            //取交集，防止选择的区域比使用区域要大；
            if (sht.UsedRange.Address == Globals.ThisAddIn.Application.Intersect(Globals.ThisAddIn.Application.Selection, sht.UsedRange).address)
            {
                return;
            }

            Globals.ThisAddIn.Application.ScreenUpdating = false; //关闭屏幕更新
            string SelectionAddress; //声明变量
            string UsedAddres;
            SelectionAddress = System.Convert.ToString(Globals.ThisAddIn.Application.Selection.Address); //记录选区地址
            UsedAddres = System.Convert.ToString(sht.UsedRange.Address); //记录当前表已用区域的地址
            dynamic with_1 = Globals.ThisAddIn.Application.Sheets.Add(); //创建一个辅助工作表 
            with_1.Range[UsedAddres] = 0; //在辅助表中与当前表已用区域一致的区域写入常量“0”
            with_1.Range[SelectionAddress] = ""; //在辅助表中与当前表选区一致的区域清除数据
            //记录剩下的0值的区域地址
            SelectionAddress = System.Convert.ToString(with_1.Range(UsedAddres).SpecialCells(Excel.XlCellType.xlCellTypeConstants, 1).Address);
            Globals.ThisAddIn.Application.DisplayAlerts = false; //关闭提示，避免弹出对话框
            with_1.Delete(); //删除辅助表
            Globals.ThisAddIn.Application.DisplayAlerts = true; //恢复提示
            Globals.ThisAddIn.Application.ScreenUpdating = true; //恢复屏幕更新
            sht.Range[SelectionAddress].Select(); //在当前表中选择变量SelectionAddress所代表的区域(即Selection的反向区域)

        }

        /// <summary>
        /// 一键生成工资条
        /// </summary>
        /// <param name="control"></param>
        public void PaySlip(Office.IRibbonControl control)
        {
            //标题行数可以选择，工资数据默认只有一行。
            //On Error Resume Next VBConversions Warning: On Error Resume Next not supported in C#        

            try
            {
                Excel.Worksheet sht = Globals.ThisAddIn.Application.ActiveSheet; ;
                string shtname;
                byte btcount; //标题行数
                shtname = System.Convert.ToString(sht.Name);
                btcount = System.Convert.ToByte(Globals.ThisAddIn.Application.InputBox(Prompt: "请输入标题共几行", Default: 1, Type: 1));
                if (btcount == 0)    //点击了取消 就是0
                {
                    return;

                }

                Globals.ThisAddIn.Application.ScreenUpdating = false; //关闭屏幕刷新，从而加快代码执行速度
                //将活动工作表复制一份，在副本上生成工资条，从而不至于影响工资明细表

                sht.Copy(After: Globals.ThisAddIn.Application.Sheets[Globals.ThisAddIn.Application.Sheets.Count]);
                sht = Globals.ThisAddIn.Application.Sheets[shtname];
                sht.Select();

                //sht.Activate();

                int Item = 0; //声明Integer型的变量Item，用于循环语句中
                //使用For Next循环，起始值为已用区域的最后一行，终止值为btcount + 2
                for (Item = sht.UsedRange.Rows.Count; Item >= btcount + 2; Item--)
                {
                    //对第Item行插入三行，其中前两行用于存放标题，最后一行作为间隔行(方便裁剪)
                    sht.Cells[Item, 1].Resize(btcount + 1, 1).EntireRow.Insert();
                }
                //声明一个Integer型的变量，用于取代已用数据区域的行数ActiveSheet.UsedRange.Rows.Count
                int RowNum = 0;
                //将已用数据区域的行数赋值给变量RowNum
                RowNum = System.Convert.ToInt32(sht.UsedRange.Rows.Count);
                Excel.Range rng = default(Excel.Range); //声明Range类型的变量
                rng = sht.Rows[(btcount + 3) + ":" + (btcount + 3)]; //首先将第btcount + 3行(标题行+数据行+空行+下一行)赋值给变量rng,它是第一个需要插入标题的行
                for (Item = btcount + 3; Item <= RowNum; Item += btcount + 2) //使用For Next循环，从第btcount + 3行到已用数据区域的最后一行
                {
                    //将变量rng与第Item行合并为一个Range对象，此后Rng变量将包含每一个需要插入标题的行
                    rng = Globals.ThisAddIn.Application.Union(rng, sht.Rows[Item]);
                }
                sht.Rows["1:" + System.Convert.ToString(btcount)].Copy(rng); //将标题行复制到刚插入的空行中
                rng.Offset[-1, 0].Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone; //取消间隔行的边框，从而使工资表更美观
                rng.Offset[-1, 0].RowHeight = 7; //指定间隔行的行高为7
                Globals.ThisAddIn.Application.ScreenUpdating = true; //恢复屏幕刷新
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

        }

        /// <summary>
        /// 将工作簿所有的外部链接转换成值
        /// </summary>
        /// <param name="control"></param>
        public void LinkToValue(Office.IRibbonControl control)
        {
            //giving the error:Unable to cast object of type 'System.Object[*]' to type 'System.Object[]'！
            //This issue is caused by different .net version, before 4.0 a simple cast would work, but in C# 4.0, you will need to cast to object first.
            if (DialogResult.Yes == MessageBox.Show("是否要删除所有的数据链接?", "Links", MessageBoxButtons.YesNo))
            {
                Array links = (Array)((object)Globals.ThisAddIn.Application.ActiveWorkbook.LinkSources(Excel.XlLink.xlExcelLinks));

                if (links != null)
                {
                    for (int i = 1; i <= links.Length; i++)
                    {
                        Globals.ThisAddIn.Application.ActiveWorkbook.BreakLink((string)links.GetValue(i), Excel.XlLinkType.xlLinkTypeExcelLinks);
                    }
                }
            }
        }

        /// <summary>
        /// 中文分词
        /// </summary>
        /// <param name="control"></param>
        public void ChWordSeg(Office.IRibbonControl control)
        {

            //MessageBox.Show(Globals.ThisAddIn.Application.AltStartupPath);
            //MessageBox.Show(Globals.ThisAddIn.Application.DefaultFilePath);//666
            //MessageBox.Show(Globals.ThisAddIn.Application.LibraryPath);
            //MessageBox.Show(Globals.ThisAddIn.Application.NetworkTemplatesPath);
            //MessageBox.Show(Globals.ThisAddIn.Application.Path);
            //MessageBox.Show(Globals.ThisAddIn.Application.PathSeparator);
            //MessageBox.Show(Globals.ThisAddIn.Application.StartupPath);
            //MessageBox.Show(Globals.ThisAddIn.Application.TemplatesPath);
            //MessageBox.Show(Globals.ThisAddIn.Application.UserLibraryPath);

            //string str = this.GetType().Assembly.Location;
            //MessageBox.Show(str);
            // MessageBox.Show(AppDomain.CurrentDomain.BaseDirectory);

            //C#窗口对话框一般分为两种类型：模态类型（modal）与非模态类型（modeless）。
            //所谓模态对话框，就是指除非采取有效的关闭手段，用户的鼠标焦点或者输入光标将一直停留在其上的对话框。form.ShowDialog();
            //非模态对话框则不会强制此种特性，用户可以在当前对话框以及其他窗口间进行切换 form.Show(); 
            Form frm = new TextToPinyinForm();
            frm.TopMost = true;
            frm.Show();
        }

        /// <summary>
        /// 拆分工作簿
        /// </summary>
        /// <param name="control"></param>
        public void SplitWorkbook(Office.IRibbonControl control)
        {
            //On Error Resume Next  '当程序出错时继续执行下一句
            string Pathstr; //声明变量
            string ExcelType;
            long i;
            byte j;
            string ActiveWB;

            dynamic fd = Globals.ThisAddIn.Application.FileDialog[Office.MsoFileDialogType.msoFileDialogFolderPicker]; //创建选择文件夹的对话框     

            //Globals.ThisAddIn.Application.FileDialog[Office.MsoFileDialogType.msoFileDialogFolderPicker].SelectedItems[1];
            //Microsoft.Office.Core.FileDialog fd = Globals.ThisAddIn.Application.get_FileDialog(Microsoft.Office.Core.MsoFileDialogType.msoFileDialogFolderPicker);

            if (fd.show() != 0) //如果在对话框中单击了“确定”
            {
                Pathstr = Convert.ToString(fd.SelectedItems[1]); //将选定的路径赋予变量
            }
            else
            {
                return; //否则结束过程
            }

            //如果不是“\”结尾则添加“\”
            if (Pathstr.Substring(Pathstr.Length - 1, 1) != "\\")
            {
                Pathstr = Pathstr + "\\";
            }

            Globals.ThisAddIn.Application.ScreenUpdating = false; //关闭屏幕更新，提升代码执行速度
            ActiveWB = Convert.ToString(Globals.ThisAddIn.Application.ActiveWorkbook.Name); //记录活动工作簿名称
            for (i = 1; i <= Globals.ThisAddIn.Application.Sheets.Count; i++) //遍历所有工作表
            {
                Globals.ThisAddIn.Application.Sheets[i].copy(); //复制工作表到新工作簿中（忽略参数时表示复制到新工作簿中）
                //另存工作簿另存，文件名称由工作表名称决定。而文件的后缀名则由Excel程序的版本决定，如果小于12则用“xls”作为后缀名，否则用“xlsx”
                string ExcelVersion = Globals.ThisAddIn.Application.Version;
                //Convert.ToInt32 会报错：输入字符串的格式不正确。因为14.0 是小数
                if (Convert.ToDecimal(ExcelVersion) < 12)
                {
                    ExcelType = ".xls";
                }
                else
                {
                    ExcelType = ".xlsx";
                }
                Globals.ThisAddIn.Application.ActiveWorkbook.SaveAs(Filename: Pathstr + Globals.ThisAddIn.Application.Workbooks[ActiveWB].Sheets[i].Name + ExcelType, FileFormat: XlFileFormat.xlWorkbookDefault, CreateBackup: false);
                dynamic with_2 = Globals.ThisAddIn.Application.ActiveSheet.UsedRange; //引用已用区域
                //获取工作簿中的外部链接
                Array aLinks = (Array)((object)Globals.ThisAddIn.Application.ActiveWorkbook.LinkSources(Excel.XlLink.xlExcelLinks));
                if (aLinks != null) //如果有外部链接
                {
                    for (j = 1; j <= (aLinks.Length - 1); j++) //遍历所有外部链接
                    {
                        Globals.ThisAddIn.Application.ActiveWorkbook.BreakLink((string)aLinks.GetValue(j), Excel.XlLinkType.xlLinkTypeExcelLinks); //中断第j个链接
                    }
                }
                Globals.ThisAddIn.Application.ActiveWindow.Close(); //关闭活动窗口

                ((Excel._Workbook)Globals.ThisAddIn.Application.Workbooks[ActiveWB]).Activate(); //激活待拆分的工作簿
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true; //恢复屏幕更新
            System.Diagnostics.Process.Start("explorer.exe", Pathstr);//打开文件夹，查看拆分结果
            //Interaction.Shell("EXPLORER.EXE " + Pathstr, Constants.vbNormalFocus, 0, -1); 
        }

        /// <summary>
        /// 创建文件链接列表
        /// </summary>
        /// <param name="control"></param>
        public void Director(Office.IRibbonControl control)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            list.Clear();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                List<String> abc = Director(dialog.SelectedPath);

                string[] arr = abc.ToArray();
                //MessageBox.Show("666");
                try
                {
                    Excel.Range rng = (Excel.Range)Globals.ThisAddIn.Application.InputBox(Prompt: "请选择存放区域，选择单个单元格即可", Type: 8);
                    rng.Offset[0, 0].Resize[arr.GetUpperBound(0) + 1, 1].Value = Globals.ThisAddIn.Application.WorksheetFunction.Transpose(arr);

                    foreach (Excel.Range temprng in rng.Offset[0, 0].Resize[arr.GetUpperBound(0) + 1, 1])
                    {
                        Excel.Worksheet sht = Globals.ThisAddIn.Application.ActiveSheet;
                        sht.Hyperlinks.Add(temprng, temprng.Value);
                    }

                }
                catch
                {
                    return;
                }
            }
        }
        //装到list里面去
        List<string> list = new List<string>();
        public List<String> Director(string dir)
        {
            DirectoryInfo d = new DirectoryInfo(dir);
            FileSystemInfo[] fsinfos = d.GetFileSystemInfos();
            foreach (FileSystemInfo fsinfo in fsinfos)
            {
                if (fsinfo is DirectoryInfo)     //判断是否为文件夹  
                {
                    Director(fsinfo.FullName);//递归调用  
                }
                else
                {
                    list.Add(fsinfo.FullName);//输出文件的全部路径 
                }
            }
            return list;

        }

        /// <summary>
        /// 选择所有工作表
        /// </summary>
        /// <param name="control"></param>
        public void SelectAllSheet(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.ActiveWorkbook.Sheets.Select();
        }

        /// <summary>
        /// 隐藏选择工作表
        /// </summary>
        /// <param name="control"></param>
        public void HideSelectSheet(Office.IRibbonControl control)
        {
            //MsgBox(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets.Count)
            int count = 0;
            Excel.Worksheet sht;
            foreach (Excel.Worksheet tempLoopVar_sht in Globals.ThisAddIn.Application.ActiveWorkbook.Sheets)
            {
                sht = tempLoopVar_sht;
                if (Convert.ToBoolean(sht.Visible))
                {
                    count++;
                }
            }
            if (Globals.ThisAddIn.Application.ActiveWindow.SelectedSheets.Count < count)
            {
                Globals.ThisAddIn.Application.ActiveWindow.SelectedSheets.Visible = false;
            }
            else
            {
                MessageBox.Show(text: "工作簿内至少含有一张可视工作表", caption: "请注意", buttons: MessageBoxButtons.OKCancel, icon: MessageBoxIcon.Exclamation);
            }
        }

        /// <summary>
        /// 创建工作表链接列表
        /// </summary>
        /// <param name="control"></param>
        public void SheetLink(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false; //关闭屏幕刷新
            //有错误时继续执行下一句(当没有“工作表目录”时下一句代码会出错)
            //On Error Resume Next VBConversions Warning: On Error Resume Next not supported in C#
            Globals.ThisAddIn.Application.DisplayAlerts = false; //关闭提示(删除非空工作表时会弹出提示框)
            try
            {
                Globals.ThisAddIn.Application.Sheets["工作表目录"].Delete(); //删除已有的工作表目录(假设有的话)
            }
            catch
            {
            }
            Globals.ThisAddIn.Application.DisplayAlerts = true; //恢复提示
            //在最前面创建一个新工作表，并命名为“工作表目录”
            Globals.ThisAddIn.Application.Sheets.Add(Globals.ThisAddIn.Application.Sheets[1]).Name = "工作表目录";
            //写入标题，采用数组形式，可以一次性写入多个单元格的值	
            Excel.Worksheet sht = Globals.ThisAddIn.Application.ActiveSheet;
            sht.Range["A1:B1"].Value = new[] { "编号", "目录" };
            for (int i = 2; i <= Globals.ThisAddIn.Application.Worksheets.Count; i++) //遍历工作表目录以外的所有工作表
            {
                sht.Cells[i, 1].Value = i - 1; //在单元格中写入序号
                //在单元格中创建链接
                sht.Hyperlinks.Add(Anchor: sht.Cells[i, 2], Address: "", SubAddress: "'" + Globals.ThisAddIn.Application.Worksheets[i].Name + "'" + "!A1", TextToDisplay: Globals.ThisAddIn.Application.Worksheets[i].Name, ScreenTip: "单击打开：" + Globals.ThisAddIn.Application.Worksheets[i].Name);
            }
            Globals.ThisAddIn.Application.Sheets[sht.Name].Activate();
            sht.Range["A2"].Select(); //选择A2单元格
            Globals.ThisAddIn.Application.ActiveWindow.FreezePanes = true; //冻结窗格(让首行固定，方便查看)
            Globals.ThisAddIn.Application.ScreenUpdating = true; //恢复屏幕刷新
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="control"></param>
        public void oracle_createtablesql(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn.CustomTaskPanes.Contains(OracleCreateTableSqlTaskPane))
            {
                Globals.ThisAddIn.CustomTaskPanes.Remove(OracleCreateTableSqlTaskPane);

            }
            else
            {
                OracleCreateTableSqlTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(new oracle_createtablesql(), "Oracle建表语句生成器");
                OracleCreateTableSqlTaskPane.Width = 415;
                OracleCreateTableSqlTaskPane.Visible = true;
            }
        }

        public void createtjson(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn.CustomTaskPanes.Contains(CreateJsonTaskPane))
            {
                Globals.ThisAddIn.CustomTaskPanes.Remove(CreateJsonTaskPane);

            }
            else
            {
                CreateJsonTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(new CreateJson(), "JSON生成器");
                CreateJsonTaskPane.Visible = true;
            }
        }


        public void oraclecomment(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn.CustomTaskPanes.Contains(oraclecommentTaskPane))
            {
                Globals.ThisAddIn.CustomTaskPanes.Remove(oraclecommentTaskPane);

            }
            else
            {
                oraclecommentTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(new Oraclecomment(), "oracle表和字段注释");
                oraclecommentTaskPane.Width = 415;
                oraclecommentTaskPane.Visible = true;
            }
        }

        public void jsontoarray(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn.CustomTaskPanes.Contains(jsontoarrayTaskPane))
            {
                Globals.ThisAddIn.CustomTaskPanes.Remove(jsontoarrayTaskPane);

            }
            else
            {
                jsontoarrayTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(new jsontoarray(), "接口的JSON转数组");
                jsontoarrayTaskPane.Width = 415;
                jsontoarrayTaskPane.Visible = true;
            }
        }


        public void GeneralBasicDemo(Office.IRibbonControl control)
        {
            BaiduAI baidu_ai = new BaiduAI();
            baidu_ai.GeneralBasicDemo();
        }

        public void AccurateBasicDemo(Office.IRibbonControl control)
        {
            BaiduAI baidu_ai = new BaiduAI();
            baidu_ai.AccurateBasicDemo();
        }
        
        public void TableRecognitionGetResultDemo(Office.IRibbonControl control)
        {            
            BaiduAI baidu_ai = new BaiduAI();
            baidu_ai.TableRecognitionGetResultDemo();
        }





        public void Help(Office.IRibbonControl control)
        {
            MessageBox.Show(text: "有疑问请联系蝌蚪文数据工作室，QQ群：208882120", caption: "蝌蚪文数据工具箱");
        }

        public string Group2getlabel(Office.IRibbonControl control)
        {
            return "今天是" + StrLabel + "【单元格与区域】";
        }


        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("kedouwenc.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1


        private Timer timer = new Timer();
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ribbonUI"></param>
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            timer.Interval = 1000;// 一秒=1000ms,以ms为单位的
            this.ribbon = ribbonUI;
            timer.Tick += new EventHandler(timer_Tick);
            timer.Start();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void timer_Tick(object sender, EventArgs e)
        {
            StrLabel = DateTime.Now.ToString();
            ribbon.InvalidateControl("Group2");
            //https://msdn.microsoft.com/en-us/library/aa433553(v=office.12).aspx
            //Invalidates the cached value for a single control on the Ribbon user interface.
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}


