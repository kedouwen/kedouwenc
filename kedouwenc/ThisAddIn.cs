using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using System.Drawing;

//ThisAddIn 类的分部定义。 此类提供了代码的入口点，并提供了对 Excel 对象模型的访问。 有关更多信息，请参见应用程序级外接程序编程。
//ThisAddIn 类的其余部分是在一个隐藏的代码文件中定义的，您不应修改此代码文件。

//this.Application.WorkbookBeforeSave += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
//如果您使用的是 C#，请将以下必需代码添加到 ThisAddIn_Startup 事件处理程序中。 此代码用于将 Application_WorkbookBeforeSave 事件处理程序与 WorkbookBeforeSave 事件连接起来。

namespace kedouwenc
{
    public partial class ThisAddIn
    {

        static Excel.Range previousSpotLightRange;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Globals.ThisAddIn.Application.SheetSelectionChange += new Excel.AppEvents_SheetSelectionChangeEventHandler(Application_SheetSelectionChange);
            Globals.ThisAddIn.Application.WindowResize += Application_WindowResize;
            cellmeau cell_meau = new cellmeau();
            cell_meau.cellmenu();

        }

        private void Application_WindowResize(Excel.Workbook Wb, Excel.Window Wn)
        {
            if (Ribbon1.isnewpressed)
            {
                Ribbon1 xlribbon = new Ribbon1();
                xlribbon.LightShine();
            }
        }

        void Application_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            // MessageBox.Show(Convert.ToString(Ribbon1.ispressed));
            if (Ribbon1.isnewpressed)
            {
                Ribbon1 xlribbon = new Ribbon1();
                xlribbon.LightShine();
            }
        }

        //void Application_SheetSelectionChange(object Sh, Excel.Range Target)
        //{
        //    // MessageBox.Show(Convert.ToString(Ribbon1.ispressed));
        //    if (Ribbon1.ispressed)
        //    {
        //        Globals.ThisAddIn.Application.EnableEvents = false; //禁止响应事件
        //        //if (Target.Rows.Count < 3 && Target.Columns.Count < 3)
        //        //{
        //        //    Globals.ThisAddIn.Application.Union(Target.EntireColumn, Target.EntireRow).Select();
        //        //    Target.Activate();
        //        //}


        //        Excel.Worksheet sht = Target.Parent;                
        //        int colIndex = Target.Column;
        //        int rowIndex = Target.Row;
        //        DeletePreviouSpotLightCondition();

        //        Excel.Range spotLightRange = Globals.ThisAddIn.Application.Union(sht.Range[sht.Cells[1, colIndex], Target.Offset[-1, 0].Resize[1, Target.Columns.Count]],
        //                                                           sht.Range[Target.Offset[Target.Rows.Count, 0], sht.Cells[sht.Rows.Count, colIndex]],
        //                                                           sht.Range[sht.Cells[rowIndex, 1], Target.Offset[0, -1].Resize[Target.Rows.Count, 1]],
        //                                                           sht.Range[Target.Offset[0, Target.Columns.Count], sht.Cells[rowIndex, sht.Columns.Count]]
        //                                                            );

        //        Excel.FormatCondition currentFormatCondition = spotLightRange.FormatConditions.Add(Type: Excel.XlFormatConditionType.xlExpression, Formula1: "=TRUE");
        //        currentFormatCondition.SetFirstPriority();
        //        currentFormatCondition.StopIfTrue = true;


        //        Color spotColor = Color.LightGoldenrodYellow;


        //        currentFormatCondition.Interior.Color = spotColor;
        //        //currentFormatCondition.Interior.Color = ColorTranslator.ToOle(spotColor);

        //        previousSpotLightRange = spotLightRange;
        //        Globals.ThisAddIn.Application.EnableEvents = true; //恢复响应事件          
        //    }
        //}

        public static void DeletePreviouSpotLightCondition()
        {
            if (previousSpotLightRange != null)
            {
                Excel.FormatCondition previousFormatCondition = previousSpotLightRange.FormatConditions[1];
                if (previousFormatCondition.Formula1 == "=TRUE")
                {
                    previousFormatCondition.Delete();
                }
                previousSpotLightRange = null;
            }
        }



        //Calling Code in an Application-Level Add-in from VBA
        private AddInUtilities utilities;
        protected override object RequestComAddInAutomationService()
        {
            if (utilities == null)
                utilities = new AddInUtilities();

            return utilities;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion














    }
}
