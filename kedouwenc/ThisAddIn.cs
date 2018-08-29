using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms; 

//ThisAddIn 类的分部定义。 此类提供了代码的入口点，并提供了对 Excel 对象模型的访问。 有关更多信息，请参见应用程序级外接程序编程。
//ThisAddIn 类的其余部分是在一个隐藏的代码文件中定义的，您不应修改此代码文件。

//this.Application.WorkbookBeforeSave += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
//如果您使用的是 C#，请将以下必需代码添加到 ThisAddIn_Startup 事件处理程序中。 此代码用于将 Application_WorkbookBeforeSave 事件处理程序与 WorkbookBeforeSave 事件连接起来。

namespace kedouwenc
{
    public partial class ThisAddIn
    {
        

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
           Globals.ThisAddIn.Application.SheetSelectionChange += new Excel.AppEvents_SheetSelectionChangeEventHandler(Application_SheetSelectionChange);
            cellmeau cell_meau = new cellmeau();
            cell_meau.cellmenu();
        }

        void Application_SheetSelectionChange(object Sh, Excel.Range Target)
        {
           // MessageBox.Show(Convert.ToString(Ribbon1.ispressed));

            if (Ribbon1.ispressed)
            {
                Globals.ThisAddIn.Application.EnableEvents = false; //禁止响应事件
                if (Target.Rows.Count < 3 && Target.Columns.Count < 3)
                {
                    Globals.ThisAddIn.Application.Union(Target.EntireColumn, Target.EntireRow).Select();
                    Target.Activate();
                }
                Globals.ThisAddIn.Application.EnableEvents = true; //恢复响应事件          
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
