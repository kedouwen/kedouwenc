using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace kedouwenc
{
    class cellmeau
    {

        private Office.CommandBarPopup cell_meau;
        private Office.CommandBarButton cell_meau_btn_upper;
        private Office.CommandBarButton cell_meau_btn_lower;
        public void cellmenu()
        {

            try
            {
                Globals.ThisAddIn.Application.CommandBars["Cell"].Controls["转成大小写"].Delete();
            }
            catch
            {

                cell_meau = (Office.CommandBarPopup)Globals.ThisAddIn.Application.CommandBars["Cell"].Controls.Add(Type: Office.MsoControlType.msoControlPopup, Before: 1);
                cell_meau.Caption = "转成大小写";
                

                cell_meau_btn_upper = (Office.CommandBarButton)cell_meau.Controls.Add(Type: Office.MsoControlType.msoControlButton);
                cell_meau_btn_upper.Caption = "转成大写";
                cell_meau_btn_upper.FaceId = 80;
                cell_meau_btn_upper.Click += cell_meau_btn_upper_Click;

                cell_meau_btn_lower = (Office.CommandBarButton)cell_meau.Controls.Add(Type: Office.MsoControlType.msoControlButton);
                cell_meau_btn_lower.Caption = "转成小写";
                cell_meau_btn_lower.FaceId = 81;
                cell_meau_btn_lower.Click += Cell_meau_btn_lower_Click;             

            }
        }

        private void Cell_meau_btn_lower_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            //throw new NotImplementedException();
            LowerCase();
        }

        private void cell_meau_btn_upper_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            //throw new NotImplementedException();           
            UpperCase();
        }

        private void UpperCase()
        {
            foreach (Excel.Range rng in Globals.ThisAddIn.Application.Selection)
            {
                string str = Convert.ToString(rng.Value2);
                rng.Value2 = str.ToUpper();
            }


        }

        private void LowerCase()
        {
            foreach (Excel.Range rng in Globals.ThisAddIn.Application.Selection)
            {
                string str = Convert.ToString(rng.Value2);
                rng.Value2 = str.ToLower();
            }


        }










    }
}
