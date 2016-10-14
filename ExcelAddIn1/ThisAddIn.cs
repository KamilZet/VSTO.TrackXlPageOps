using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {  

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.SheetBeforeDelete += NofityPageDeletion;
            this.Application.WorkbookNewSheet += NotifyPageAdded;
            this.Application.WorkbookBeforeClose += Application_CancelBookClose;
        }


        private void Application_CancelBookClose(Excel.Workbook Wb, ref bool cancel)
        {
            
            if (MessageBox.Show("Are you sure!","close",MessageBoxButtons.YesNo) == DialogResult.Yes)
                cancel = true;       
        }

        private void NofityPageDeletion(object sh)
        {
            Globals.Ribbons.Ribbon1.LogPagesOperations(
                "Deleted page: " + ((Excel.Worksheet)sh).Name);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        private void NotifyPageAdded(Excel.Workbook w,object sh)
        {
            Globals.Ribbons.Ribbon1.LogPagesOperations(
                "Added page: " + ((Excel.Worksheet)sh).Name);                
                
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
