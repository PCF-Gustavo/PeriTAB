using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System.Windows;

namespace PeriTAB
{
    public partial class ThisAddIn
    {
        public UserControl1 iUserControl1;
        public Microsoft.Office.Tools.CustomTaskPane TaskPane1;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Configura Task Pane
            iUserControl1 = new UserControl1();
            TaskPane1 = this.CustomTaskPanes.Add(iUserControl1, "Estilos (PeriTAB)");
            TaskPane1.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
            TaskPane1.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
            TaskPane1.Height = 90;
            TaskPane1.VisibleChanged += MyCustomTaskPane_VisibleChanged;

            Class_AnyButtonClick_Event iClass_AnyButtonClick_Event = new Class_AnyButtonClick_Event(); iClass_AnyButtonClick_Event.Evento_AnyButtonClick();
        }

        private void MyCustomTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.TaskPane1.Visible == false & Globals.Ribbons.Ribbon1.toggleButton1.Checked == true) Globals.Ribbons.Ribbon1.toggleButton1.Checked = false;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Código gerado por VSTO

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
