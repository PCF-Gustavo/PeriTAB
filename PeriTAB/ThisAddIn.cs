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
using Microsoft.Office.Tools.Ribbon;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using System.Windows;

namespace PeriTAB
{
    public partial class ThisAddIn
    {
        public UserControl1 myUserControl1;
        public Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Configura Task Pane
            myUserControl1 = new UserControl1();
            myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl1, "Estilos (PeriTAB)");
            myCustomTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
            myCustomTaskPane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
            myCustomTaskPane.Height = 90;
            myCustomTaskPane.VisibleChanged += MyCustomTaskPane_VisibleChanged;
        }

        private void MyCustomTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.myCustomTaskPane.Visible == false & Globals.Ribbons.Ribbon1.toggleButton1.Checked == true) Globals.Ribbons.Ribbon1.toggleButton1.Checked = false;
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
