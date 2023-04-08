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
            //Configura o Task Pane
            iUserControl1 = new UserControl1();
            TaskPane1 = Globals.ThisAddIn.CustomTaskPanes.Add(iUserControl1, "Estilos (PeriTAB)");
            TaskPane1.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
            TaskPane1.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
            TaskPane1.Height = 90;
            TaskPane1.VisibleChanged += MyCustomTaskPane_VisibleChanged;

            //Inicia Eventos
            Class_New_or_Open_Event iClass_New_or_Open_Event = new Class_New_or_Open_Event(); iClass_New_or_Open_Event.Evento_New_or_Open();
            Class_AnyButtonClick_Event iClass_AnyButtonClick_Event = new Class_AnyButtonClick_Event(); iClass_AnyButtonClick_Event.Evento_AnyButtonClick();
            Class_Buttons iClass_Buttons = new Class_Buttons(); iClass_Buttons.DefaultAll();
            Class_DocChange_Event iClass_DocChange_Event = new Class_DocChange_Event(); iClass_DocChange_Event.Evento_DocChange();
            Class_DocSave_Event iClass_DocSave_Event = new Class_DocSave_Event(); iClass_DocSave_Event.Evento_DocSave();            
            Class_SelectionChange_Event iClass_SelectionChange_Event = new Class_SelectionChange_Event(); iClass_SelectionChange_Event.Evento_SelectionChange();
            Class_WindowActivate_Event iClass_WindowActivate_Event = new Class_WindowActivate_Event(); iClass_WindowActivate_Event.Evento_WindowActivate();
        }

        private void MyCustomTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.TaskPane1.Visible == false & Globals.Ribbons.Ribbon1.toggleButton_estilos.Checked == true) Globals.Ribbons.Ribbon1.toggleButton_estilos.Checked = false;
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
