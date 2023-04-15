using System;
using System.IO;
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
            Class_WindowDeactivate_Event iClass_WindowDeactivate_Event = new Class_WindowDeactivate_Event(); iClass_WindowDeactivate_Event.Evento_WindowDeactivate();
        }

        private void MyCustomTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.TaskPane1.Visible == false & Globals.Ribbons.Ribbon1.toggleButton_estilos.Checked == true) Globals.Ribbons.Ribbon1.toggleButton_estilos.Checked = false;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {            
            // Atualiza preferências
            //string preferences_path = Path.Combine(Ribbon1.Variables.caminho_AppData_Roaming_PeriTAB, "preferences.txt");

            //File.WriteAllText(preferences_path, "teste1");

            //try
            //{
            //    Globals.ThisAddIn.Application.ActiveDocument.SaveAs2(tmpsave);
            //}
            //catch
            //{
            //    System.IO.Directory.CreateDirectory(Variables.caminho_tmp);
            //    Globals.ThisAddIn.Application.ActiveDocument.SaveAs2("tmpsave.docx");
            //}


            //            // Create a file to write to.
            //string createText = "Hello and Welcome" + Environment.NewLine;
            //File.WriteAllText(path, createText);

            //...

            //// Open the file to read from.
            //string readText = File.ReadAllText(path);
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
