using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace PeriTAB
{
    internal class Class_DocumentBeforeClose_Event
    {
        public void Evento_DocumentBeforeClose()
        {            
            Globals.ThisAddIn.Application.DocumentBeforeClose += new ApplicationEvents4_DocumentBeforeCloseEventHandler(Metodo_DocumentBeforeClose);
        }
        private void Metodo_DocumentBeforeClose(Document Doc, ref bool Cancel)
        {
            //Class_New_or_Open_Event.list_TaskPane.Remove(Class_New_or_Open_Event.TaskPane_1);
            //bool b = Globals.ThisAddIn.CustomTaskPanes.Remove(Class_New_or_Open_Event.TaskPane_1);
            try
            {
                Microsoft.Office.Tools.CustomTaskPane CTP = Class_New_or_Open_Event.Dicionario_Doc_e_TaskPane[Doc];
                //bool b = Globals.ThisAddIn.CustomTaskPanes.Remove(CTP);
                Globals.ThisAddIn.CustomTaskPanes.Remove(CTP);
                Class_New_or_Open_Event.Dicionario_Doc_e_TaskPane.Remove(Doc);
                Globals.ThisAddIn.Dicionario_Doc_e_UserControl.Remove(Doc);
            }
            catch (System.Collections.Generic.KeyNotFoundException ex) { }
            //if (b) MessageBox.Show("Taskpane removed");
        }
    }
}