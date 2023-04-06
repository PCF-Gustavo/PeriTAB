using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PeriTAB
{
    internal class Class_DocSave_Event
    {
        public void Evento_DocSave()
        {
            Globals.ThisAddIn.Application.DocumentBeforeSave += new ApplicationEvents4_DocumentBeforeSaveEventHandler(Metodo_DocSave);
        }
        private void Metodo_DocSave(Microsoft.Office.Interop.Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            Class_Buttons iClass_Buttons = new Class_Buttons(); iClass_Buttons.Default();
            if (Globals.ThisAddIn.Application.ActiveDocument.Path == "") { Globals.Ribbons.Ribbon1.button3.Enabled = false; Globals.Ribbons.Ribbon1.button3.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon1.button3.SuperTip = "Este documento ainda não foi salvo."; }
            if (Globals.ThisAddIn.Application.ActiveDocument.FullName == "http") { Globals.Ribbons.Ribbon1.button3.Enabled = false; Globals.Ribbons.Ribbon1.button3.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon1.button3.SuperTip = "Este documento não pode ser renomeado porque está salvo online."; }
        }
    }
}
