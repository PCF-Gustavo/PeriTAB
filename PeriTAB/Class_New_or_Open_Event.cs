using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PeriTAB
{
    internal class Class_New_or_Open_Event
    {
        public void Evento_New_or_Open()
        {
            ((Microsoft.Office.Interop.Word.ApplicationEvents4_Event)Globals.ThisAddIn.Application).NewDocument += new ApplicationEvents4_NewDocumentEventHandler(Metodo_New_or_Open);
            Globals.ThisAddIn.Application.DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(Metodo_New_or_Open);
        }
        private void Metodo_New_or_Open(Microsoft.Office.Interop.Word.Document Doc)
        {
            Globals.ThisAddIn.iUserControl1.Habilita_button9(true);
            if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count > 1 | Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListType == 0 | Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListValue == 1) { Globals.ThisAddIn.iUserControl1.Habilita_button9(false); }

        }
    }
}
