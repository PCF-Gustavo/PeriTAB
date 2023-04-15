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
    internal class Class_WindowDeactivate_Event
    {
        public void Evento_WindowDeactivate()
        {            
            Globals.ThisAddIn.Application.WindowDeactivate += new ApplicationEvents4_WindowDeactivateEventHandler(Metodo_WindowDeactivate);
        }
        private void Metodo_WindowDeactivate(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn)
        {

        }
    }
}
