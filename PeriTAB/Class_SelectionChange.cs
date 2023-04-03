using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PeriTAB
{
    internal class Class_SelectionChange
    {
        public void Evento_WindowSelectionChange()
        {
            Globals.ThisAddIn.Application.WindowSelectionChange += new ApplicationEvents4_WindowSelectionChangeEventHandler(Metodo_WindowSelectionChange);
        }

        private UserControl a;

        private void Metodo_WindowSelectionChange(Selection Sel)
        {
            Globals.Ribbons.Ribbon1.button1.Enabled = false;
            Globals.ThisAddIn.myUserControl1.Metodo_button1(false);
        }


    }
}
