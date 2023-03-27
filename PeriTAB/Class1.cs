using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace PeriTAB
{
    public partial class Class1
    {
        public static void Metodo_DocumentOpen(Microsoft.Office.Interop.Word.Document Doc)
        {
            MessageBox.Show("DocumentOpen");            
        }

        internal static void Metodo_DocumentOpen()
        {
            //throw new NotImplementedException();
            MessageBox.Show("DocumentOpen2");
        }

        public void Evento_DocumentOpen()
        {
            Globals.ThisAddIn.Application.DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(Metodo_DocumentOpen);
        }

    }
}
