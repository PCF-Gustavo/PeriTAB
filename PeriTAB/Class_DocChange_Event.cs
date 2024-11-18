using Microsoft.Office.Interop.Word;

namespace PeriTAB
{
    internal class Class_DocChange_Event
    {
        public void Evento_DocChange()
        {
            Globals.ThisAddIn.Application.DocumentChange += new ApplicationEvents4_DocumentChangeEventHandler(Metodo_DocChange);
        }
        private void Metodo_DocChange()
        {
            //MessageBox.Show("docchange");
        }
    }
}
