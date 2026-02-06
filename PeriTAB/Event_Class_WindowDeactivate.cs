using Microsoft.Office.Interop.Word;
using System.Linq;

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
            Clean_Windows();
        }

        private void Clean_Windows()
        {
            foreach (Window window in Globals.ThisAddIn.Dicionario_Window_e_UserControl.Keys.ToList())
            {
                try
                {
                    _ = window.Caption;
                }
                catch
                {
                    //Globals.ThisAddIn.Dicionario_Window_e_Doc.Remove(window);
                    Globals.ThisAddIn.Dicionario_Window_e_UserControl.Remove(window);
                    //Globals.ThisAddIn.Dicionario_Window_e_TaskPane.Remove(window);
                }
            }
        }

        }
}
