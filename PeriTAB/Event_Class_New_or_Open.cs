using Microsoft.Office.Interop.Word;
using System.Collections.Generic;

namespace PeriTAB
{
    public class Class_New_or_Open_Event
    {
        Class_DocumentClose_Event iClass_DocumentClose_Event = new Class_DocumentClose_Event();
        Class_AnyButtonClick_Event iClass_AnyButtonClick_Event = new Class_AnyButtonClick_Event();
        Class_ContentControlOnExit_Event iClass_ContentControlOnExit_Event = new Class_ContentControlOnExit_Event();

        Class_CustomTaskPanes iClass_CustomTaskPanes = new Class_CustomTaskPanes();

        public static Microsoft.Office.Tools.CustomTaskPane iTaskPane;
        public static Dictionary<Microsoft.Office.Interop.Word.Document, Microsoft.Office.Tools.CustomTaskPane> Dicionario_Doc_e_TaskPane = new Dictionary<Microsoft.Office.Interop.Word.Document, Microsoft.Office.Tools.CustomTaskPane>();

        public void Evento_New_or_Open()
        {
            ((Microsoft.Office.Interop.Word.ApplicationEvents4_Event)Globals.ThisAddIn.Application).NewDocument += new ApplicationEvents4_NewDocumentEventHandler(Metodo_New_or_Open);
            Globals.ThisAddIn.Application.DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(Metodo_New_or_Open);
        }
        public void Metodo_New_or_Open(Microsoft.Office.Interop.Word.Document Doc)
        {
            //System.Windows.Forms.MessageBox.Show("new or open");
            iClass_DocumentClose_Event.Tracking_OpenDocumentNumber();

            iClass_ContentControlOnExit_Event.Metodo_ContentControlOnExit();

            // Cria um novo UserControl e um novo CustomTaskPane para cada documento aberto
            if (!Globals.ThisAddIn.Dicionario_Doc_e_UserControl.ContainsKey(Doc))
            {
                //************** precisa criar 1 usercontrol para cada documento aberto????? **************
                Globals.ThisAddIn.iMyUserControl = new MyUserControl();
                Globals.ThisAddIn.iMyUserControl.AutoScroll = true;

                iClass_AnyButtonClick_Event.Evento_AnyButtonClick(Globals.ThisAddIn.iMyUserControl);

                iTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(Globals.ThisAddIn.iMyUserControl, "Painel de Estilos (PeriTAB)");
                Globals.ThisAddIn.Dicionario_Doc_e_UserControl.Add(Doc, Globals.ThisAddIn.iMyUserControl);
                Dicionario_Doc_e_TaskPane.Add(Doc, iTaskPane);
                iTaskPane.VisibleChanged += iClass_CustomTaskPanes.MyCustomTaskPane_VisibleChanged;
                iClass_CustomTaskPanes.Redimensionar(Globals.ThisAddIn.iMyUserControl, iTaskPane);
                if (Globals.Ribbons.Ribbon.toggleButton_painel_de_estilos.Checked) iTaskPane.Visible = true; //Checa se deve mostrar o "Painel de Estilos" do Ribbon
            }
        }
    }
}
