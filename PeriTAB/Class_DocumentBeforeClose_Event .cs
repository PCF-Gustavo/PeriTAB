using Microsoft.Office.Interop.Word;
using System.Windows.Forms;

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
            //MessageBox.Show("Evento_DocumentBeforeClose");
            //try
            //{
            //    Microsoft.Office.Tools.CustomTaskPane CTP = Class_New_or_Open_Event.Dicionario_Doc_e_TaskPane[Doc];
            //    Globals.ThisAddIn.CustomTaskPanes.Remove(CTP);
            //    Class_New_or_Open_Event.Dicionario_Doc_e_TaskPane.Remove(Doc);
            //    Globals.ThisAddIn.Dicionario_Doc_e_UserControl.Remove(Doc);
            //}
            //catch (System.Collections.Generic.KeyNotFoundException) { }

            //*********************************************
            // Exclusão do Painel de Estilos. Remove também dos dicionarios.
            if (Class_New_or_Open_Event.Dicionario_Doc_e_TaskPane.ContainsKey(Doc))
            {
                Globals.ThisAddIn.CustomTaskPanes.Remove(Class_New_or_Open_Event.Dicionario_Doc_e_TaskPane[Doc]);
                Class_New_or_Open_Event.Dicionario_Doc_e_TaskPane.Remove(Doc);
                Globals.ThisAddIn.Dicionario_Doc_e_UserControl.Remove(Doc);
            }

            if (Globals.ThisAddIn.Dicionario_Doc_e_UserControl.ContainsKey(Doc))
            {
                Globals.ThisAddIn.Dicionario_Doc_e_UserControl.Remove(Doc);
            }

            //Monitoramento do Painel de estilos
            if (Ribbon.Variables.debugging)
            {
                string string_Documents_Count = (Globals.ThisAddIn.Application.Documents.Count - 1).ToString();
                string string_CustomTaskPanes_Count = Globals.ThisAddIn.CustomTaskPanes.Count.ToString();
                string string_Dicionario_Doc_e_UserControl_Count = Globals.ThisAddIn.Dicionario_Doc_e_UserControl.Count.ToString();
                string string_Dicionario_Doc_e_TaskPane_Count = Class_New_or_Open_Event.Dicionario_Doc_e_TaskPane.Count.ToString();

                if (!(string_Documents_Count == string_CustomTaskPanes_Count && string_CustomTaskPanes_Count == string_Dicionario_Doc_e_UserControl_Count && string_Dicionario_Doc_e_UserControl_Count == string_Dicionario_Doc_e_TaskPane_Count))
                {
                    foreach (var taskPane in Globals.ThisAddIn.CustomTaskPanes)
                    {
                        string doc_name = ObterNomeDocumentoPorTaskPane(taskPane);
                        if (!(taskPane.Control is UserControl userControl))
                        {
                            MessageBox.Show("O taskPane do documento " + doc_name + " não possui um UserControl associado.");
                        }
                    }
                    MessageBox.Show("\nDocuments.Count: " + string_Documents_Count + "\nCustomTaskPanes_Count: " + string_CustomTaskPanes_Count + "\nDicionario_Doc_e_UserControl.Count: " + string_Dicionario_Doc_e_UserControl_Count + "\nDicionario_Doc_e_TaskPane.Count: " + string_Dicionario_Doc_e_TaskPane_Count);
                }
            }
        }

        public string ObterNomeDocumentoPorTaskPane(Microsoft.Office.Tools.CustomTaskPane taskPane)
        {
            foreach (var entry in Class_New_or_Open_Event.Dicionario_Doc_e_TaskPane)
            {
                // Verifica se o TaskPane corresponde ao valor no dicionário
                if (entry.Value == taskPane)
                {
                    // Retorna o nome do documento associado
                    return entry.Key.Name;  // entry.Key é o documento associado
                }
            }
            return null;  // Retorna null se o TaskPane não estiver no dicionário
        }
    }
}