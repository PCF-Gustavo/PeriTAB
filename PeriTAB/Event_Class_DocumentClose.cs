using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Threading;
using Tarefa = System.Threading.Tasks.Task;

namespace PeriTAB
{
    internal class Class_DocumentClose_Event
    {
        private List<Microsoft.Office.Interop.Word.Document> List_docs_esperando_fechar = new List<Microsoft.Office.Interop.Word.Document>();

        private static int private_OpenDocumentNumber;
        public static int OpenDocumentNumber { get { return private_OpenDocumentNumber; } set { private_OpenDocumentNumber = value; } }
        public void Evento_DocumentClose()
        {
            Globals.ThisAddIn.Application.DocumentBeforeClose += new ApplicationEvents4_DocumentBeforeCloseEventHandler(Metodo_DocumentBeforeClose);
        }
        private void Metodo_DocumentBeforeClose(Document Doc, ref bool Cancel)
        {
            if (!List_docs_esperando_fechar.Contains(Doc))
            {
                List_docs_esperando_fechar.Add(Doc);
                espera_fechar(Doc); //Cria a thread  que chama o Metodo_DocumentClose()
            }
        }

        private void espera_fechar(Document Doc)
        {
            Tarefa.Run(() =>
            {
                while (true)
                {
                    if (Globals.ThisAddIn.Application.Documents.Count != OpenDocumentNumber) break;
                    Thread.Sleep(1000);
                }
                Metodo_DocumentAfterClose(Doc);
            });
        }

        public void Tracking_OpenDocumentNumber()
        {
            OpenDocumentNumber = Globals.ThisAddIn.Application.Documents.Count;
        }
        public void Metodo_DocumentAfterClose(Document Doc)
        {
            Tarefa.Run(() =>
            {
                Thread.Sleep(100);
                if (IsDocumentOpen(Doc))
                {
                    return;
                }
                Tracking_OpenDocumentNumber();

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
            });
        }

        public bool IsDocumentOpen(Document Doc)
        {
            foreach (Document document in Globals.ThisAddIn.Application.Documents)
            {
                if (Document.Equals(Doc, document))
                {
                    return true; // Documento encontrado, está aberto
                }
            }
            return false; // Documento não encontrado, está fechado
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