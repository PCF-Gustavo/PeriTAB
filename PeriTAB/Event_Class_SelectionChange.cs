using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Tarefa = System.Threading.Tasks.Task;

namespace PeriTAB
{
    internal class Class_SelectionChange_Event
    {
        private static CancellationTokenSource cancellationTokenSource = null;

        public void Evento_SelectionChange()
        {
            Globals.ThisAddIn.Application.WindowSelectionChange += new ApplicationEvents4_WindowSelectionChangeEventHandler(Metodo_SelectionChange);
        }

        private async void Metodo_SelectionChange(Selection Sel)
        {
            // Se houver uma operação em andamento, cancelamos a execução anterior
            cancellationTokenSource?.Cancel();



            //if (cancellationTokenSource != null)
            //cancellationTokenSource.Cancel(); // Cancela a execução anterior

            // Criamos uma nova fonte de cancelamento para a próxima execução
            cancellationTokenSource = new CancellationTokenSource();
            CancellationToken token = cancellationTokenSource.Token;

            //Declara instacias das classes
            MyUserControl UserControl_ActiveDocument = Globals.ThisAddIn.Dicionario_Doc_e_UserControl[Globals.ThisAddIn.Application.ActiveDocument];

            //Revisa a habilitação do CheckBox "Destacar campos" do Ribbon
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)1) { Globals.Ribbons.Ribbon.checkBox_destaca_campos.Checked = true; }
                if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)0 | Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)2) { Globals.Ribbons.Ribbon.checkBox_destaca_campos.Checked = false; }

                //Revisa a habilitação do CheckBox "Mostrar indicadores" do Ribbon
                if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowBookmarks == true) { Globals.Ribbons.Ribbon.checkBox_mostra_indicadores.Checked = true; }
                if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowBookmarks == false) { Globals.Ribbons.Ribbon.checkBox_mostra_indicadores.Checked = false; }

                //Revisa a habilitação do CheckBox "Ver código" do Ribbon
                if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes == true) { Globals.Ribbons.Ribbon.checkBox_vercodigo_campos.Checked = true; }
                if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes == false) { Globals.Ribbons.Ribbon.checkBox_vercodigo_campos.Checked = false; }

                //Revisa a habilitação do CheckBox "Atualizar antes de imprimir" do Ribbon
                if (Globals.ThisAddIn.Application.Options.UpdateFieldsAtPrint == true) { Globals.Ribbons.Ribbon.checkBox_atualizar_antes_de_imprimir_campos.Checked = true; }
                if (Globals.ThisAddIn.Application.Options.UpdateFieldsAtPrint == false) { Globals.Ribbons.Ribbon.checkBox_atualizar_antes_de_imprimir_campos.Checked = false; }
            }
            catch (System.Runtime.InteropServices.COMException) { }


            //Revisa o destaque dos botoes do TaskPane
            if (Globals.Ribbons.Ribbon.toggleButton_painel_de_estilos.Checked)
            {
                Globals.ThisAddIn.iMyUserControl.Remove_Destaque_Botoes(UserControl_ActiveDocument);

                await Tarefa.Run(() =>
                {
                    if (Globals.ThisAddIn.Application.Selection.Tables.Count == 0) // Inseri pq selectionar paragrafos com tabela causa problemas de seleção.
                    {
                        List<Paragraph> paragrafosSelecionados = Globals.ThisAddIn.Application.Selection.Paragraphs.Cast<Paragraph>().ToList();

                        foreach (Paragraph p in paragrafosSelecionados)
                        {
                            if (token.IsCancellationRequested)
                                break;

                            Style estilo = null;
                            if (p.Range.StoryType == WdStoryType.wdMainTextStory)
                            {
                                try { estilo = p.Range.get_Style(); } catch (System.Runtime.InteropServices.COMException) { }

                                if (estilo != null && UserControl_ActiveDocument.dict_estilo_e_botao.ContainsKey(estilo.NameLocal))
                                {
                                    System.Windows.Forms.Button botao = UserControl_ActiveDocument.dict_estilo_e_botao[estilo.NameLocal];
                                    UserControl_ActiveDocument.Habilita_Destaca(botao, true, true);
                                }
                            }
                            if (p.Range.StoryType == WdStoryType.wdFootnotesStory)
                            {
                                Range Selecao_inicial = Globals.ThisAddIn.Application.Selection.Range; //Salva a seleção inicial (Inseri pq estilo = p.Range.ParagraphFormat.get_Style(); estava modificando implicitamente a selação)
                                try { estilo = p.Range.ParagraphFormat.get_Style(); } catch (System.Runtime.InteropServices.COMException) { }
                                Selecao_inicial.Select(); // Restaura a seleção inicial
                                if (estilo != null && UserControl_ActiveDocument.dict_estilo_e_botao.ContainsKey(estilo.NameLocal))
                                {
                                    System.Windows.Forms.Button botao = UserControl_ActiveDocument.dict_estilo_e_botao[estilo.NameLocal];
                                    UserControl_ActiveDocument.Habilita_Destaca(botao, true, true);
                                }
                            }
                        }
                    }
                }, token);
            }
        }
    }
}
