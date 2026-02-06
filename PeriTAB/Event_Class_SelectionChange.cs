using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace PeriTAB
{
    internal class Class_SelectionChange_Event
    {
        //private static CancellationTokenSource cancellationTokenSource = null;

        

        public void Evento_SelectionChange()
        {
            Globals.ThisAddIn.Application.WindowSelectionChange += new ApplicationEvents4_WindowSelectionChangeEventHandler(Metodo_SelectionChange);
        }

        private /*async*/ void Metodo_SelectionChange(Selection Sel)
        {
            // Se houver uma operação em andamento, cancelamos a execução anterior
            //cancellationTokenSource?.Cancel();

            //if (cancellationTokenSource != null)
            //cancellationTokenSource.Cancel(); // Cancela a execução anterior

            // Criamos uma nova fonte de cancelamento para a próxima execução
            //cancellationTokenSource = new CancellationTokenSource();
            //CancellationToken token = cancellationTokenSource.Token;

            //Declara instacias das classes
            //MyUserControl UserControl_ActiveDocument = Globals.ThisAddIn.Dicionario_Doc_e_UserControl[Globals.ThisAddIn.Application.ActiveDocument];
            if (!Globals.ThisAddIn.Dicionario_Window_e_UserControl.TryGetValue(Globals.ThisAddIn.Application.ActiveWindow, out MyUserControl UserControl_ActiveWindow)) return;

            if (Globals.Ribbons.Ribbon.ToggleButton_painel_de_estilos.Checked) UserControl_ActiveWindow.Atualiza_Destaque_Botoes();



            //try
            //{
            //    //Revisa a habilitação do CheckBox "Destacar campos" do Ribbon
            //    if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)1) { Globals.Ribbons.Ribbon.CheckBox_destaca_campos.Checked = true; }
            //    if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)0 | Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)2) { Globals.Ribbons.Ribbon.CheckBox_destaca_campos.Checked = false; }

            //    //Revisa a habilitação do CheckBox "Mostrar indicadores" do Ribbon
            //    if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowBookmarks == true) { Globals.Ribbons.Ribbon.CheckBox_mostra_indicadores.Checked = true; }
            //    if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowBookmarks == false) { Globals.Ribbons.Ribbon.CheckBox_mostra_indicadores.Checked = false; }
            //}
            //catch (System.Runtime.InteropServices.COMException) { }

            

            //Class_CustomTaskPanes.Atualiza_Destaques(UserControl_ActiveWindow);

            //Revisa o destaque dos botoes do TaskPane
            //if (Globals.Ribbons.Ribbon.toggleButton_painel_de_estilos.Checked)
            //{
            //    UserControl_ActiveWindow.Remove_Destaque_Botoes();

            //    await Task.Run(() =>
            //    {
            //        try
            //        {
            //            if (Globals.ThisAddIn.Application.Selection.Tables.Count == 0) // Inseri pq selecionar paragrafos com tabela causa problemas de seleção.
            //            {
            //                List<Paragraph> paragrafosSelecionados = Globals.ThisAddIn.Application.Selection.Paragraphs.Cast<Paragraph>().ToList();

            //                foreach (Paragraph p in paragrafosSelecionados)
            //                {
            //                    if (token.IsCancellationRequested)
            //                        break;

            //                    Style estilo = null;
            //                    if (p.Range.StoryType == WdStoryType.wdMainTextStory)
            //                    {
            //                        //try { estilo = p.Range.get_Style(); } catch (System.Runtime.InteropServices.COMException) { }
            //                        estilo = p.Range.get_Style();

            //                        if (estilo != null && UserControl_ActiveWindow.Dicionario_Estilo_e_Botao.ContainsKey(estilo.NameLocal))
            //                        {
            //                            System.Windows.Forms.Button botao = UserControl_ActiveWindow.Dicionario_Estilo_e_Botao[estilo.NameLocal];
            //                            UserControl_ActiveWindow.Habilita_Destaca(botao, true, true);
            //                        }
            //                    }
            //                    if (p.Range.StoryType == WdStoryType.wdFootnotesStory)
            //                    {
            //                        Range Selecao_inicial = Globals.ThisAddIn.Application.Selection.Range; //Salva a seleção inicial (Inseri pq estilo = p.Range.ParagraphFormat.get_Style(); estava modificando implicitamente a selação)
            //                        //try { estilo = p.Range.ParagraphFormat.get_Style(); } catch (System.Runtime.InteropServices.COMException) { }
            //                        estilo = p.Range.ParagraphFormat.get_Style();
            //                        Selecao_inicial.Select(); // Restaura a seleção inicial
            //                        if (estilo != null && UserControl_ActiveWindow.Dicionario_Estilo_e_Botao.ContainsKey(estilo.NameLocal))
            //                        {
            //                            System.Windows.Forms.Button botao = UserControl_ActiveWindow.Dicionario_Estilo_e_Botao[estilo.NameLocal];
            //                            UserControl_ActiveWindow.Habilita_Destaca(botao, true, true);
            //                        }
            //                    }
            //                }
            //            }
            //        }
            //        catch (System.Runtime.InteropServices.COMException) { }
            //    }, token);
            //}
        }
    }
}
