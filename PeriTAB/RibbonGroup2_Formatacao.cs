using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualBasic;
using System.Diagnostics;
using System.Windows.Controls;
using Tarefa = System.Threading.Tasks.Task;


namespace PeriTAB
{
    public partial class Ribbon
    {
        // Cria instância das classes
        public MyUserControl iMyUserControl;

        private async void button_alinha_legenda_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;

            // Executa as tarefas em segundo plano
            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.Run("alinha_legenda");
            });

            // Após a execução das tarefas, atualiza a UI na Thread principal
            RibbonButton.Image = Properties.Resources.lupa;
            RibbonButton.Enabled = true;
        }

        private void toggleButton_painel_de_estilos_Click(object sender, RibbonControlEventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (PeriTAB.Ribbon.Variables.debugging) { stopwatch.Start(); }
            string msg_StatusBar = "";
            var botao_toggle = (Microsoft.Office.Tools.Ribbon.RibbonToggleButton)sender;
            if (botao_toggle.Checked == true)
            {
                iClass_CustomTaskPanes.Visible(true);
                if (PeriTAB.Ribbon.Variables.debugging) msg_StatusBar = "Painel de Estilos: Aberto";
            }
            if (botao_toggle.Checked == false)
            {
                iClass_CustomTaskPanes.Visible(false);
                if (PeriTAB.Ribbon.Variables.debugging) msg_StatusBar = "Painel de Estilos: Fechado";
            }

            if (PeriTAB.Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
        }

        private void button_formata_cabecalhos_e_preambulo_Click(object sender, RibbonControlEventArgs e)
        {
            // Page Setup
            Globals.ThisAddIn.Application.ActiveDocument.PageSetup.Orientation = WdOrientation.wdOrientPortrait;
            Globals.ThisAddIn.Application.ActiveDocument.PageSetup.PageWidth = Globals.ThisAddIn.Application.CentimetersToPoints(21);
            Globals.ThisAddIn.Application.ActiveDocument.PageSetup.PageHeight = Globals.ThisAddIn.Application.CentimetersToPoints(29.7f);
            Globals.ThisAddIn.Application.ActiveDocument.PageSetup.TopMargin = Globals.ThisAddIn.Application.CentimetersToPoints(2);
            Globals.ThisAddIn.Application.ActiveDocument.PageSetup.TopMargin = Globals.ThisAddIn.Application.CentimetersToPoints(2);
            Globals.ThisAddIn.Application.ActiveDocument.PageSetup.BottomMargin = Globals.ThisAddIn.Application.CentimetersToPoints(2);
            Globals.ThisAddIn.Application.ActiveDocument.PageSetup.LeftMargin = Globals.ThisAddIn.Application.CentimetersToPoints(3);
            Globals.ThisAddIn.Application.ActiveDocument.PageSetup.RightMargin = Globals.ThisAddIn.Application.CentimetersToPoints(2);
            Globals.ThisAddIn.Application.ActiveDocument.PageSetup.HeaderDistance = Globals.ThisAddIn.Application.CentimetersToPoints(1);
            Globals.ThisAddIn.Application.ActiveDocument.PageSetup.FooterDistance = Globals.ThisAddIn.Application.CentimetersToPoints(.5f);
            Globals.ThisAddIn.Application.ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter = -1;
            Globals.ThisAddIn.Application.ActiveDocument.PageSetup.OddAndEvenPagesHeaderFooter = 0;

            Globals.ThisAddIn.Dicionario_Doc_e_UserControl[Globals.ThisAddIn.Application.ActiveDocument].Importa_todos_estilos();

            // Apaga cabeçalhos
            Globals.ThisAddIn.Application.ActiveDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "";
            Globals.ThisAddIn.Application.ActiveDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "";

            // Insere cabeçalho1
            Globals.Ribbons.Ribbon.inserir_autotexto(Globals.ThisAddIn.Application.ActiveDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range, "cabecalho1");
            Globals.ThisAddIn.Application.ActiveDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Paragraphs[Globals.ThisAddIn.Application.ActiveDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Paragraphs.Count].Range.Delete();

            // Insere preâmbulo
            Globals.ThisAddIn.Application.ActiveDocument.Range(0).InsertParagraphBefore();
            Globals.Ribbons.Ribbon.inserir_autotexto(Globals.ThisAddIn.Application.ActiveDocument.Range(0).Paragraphs[1].Range, "inicio_do_laudo");

            //// Insere cabeçalho2
            //Globals.Ribbons.Ribbon.inserir_autotexto(Globals.ThisAddIn.Application.ActiveDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range, "cabecalho2");
            //Globals.ThisAddIn.Application.ActiveDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[Globals.ThisAddIn.Application.ActiveDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs.Count].Range.Delete();
        }
    }
}