using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualBasic;
using System.Diagnostics;
using System.Windows.Controls;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using Tarefa = System.Threading.Tasks.Task;
using System.Linq;
using System.Text.RegularExpressions;


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

        private void button_formata_pagina_Click(object sender, RibbonControlEventArgs e)
        {
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
            Globals.ThisAddIn.Application.ActiveDocument.PageSetup.MirrorMargins = 0;
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
            Globals.ThisAddIn.Dicionario_Doc_e_UserControl[Globals.ThisAddIn.Application.ActiveDocument].Importa_todos_estilos();

            // CABEÇALHO DA PRIMEIRA PÁGINA
            // Apaga texto do cabeçalho da primeira pagina, inclusive os bookmarks e content controls
            Range cabecalho_1a_pagina = Globals.ThisAddIn.Application.ActiveDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
            Exclui_Bookmarks(cabecalho_1a_pagina);
            Exclui_ContentControls(cabecalho_1a_pagina);
            cabecalho_1a_pagina.Text = "";
            // Insere cabeçalho da primeira pagina
            Globals.Ribbons.Ribbon.inserir_autotexto(cabecalho_1a_pagina, "cabecalho1");
            // Deleta o último parágrafo do cabeçalho da primeira página
            cabecalho_1a_pagina.Paragraphs[cabecalho_1a_pagina.Paragraphs.Count].Range.Delete();

            // INICIO DO LAUDO
            // Apaga texto do início do laudo, inclusive os bookmarks e content controls
            Range inicio_do_laudo = encontrarRangedoIniciodoLaudo(Globals.ThisAddIn.Application.ActiveDocument);
            if (inicio_do_laudo == null)
            {
                Globals.ThisAddIn.Application.ActiveDocument.Range(0).InsertParagraphBefore();
                Globals.Ribbons.Ribbon.inserir_autotexto(Globals.ThisAddIn.Application.ActiveDocument.Range(0).Paragraphs[1].Range, "inicio_do_laudo");
            }
            else
            {
                string unidade = Tenta_deduzir_o_valor_de_unidade(inicio_do_laudo);
                string subtitulo = Tenta_deduzir_o_valor_de_subtitulo(inicio_do_laudo);
                Exclui_Bookmarks(inicio_do_laudo);
                Exclui_ContentControls(inicio_do_laudo);
                inicio_do_laudo.Text = "";
                // Insere início do laudo
                Globals.Ribbons.Ribbon.inserir_autotexto(inicio_do_laudo, "inicio_do_laudo");
                // Ajusta os ContentControl DropdownList Unidade e Subtítulo
                if (unidade != null) { }
                if (subtitulo != null) { }
            }

            // CABEÇALHO DAS OUTRAS PÁGINAS
            // Apaga texto do cabeçalho das outras páginas, inclusive os bookmarks e content controls
            Range cabecalho_outras_paginas = Globals.ThisAddIn.Application.ActiveDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
            Exclui_Bookmarks(cabecalho_outras_paginas);
            Exclui_ContentControls(cabecalho_outras_paginas);
            cabecalho_outras_paginas.Text = "";
            // Insere cabeçalho das outras páginas
            Globals.Ribbons.Ribbon.inserir_autotexto(cabecalho_outras_paginas, "cabecalho2");
            // Deleta o último parágrafo do cabeçalho das outras páginas
            cabecalho_outras_paginas.Paragraphs[cabecalho_outras_paginas.Paragraphs.Count].Range.Delete();

            // SEÇÃO DE CONCLUSÃO
            // Insere secao_de_conclusao
            Range UltimoParagrafo = EncontrarUltimoParagrafo("resposta aos quesitos");
            if (UltimoParagrafo == null) UltimoParagrafo = EncontrarUltimoParagrafo("conclusão");
            if (UltimoParagrafo != null)
            {
                Exclui_Bookmarks(UltimoParagrafo);
                Exclui_ContentControls(UltimoParagrafo);
                UltimoParagrafo.Text = "";
                Globals.Ribbons.Ribbon.inserir_autotexto(UltimoParagrafo, "secao_de_conclusao");
            }

        }

        private Range encontrarRangedoIniciodoLaudo(Document doc)
        {
            Range range = doc.Content;

            // Usa o Find para procurar pelo padrão "LAUDO Nº" (ou "LAUDO N°")
            // A busca deve cobrir o texto inicial do laudo, ajustando o padrão conforme necessário
            Find find = range.Find;
            find.ClearFormatting();
            find.Text = @"[lL][aA][uU][dD][oO]([ ]*)N* abaixo transcrito[os]";
            find.MatchCase = false;
            find.IgnorePunct = true;
            find.IgnoreSpace = true;
            find.MatchWildcards = true;  // Permite usar expressões regulares no Find

            // Executa a busca até encontrar o título do laudo
            bool encontrado = find.Execute();

            if (encontrado)
            {
                return range;
            }
            return null;
        }

        private string Tenta_deduzir_o_valor_de_unidade(Range range) { return null; }
        private string Tenta_deduzir_o_valor_de_subtitulo(Range range) { return null; }

        // Função para procurar o último parágrafo com o critério especificado
        static Range EncontrarUltimoParagrafo(string textoBusca)
        {
            // Obter o intervalo do conteúdo do documento
            Range range = Globals.ThisAddIn.Application.ActiveDocument.Content;
            Find find = range.Find;

            // Configurar o critério de busca
            find.Text = textoBusca;
            find.MatchCase = false; // Ignorar maiúsculas/minúsculas
            find.MatchWholeWord = true; // Procurar palavra completa
            find.Forward = false; // Buscar na direção do começo do documento (trás para frente)
            find.Wrap = WdFindWrap.wdFindStop; // Não reiniciar ao encontrar a primeira ocorrência

            // Realizar a busca
            Range ultimoRangeEncontrado = null;

            // Realizar a busca de trás para frente
            while (find.Execute())
            {
                // Armazenar o último range encontrado
                ultimoRangeEncontrado = range.Duplicate;
                range.SetRange(range.Start - 1, Globals.ThisAddIn.Application.ActiveDocument.Content.Start); // Avançar a busca para o início
            }

            return ultimoRangeEncontrado;
        }

        private void Exclui_Bookmarks(Range range) 
        {
            if (range == null) return;
            foreach (Microsoft.Office.Interop.Word.Bookmark bookmark in range.Bookmarks)
            {
                bookmark.Delete();
            }
        }
        private void Exclui_ContentControls(Range range)
        {
            if (range == null) return;
            while (range.ContentControls.Count > 0)
            {
                foreach (Microsoft.Office.Interop.Word.ContentControl cc in range.ContentControls)
                {
                    cc.LockContentControl = false;
                    cc.Delete();
                }
            }
        }

        private void button_habilita_edicao_Click(object sender, RibbonControlEventArgs e)
        {
            //Range range = Globals.ThisAddIn.Application.Selection.Range;
            //while (range.ContentControls.Count > 0)
            //{
            //    foreach (Microsoft.Office.Interop.Word.ContentControl cc in range.ContentControls)
            //    {
            //        cc.LockContentControl = false;
            //        cc.Delete();
            //    }
            //}
            Exclui_ContentControls(Globals.ThisAddIn.Application.Selection.Range);
        }
    }
}