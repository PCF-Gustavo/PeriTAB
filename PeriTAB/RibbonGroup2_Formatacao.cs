﻿using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Tarefa = System.Threading.Tasks.Task;


namespace PeriTAB
{
    public partial class Ribbon
    {
        // Cria instância das classes
        Class_ContentControlOnExit_Event iClass_ContentControlOnExit_Event = new Class_ContentControlOnExit_Event();
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

            DeleteEmptyParagraphsAtStart(Globals.ThisAddIn.Application.ActiveDocument.Content);
        }
        public void DeleteEmptyParagraphsAtStart(Range range)
        {
            // Enquanto o primeiro parágrafo for vazio (só com quebras de linha ou espaços)
            while (range.Paragraphs.Count > 0)
            {
                // Verifica se o primeiro parágrafo é vazio (somente espaços ou quebras de linha)
                string text = range.Paragraphs[1].Range.Text.Trim(); // Obtém o texto do parágrafo e remove espaços extras
                if (string.IsNullOrEmpty(text))  // Se o parágrafo for vazio
                {
                    range.Paragraphs[1].Range.Delete();  // Deleta o parágrafo
                }
                else
                {
                    break;  // Sai do loop se o parágrafo não for vazio
                }
            }
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
            string unidade = null;
            string fim_do_preambulo = null;

            // INICIO DO LAUDO
            // Apaga texto do início do laudo, inclusive os bookmarks e content controls
            Range inicio_do_laudo = encontrarRangedoIniciodoLaudo(Globals.ThisAddIn.Application.ActiveDocument);
            //MessageBox.Show(inicio_do_laudo.Text);
            if (inicio_do_laudo == null)
            {
                Globals.ThisAddIn.Application.ActiveDocument.Range(0).InsertParagraphBefore();
                Globals.Ribbons.Ribbon.inserir_autotexto(Globals.ThisAddIn.Application.ActiveDocument.Range(0).Paragraphs[1].Range, "inicio_do_laudo");
            }
            else
            {
                unidade = SearchTextWithRegex(inicio_do_laudo, @"\b([A-Z]{2,})(\s*/\s*([A-Z]{2,}))+(\s*/\s*([A-Z]{2,}))*\b");
                string subtitulo = SearchTextWithRegex(inicio_do_laudo, @"\((.*?)\)");
                fim_do_preambulo = string.Join(" ", inicio_do_laudo.Text.Split(' ').Where(word => !string.IsNullOrEmpty(word)).Reverse().Take(7).Reverse());
                Exclui_Bookmarks(inicio_do_laudo);
                Exclui_ContentControls(inicio_do_laudo);
                List<Paragraph> lista_de_paragrafos_de_inicio_do_laudo = inicio_do_laudo.Paragraphs.Cast<Paragraph>().ToList();
                foreach (Paragraph p in lista_de_paragrafos_de_inicio_do_laudo)
                {
                    p.Range.Delete();
                }
                // Insere início do laudo
                Globals.Ribbons.Ribbon.inserir_autotexto(inicio_do_laudo, "inicio_do_laudo");
                // Ajusta os ContentControl DropdownList Unidade e Subtítulo
                if (unidade != null)
                {
                    string maisProximo = EncontrarMaisProximo(unidade, Class_ContentControlOnExit_Event.Lista_Unidade);
                    iClass_ContentControlOnExit_Event.ChangeEntry(iClass_ContentControlOnExit_Event.GetContentControl("Unidade"), maisProximo);
                }
                if (subtitulo != null)
                {
                    string maisProximo = EncontrarMaisProximo(subtitulo, Class_ContentControlOnExit_Event.Lista_Subtitulos);
                    iClass_ContentControlOnExit_Event.ChangeEntry(iClass_ContentControlOnExit_Event.GetContentControl("Subtítulo"), maisProximo);
                }
                if (fim_do_preambulo != null)
                {
                    string maisProximo = EncontrarMaisProximo(fim_do_preambulo, Class_ContentControlOnExit_Event.Lista_fim_do_preambulo);
                    iClass_ContentControlOnExit_Event.ChangeEntry(iClass_ContentControlOnExit_Event.GetContentControl("Fim do preâmbulo"), maisProximo);
                }

            }

            // CABEÇALHO DA PRIMEIRA PÁGINA
            // Apaga texto do cabeçalho da primeira pagina, inclusive os bookmarks e content controls
            Range cabecalho_1a_pagina = Globals.ThisAddIn.Application.ActiveDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
            Exclui_Bookmarks(cabecalho_1a_pagina);
            Exclui_ContentControls(cabecalho_1a_pagina);
            cabecalho_1a_pagina.Text = "";
            // Insere cabeçalho da primeira pagina
            Globals.Ribbons.Ribbon.inserir_autotexto(cabecalho_1a_pagina, "cabecalho1");
            if (unidade != null)
            {
                string maisProximo = EncontrarMaisProximo(unidade, Class_ContentControlOnExit_Event.Lista_Unidade);
                iClass_ContentControlOnExit_Event.ChangeEntry(iClass_ContentControlOnExit_Event.GetContentControl("Unidade da PF"), Class_ContentControlOnExit_Event.dict_Unidade_e_Unidade_da_PF[maisProximo]);
                iClass_ContentControlOnExit_Event.Add_or_remove_ultima_linha_cabecalho1();
                iClass_ContentControlOnExit_Event.Muda_Tipo_de_unidade_de_criminalistica();
            }
            // Deleta o último parágrafo do cabeçalho da primeira página
            //cabecalho_1a_pagina.Paragraphs[cabecalho_1a_pagina.Paragraphs.Count].Range.Delete();

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
                string dasd = UltimoParagrafo.Text;
                UltimoParagrafo.Text = "";
                Globals.Ribbons.Ribbon.inserir_autotexto(UltimoParagrafo, "secao_de_conclusao");
                if (fim_do_preambulo != null)
                {
                    string maisProximo = EncontrarMaisProximo(fim_do_preambulo, Class_ContentControlOnExit_Event.Lista_fim_do_preambulo);
                    iClass_ContentControlOnExit_Event.ChangeEntry(iClass_ContentControlOnExit_Event.GetContentControl("Seção de conclusão"), Class_ContentControlOnExit_Event.dict_Fim_do_preambulo_e_Secao_de_conclusao[maisProximo]);
                }
            }

        }

        private Range encontrarRangedoIniciodoLaudo(Document doc)
        {
            Range range = doc.Content;

            // Usa o Find para procurar pelo padrão "LAUDO Nº" (ou "LAUDO N°")
            // A busca deve cobrir o texto inicial do laudo, ajustando o padrão conforme necessário
            Find find = range.Find;
            find.ClearFormatting();
            find.Text = @"[lL][aA][uU][dD][oO]([ ]*)N* abaixo transcrito";
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

        private string SearchTextWithRegex(Range range, string regex)
        {
            // Regex para encontrar a correspondência no texto
            Match match = Regex.Match(range.Text, regex, RegexOptions.IgnoreCase | RegexOptions.Multiline);

            if (match.Success)
            {
                // Retorna o grupo encontrado (SETEC/SR/PF/MA ou outras variações)
                return match.Value;
            }
            else
            {
                // Caso não encontre, retorna uma string vazia ou uma mensagem de erro
                return null;
            }
        }

        public static int CalcularDistanciaLevenshtein(string s1, string s2)
        {
            int n = s1.Length;
            int m = s2.Length;
            int[,] matriz = new int[n + 1, m + 1];

            // Preenche a primeira linha e a primeira coluna
            for (int i = 0; i <= n; i++) matriz[i, 0] = i;
            for (int j = 0; j <= m; j++) matriz[0, j] = j;

            // Preenche o resto da matriz
            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int custo = (s1[i - 1] == s2[j - 1]) ? 0 : 1;
                    matriz[i, j] = Math.Min(Math.Min(matriz[i - 1, j] + 1, matriz[i, j - 1] + 1), matriz[i - 1, j - 1] + custo);
                }
            }

            return matriz[n, m];
        }

        // Função para encontrar o valor mais próximo
        public static string EncontrarMaisProximo(string referencia, List<string> valores)
        {
            string maisProximo = null;
            int menorDistancia = int.MaxValue;

            foreach (string valor in valores)
            {
                int distancia = CalcularDistanciaLevenshtein(referencia, valor);
                if (distancia < menorDistancia)
                {
                    menorDistancia = distancia;
                    maisProximo = valor;
                }
            }

            return maisProximo;
        }

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
            foreach (Paragraph p in range.Paragraphs)
            {
                List<ContentControl> contentControls = new List<ContentControl>(p.Range.ContentControls.Cast<ContentControl>());
                foreach (ContentControl cc in contentControls)
                {
                    cc.LockContentControl = false;
                    cc.Delete();
                }
            }
        }
        

        private void button_habilita_edicao_Click(object sender, RibbonControlEventArgs e)
        {
            Exclui_ContentControls(Globals.ThisAddIn.Application.Selection.Range);
        }
    }
}