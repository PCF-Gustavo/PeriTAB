﻿using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Drawing;
using Tarefa = System.Threading.Tasks.Task;
using System.Windows.Threading;
using System.Threading;

namespace PeriTAB
{
    public partial class Ribbon
    {
        // Cria instância das classes
        Class_ContentControlOnExit_Event iClass_ContentControlOnExit_Event = new Class_ContentControlOnExit_Event();
        public MyUserControl iMyUserControl;

        //private void button_adiciona_indicador_Click(object sender, RibbonControlEventArgs e)
        //{
        //    // Inicializa o índice do bookmark
        //    int i = 1;
        //    string bookmarkName = $"indicador{i}_PeriTAB";

        //    // Verifica se o bookmark já existe e incrementa o número até encontrar um nome disponível
        //    while (BookmarkExists(Globals.ThisAddIn.Application.ActiveDocument, bookmarkName))
        //    {
        //        i++;
        //        bookmarkName = $"indicador{i}_PeriTAB";
        //    }

        //    // Adiciona o bookmark com o nome encontrado
        //    try { Globals.ThisAddIn.Application.Selection.Bookmarks.Add(bookmarkName); } catch { }
        //}

        // Função para verificar se o bookmark já existe
        //static bool BookmarkExists(Document doc, string bookmarkName)
        //{
        //    foreach (Bookmark bookmark in doc.Bookmarks)
        //    {
        //        if (bookmark.Name == bookmarkName)
        //        {
        //            return true;
        //        }
        //    }
        //    return false;
        //}

        private async void button_alinha_legenda_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;

            // Executa as tarefas em segundo plano
            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                Globals.ThisAddIn.Application.Run("alinha_legenda");
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            // Após a execução das tarefas, atualiza a UI na Thread principal
            RibbonButton.Image = Properties.Resources.seta3;
            RibbonButton.Enabled = true;
        }

        private async void button_pagina_em_paisagem_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;
            bool success = true;
            string msg_StatusBar = RibbonButton.Label + ": ";
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            // Executa as tarefas em segundo plano
            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                Range r1 = Globals.ThisAddIn.Application.Selection.Range.Duplicate;
                Range r2 = Globals.ThisAddIn.Application.Selection.Range.Duplicate;
                Range r3 = Globals.ThisAddIn.Application.Selection.Range.Duplicate;
                r1.Collapse(WdCollapseDirection.wdCollapseStart);
                int pagina_inicio_selecao = r1.Information[WdInformation.wdActiveEndPageNumber];
                r2.Collapse(WdCollapseDirection.wdCollapseEnd);
                int pagina_fim_selecao = r2.Information[WdInformation.wdActiveEndPageNumber];

                r1 = r1.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, pagina_inicio_selecao);
                r2 = r2.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, pagina_fim_selecao + 1);

                if (pagina_inicio_selecao == 1 & pagina_fim_selecao == Globals.ThisAddIn.Application.ActiveDocument.ComputeStatistics(WdStatistic.wdStatisticPages))
                {
                    Globals.ThisAddIn.Application.ActiveDocument.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
                }
                else if (pagina_inicio_selecao == 1)
                {
                    r2.InsertBreak(WdBreakType.wdSectionBreakNextPage);
                    r3.Sections[1].PageSetup.Orientation = WdOrientation.wdOrientLandscape;
                    r3.Sections[1].PageSetup.DifferentFirstPageHeaderFooter = -1;
                    r3.Sections[1].PageSetup.OddAndEvenPagesHeaderFooter = 0;
                    Section next_section = GetNextSection(r3.Sections[1]);
                    next_section.PageSetup.DifferentFirstPageHeaderFooter = 0;
                    next_section.PageSetup.OddAndEvenPagesHeaderFooter = 0;
                }
                else if (pagina_fim_selecao == Globals.ThisAddIn.Application.ActiveDocument.ComputeStatistics(WdStatistic.wdStatisticPages))
                {
                    r1.InsertBreak(WdBreakType.wdSectionBreakNextPage);
                    r3.Sections[1].PageSetup.Orientation = WdOrientation.wdOrientLandscape;
                    r3.Sections[1].PageSetup.DifferentFirstPageHeaderFooter = 0;
                    r3.Sections[1].PageSetup.OddAndEvenPagesHeaderFooter = 0;
                }
                else
                {
                    r1.InsertBreak(WdBreakType.wdSectionBreakNextPage);
                    r2.InsertBreak(WdBreakType.wdSectionBreakNextPage);
                    r3.Sections[1].PageSetup.Orientation = WdOrientation.wdOrientLandscape;
                    r3.Sections[1].PageSetup.DifferentFirstPageHeaderFooter = 0;
                    r3.Sections[1].PageSetup.OddAndEvenPagesHeaderFooter = 0;
                    Section next_section = GetNextSection(r3.Sections[1]);
                    next_section.PageSetup.DifferentFirstPageHeaderFooter = 0;
                    next_section.PageSetup.OddAndEvenPagesHeaderFooter = 0;
                }
                foreach (Section section in Globals.ThisAddIn.Application.ActiveDocument.Sections)
                {
                    foreach (HeaderFooter footer in section.Footers)
                    {
                        try
                        {
                            footer.LinkToPrevious = false;
                        }
                        catch { }
                    }
                }
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;

            // Após a execução das tarefas, atualiza a UI na Thread principal
            RibbonButton.Image = null;
            RibbonButton.Enabled = true;
        }


        public void DeleteEmptyParagraphsAtStart(Range range)
        {
            // Enquanto o primeiro parágrafo for vazio (só com quebras de linha ou espaços)
            while (range.Paragraphs.Count > 0 && Globals.ThisAddIn.Application.ActiveDocument.Paragraphs.Count > 1)
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

        static Section GetNextSection(Section section)
        {
            // Percorre as seções do documento
            for (int i = 1; i < Globals.ThisAddIn.Application.ActiveDocument.Sections.Count; i++)
            {
                // Se encontramos a seção corrente, verificamos se existe uma próxima seção
                if (Globals.ThisAddIn.Application.ActiveDocument.Sections[i].Range.Start == section.Range.Start)
                {
                    // Verifica se não é a última seção
                    if (i + 1 <= Globals.ThisAddIn.Application.ActiveDocument.Sections.Count)
                    {
                        return Globals.ThisAddIn.Application.ActiveDocument.Sections[i + 1]; // Retorna a próxima seção
                    }
                    break; // Se for a última seção, sai do loop
                }
            }

            // Se não encontrar uma próxima seção, retorna null
            return null;
        }

        private void toggleButton_painel_de_estilos_Click(object sender, RibbonControlEventArgs e)
        {
            bool success = true;
            string msg_StatusBar = ((RibbonToggleButton)sender).Label + ": ";

            var botao_toggle = (Microsoft.Office.Tools.Ribbon.RibbonToggleButton)sender;
            if (botao_toggle.Checked == true) iClass_CustomTaskPanes.Visible(true);
            if (botao_toggle.Checked == false) iClass_CustomTaskPanes.Visible(false);

            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
        }

        private string barra_de_progresso(int progress) 
        {
            char filledSquare = (char)0x2588;  // Caractere '█' (quadrado preenchido).
            char emptySquare = (char)0x2591;   // Caractere '░' (quadrado não preenchido).
            string progressBar = new string(filledSquare, progress) + new string(emptySquare, 10 - progress); // Cria a "barra de progresso".
            //Globals.ThisAddIn.Application.StatusBar = $"[{progressBar}]";
            return progressBar;
        }


        private /*async*/ void button_autoformata_laudo_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonButton RibbonButton = (RibbonButton)sender;
            
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;

            bool success = true;
            string msg_StatusBar = RibbonButton.Label + ": ";
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar + " " + barra_de_progresso(0);
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            foreach (Range storyRange in Globals.ThisAddIn.Application.ActiveDocument.StoryRanges) storyRange.Fields.Update();
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar + " " + barra_de_progresso(1);
            Globals.ThisAddIn.iMyUserControl.Importa_todos_estilos();
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar + " " + barra_de_progresso(2);

            //Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
            //await Tarefa.Run(() =>
            //{
            //Thread.CurrentThread.Priority = ThreadPriority.Highest;
            foreach (Section section in Globals.ThisAddIn.Application.ActiveDocument.Sections)
            {
                section.PageSetup.PaperSize = Microsoft.Office.Interop.Word.WdPaperSize.wdPaperA4;
                // Verificar a orientação e aplicar configurações específicas
                if (section.PageSetup.Orientation == WdOrientation.wdOrientPortrait)
                {
                    // Configurações para retrato (Portrait)
                    section.PageSetup.PageWidth = Globals.ThisAddIn.Application.CentimetersToPoints(21);
                    section.PageSetup.PageHeight = Globals.ThisAddIn.Application.CentimetersToPoints(29.7f);
                }
                else if (section.PageSetup.Orientation == WdOrientation.wdOrientLandscape)
                {
                    // Configurações para paisagem (Landscape)
                    section.PageSetup.PageWidth = Globals.ThisAddIn.Application.CentimetersToPoints(29.7f);
                    section.PageSetup.PageHeight = Globals.ThisAddIn.Application.CentimetersToPoints(21);
                }
                section.PageSetup.TopMargin = Globals.ThisAddIn.Application.CentimetersToPoints(2);
                section.PageSetup.BottomMargin = Globals.ThisAddIn.Application.CentimetersToPoints(2);
                section.PageSetup.LeftMargin = Globals.ThisAddIn.Application.CentimetersToPoints(3);
                section.PageSetup.RightMargin = Globals.ThisAddIn.Application.CentimetersToPoints(2);
                section.PageSetup.HeaderDistance = Globals.ThisAddIn.Application.CentimetersToPoints(1);
                section.PageSetup.FooterDistance = Globals.ThisAddIn.Application.CentimetersToPoints(.5f);
                //section.PageSetup.DifferentFirstPageHeaderFooter = -1;
                //section.PageSetup.OddAndEvenPagesHeaderFooter = 0;
                section.PageSetup.MirrorMargins = 0;
            }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar + " " + barra_de_progresso(3);
            DeleteEmptyParagraphsAtStart(Globals.ThisAddIn.Application.ActiveDocument.Content);

            string unidade = null;
            string fim_do_preambulo = null;

            // INICIO DO LAUDO
            // Apaga texto do início do laudo, inclusive os bookmarks e content controls
            Range inicio_do_laudo = encontrarRangedoIniciodoLaudo(Globals.ThisAddIn.Application.ActiveDocument);
            //MessageBox.Show(inicio_do_laudo.Text);
            if (inicio_do_laudo == null)
            {
                Globals.ThisAddIn.Application.ActiveDocument.Range(0).InsertParagraphBefore();
                Globals.Ribbons.Ribbon.inserir_autotexto(Globals.ThisAddIn.Application.ActiveDocument.Range(0).Paragraphs[1].Range, "inicio_do_laudo_PeriTAB");
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
                Globals.Ribbons.Ribbon.inserir_autotexto(inicio_do_laudo, "inicio_do_laudo_PeriTAB");
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
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar + " " + barra_de_progresso(4);
            bool isFirstSection = true;
            foreach (Section section in Globals.ThisAddIn.Application.ActiveDocument.Sections)
            {
                foreach (HeaderFooter header in section.Headers)
                {
                    if (isFirstSection)
                    {
                        section.PageSetup.DifferentFirstPageHeaderFooter = -1;
                        section.PageSetup.OddAndEvenPagesHeaderFooter = 0;
                    }
                    else
                    {
                        section.PageSetup.DifferentFirstPageHeaderFooter = 0;
                        section.PageSetup.OddAndEvenPagesHeaderFooter = 0;
                    }
                    if (isFirstSection && header.Index == WdHeaderFooterIndex.wdHeaderFooterPrimary)
                    {
                        // CABEÇALHO DA PRIMEIRA PÁGINA
                        // Apaga texto do cabeçalho da primeira pagina, inclusive os bookmarks e content controls
                        Range cabecalho_1a_pagina = section.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                        Exclui_Bookmarks(cabecalho_1a_pagina);
                        Exclui_ContentControls(cabecalho_1a_pagina);
                        cabecalho_1a_pagina.Text = "";
                        // Insere cabeçalho da primeira pagina
                        Globals.Ribbons.Ribbon.inserir_autotexto(cabecalho_1a_pagina, "cabecalho_1a_pagina_PeriTAB");
                        if (unidade != null)
                        {
                            string maisProximo = EncontrarMaisProximo(unidade, Class_ContentControlOnExit_Event.Lista_Unidade);
                            iClass_ContentControlOnExit_Event.ChangeEntry(iClass_ContentControlOnExit_Event.GetContentControl("Unidade da PF"), Class_ContentControlOnExit_Event.dict_Unidade_e_Unidade_da_PF[maisProximo]);
                            iClass_ContentControlOnExit_Event.Add_or_remove_ultima_linha_cabecalho1();
                            iClass_ContentControlOnExit_Event.Muda_Tipo_de_unidade_de_criminalistica();
                        }
                        // Deleta o último parágrafo do cabeçalho da primeira página
                        try { cabecalho_1a_pagina.Paragraphs[cabecalho_1a_pagina.Paragraphs.Count].Range.Delete(); }
                        catch (System.Runtime.InteropServices.COMException) { }
                    }
                    else
                    {
                        // CABEÇALHO DAS OUTRAS PÁGINAS
                        // Apaga texto do cabeçalho das outras páginas, inclusive os bookmarks e content controls
                        Range cabecalho_outras_paginas = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        Exclui_Bookmarks(cabecalho_outras_paginas);
                        Exclui_ContentControls(cabecalho_outras_paginas);
                        cabecalho_outras_paginas.Text = "";
                        // Insere cabeçalho das outras páginas
                        Globals.Ribbons.Ribbon.inserir_autotexto(cabecalho_outras_paginas, "cabecalho_exceto_1a_pag_PeriTAB");
                        // Deleta o último parágrafo do cabeçalho das outras páginas
                        cabecalho_outras_paginas.Paragraphs[cabecalho_outras_paginas.Paragraphs.Count].Range.Delete();
                    }
                }
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar + " " + barra_de_progresso(5);
                foreach (HeaderFooter footer in section.Footers)
                {
                    if (isFirstSection && footer.Index == WdHeaderFooterIndex.wdHeaderFooterPrimary)
                    {
                        // RODAPÉ DA PRIMEIRA PÁGINA
                        // Apaga texto do rodapé da primeira pagina, inclusive os bookmarks e content controls
                        Range rodape_1a_pagina = section.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                        Exclui_Bookmarks(rodape_1a_pagina);
                        Exclui_ContentControls(rodape_1a_pagina);
                        rodape_1a_pagina.Text = "";
                        // Insere cabeçalho da primeira pagina
                        Globals.Ribbons.Ribbon.inserir_autotexto(rodape_1a_pagina, "rodape_1a_pagina_PeriTAB");
                    }
                    else
                    {
                        // RODAPE DAS OUTRAS PÁGINAS
                        // Apaga texto do rodape das outras páginas, inclusive os bookmarks e content controls
                        Range rodape_outras_paginas = section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        Exclui_Bookmarks(rodape_outras_paginas);
                        Exclui_ContentControls(rodape_outras_paginas);
                        rodape_outras_paginas.Text = "";
                        // Insere cabeçalho da primeira pagina
                        Globals.Ribbons.Ribbon.inserir_autotexto(rodape_outras_paginas, "rodape_exceto_1a_pag_PeriTAB");
                    }
                }
                isFirstSection = false;
            }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar + " " + barra_de_progresso(6);
            // SEÇÃO DE CONCLUSÃO
            // Insere Seção de conclusão
            Range secao_de_conclusao_range = EncontrarUltimoParagrafo("resposta aos quesitos");
            if (secao_de_conclusao_range == null) secao_de_conclusao_range = EncontrarUltimoParagrafo_wildcard("[rR]ESPOSTA*[qQ]UESITO?");
            if (secao_de_conclusao_range == null) secao_de_conclusao_range = EncontrarUltimoParagrafo_wildcard("[rR]esposta*[qQ]uesito?");
            if (secao_de_conclusao_range == null) secao_de_conclusao_range = EncontrarUltimoParagrafo("conclusão");
            if (secao_de_conclusao_range == null) secao_de_conclusao_range = EncontrarUltimoParagrafo("conclusao");
            if (secao_de_conclusao_range != null)
            {
                Exclui_Bookmarks(secao_de_conclusao_range);
                Exclui_ContentControls(secao_de_conclusao_range);
                secao_de_conclusao_range.Text = "";
                Globals.Ribbons.Ribbon.inserir_autotexto(secao_de_conclusao_range, "secao_de_conclusao_PeriTAB");
                if (fim_do_preambulo != null)
                {
                    string maisProximo = EncontrarMaisProximo(fim_do_preambulo, Class_ContentControlOnExit_Event.Lista_fim_do_preambulo);
                    iClass_ContentControlOnExit_Event.ChangeEntry(iClass_ContentControlOnExit_Event.GetContentControl("Seção de conclusão"), Class_ContentControlOnExit_Event.dict_Fim_do_preambulo_e_Secao_de_conclusao[maisProximo]);
                }
            }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar + " " + barra_de_progresso(7);
            // FECHO
            Range fecho_range = EncontrarUltimoParagrafo_wildcard("[nN]ada([ ]*)mais([ ]*)havendo*CRIMINAL([ ]*)FEDERAL"); // Procura texto de fecho e assinatura
            if (fecho_range != null)
            {
                Range assinado_digitalmente_range = EncontrarUltimoParagrafo_wildcard("[aA]ssinado [dD]igitalmente");
                if (assinado_digitalmente_range == null) assinado_digitalmente_range = EncontrarUltimoParagrafo_wildcard("[dD]igitalmente [aA]ssinado");
                if (assinado_digitalmente_range != null)
                {
                    // Procura complemento de apêndices e anexos
                    int startIndex = fecho_range.Text.IndexOf("páginas");
                    int endIndex = fecho_range.Text.IndexOf(", digitalmente");
                    string complemento_texto = null;
                    if (startIndex != -1 && endIndex != -1)
                    {
                        complemento_texto = fecho_range.Text.Substring(startIndex + 7, endIndex - startIndex - 7);
                    }

                    string nome_do_perito = assinado_digitalmente_range.Paragraphs[1].Next().Range.Text; //Guarda nome do perito
                    Globals.Ribbons.Ribbon.inserir_autotexto(fecho_range, "fecho_1_PeriTAB"); // Insere autotexto do fecho e assinatura
                    Range assinado_digitalmente_range2 = EncontrarUltimoParagrafo("assinado digitalmente");
                    assinado_digitalmente_range2.Paragraphs[1].Next().Range.HighlightColorIndex = WdColorIndex.wdAuto;
                    assinado_digitalmente_range2.Paragraphs[1].Next().Range.Text = nome_do_perito; // Insere nome do perito

                    Range fecho_range2 = assinado_digitalmente_range2.Paragraphs[1].Previous().Range;
                    Range fecho_range3 = fecho_range2.Duplicate;
                    Range fecho_range4 = fecho_range2.Duplicate;

                    bool encontrado = fecho_range2.Find.Execute(", digitalmente");
                    if (encontrado)
                    { // Insere texto do antigo destacado em amarelo
                        fecho_range2.SetRange(fecho_range2.Start, fecho_range2.Start);
                        fecho_range2.InsertAfter(complemento_texto);
                        fecho_range3.SetRange(fecho_range2.Start, fecho_range2.Start + complemento_texto.Length);

                        // Alterar o fundo para amarelo (destaque)
                        fecho_range3.HighlightColorIndex = WdColorIndex.wdYellow;

                        // Atualiza o numero de páginas
                        fecho_range4.Fields.Update();
                    }
                }
            }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar + " " + barra_de_progresso(8);
            // DESTACA EM AMARELO TEXTO PRETOS (Red = 1, Green = 1, Blue = 0)
            Color quase_preto = Color.FromArgb(1, 1, 0);
            WdColor WdColor_quase_preto = (WdColor)(quase_preto.R + 0x100 * quase_preto.G + 0x10000 * quase_preto.B);
            foreach (Range StoryRanges in Globals.ThisAddIn.Application.ActiveDocument.StoryRanges)
            {
                foreach (Range word in StoryRanges.Words)
                {

                    if (word.Font.Color == WdColor_quase_preto)
                    {
                        // Destaca o texto em amarelo
                        word.HighlightColorIndex = WdColorIndex.wdYellow;
                    }
                }
            }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar + " " + barra_de_progresso(9);
            // DESTACA TODAS AS IMAGENS COM BORDAS AMARELAS
            foreach (InlineShape ishape in Globals.ThisAddIn.Application.ActiveDocument.StoryRanges[WdStoryType.wdMainTextStory].InlineShapes)
            {
                if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                {
                    ishape.Line.Visible = MsoTriState.msoTrue;
                    ishape.Line.Weight = 3;
                    ishape.Line.ForeColor.RGB = Color.FromArgb(0, 255, 255).ToArgb();
                }
            }
            //});
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar + " " + barra_de_progresso(10);
            //Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();

            Globals.ThisAddIn.iMyUserControl.Importa_todos_estilos(); // não sei pq precisa repetir essa importacao, mas tem laudo que perde a formatacao se nao faço isso.
            Globals.ThisAddIn.Application.ScreenUpdating = true;

            if (success) { msg_StatusBar += barra_de_progresso(10) + " Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;

            // Após a execução das tarefas, atualiza a UI na Thread principal
            RibbonButton.Image = Properties.Resources.checklist2;
            RibbonButton.Enabled = true;
        }

        private Range encontrarRangedoIniciodoLaudo(Document doc)
        {
            Range range = doc.Content;

            // Usa o Find para procurar pelo padrão "LAUDO Nº" (ou "LAUDO N°")
            Find find = range.Find;
            find.ClearFormatting();
            find.Text = @"[lL][aA][uU][dD][oO]([ ]*)[nN]* abaix*transcrit*";
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
                return match.Value;
            }
            else
            {
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

                if (range.Start <= Globals.ThisAddIn.Application.ActiveDocument.Content.Start)
                {
                    break; // Finalizar a busca quando atingir o início do documento
                }
            }
            return ultimoRangeEncontrado;
        }

        static Range EncontrarUltimoParagrafo_wildcard(string textoBusca)
        {
            // Obter o intervalo do conteúdo do documento
            Range range = Globals.ThisAddIn.Application.ActiveDocument.Content;
            Find find = range.Find;

            // Configurar o critério de busca
            find.Text = textoBusca;
            find.MatchWildcards = true;
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

                if (range.Start <= Globals.ThisAddIn.Application.ActiveDocument.Content.Start)
                {
                    break; // Finalizar a busca quando atingir o início do documento
                }
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

        private async void button_habilita_edicao_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Enabled = false;

            bool success = true;
            string msg_StatusBar = RibbonButton.Label + ": ";

            // Executa as tarefas em segundo plano
            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                Exclui_ContentControls(Globals.ThisAddIn.Application.Selection.Range);
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;

            // Após a execução das tarefas, atualiza a UI na Thread principal
            RibbonButton.Enabled = true;
        }
    }
}