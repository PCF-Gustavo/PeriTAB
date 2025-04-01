using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Tarefa = System.Threading.Tasks.Task;


namespace PeriTAB
{
    public partial class Ribbon
    {
        private async void button_cola_imagem_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
            object obj = System.Windows.Clipboard.GetData("FileDrop");

            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            // Configurações iniciais
            bool success = true;
            string msg_StatusBar = RibbonButton.Label + ": ";
            string msg_Falha = "";

            await Tarefa.Run(() =>
            {
                if (System.Windows.Clipboard.ContainsData("FileDrop"))
                {
                    string[] pathfile = (string[])obj;
                    string[] pathfile2 = { "" };
                    string[] pathfile3 = { "" };
                    int n = 0;
                    for (int i = 0; i <= pathfile.Length - 1; i++)
                    {
                        if (File.Exists(pathfile[i]))
                        {
                            string extensao = (pathfile[i].Substring(pathfile[i].Length - 4)).ToLower();
                            if (extensao == ".jpg" | extensao == "jpeg" | extensao == ".png" | extensao == ".bmp" | extensao == ".gif" | extensao == "tiff") //Se tem extensao de imagem
                            {
                                Array.Resize(ref pathfile2, n + 1);
                                pathfile2[n] = pathfile[i];
                                n++;
                            }
                        }
                    }

                    if (pathfile2[0] != "")
                    {
                        // Evita a exclusao de \r (Carrige Return) ao final da seleção
                        if (Globals.ThisAddIn.Application.Selection.Text.EndsWith("\r") && Globals.ThisAddIn.Application.Selection.InlineShapes.Count > 0)
                        {
                            Globals.ThisAddIn.Application.Selection.MoveEnd(WdUnits.wdCharacter, -1);
                        }
                        for (int i = 0; i <= pathfile2.Length - 1; i++)
                        {

                            bool link = false; bool save = true;
                            if (Globals.Ribbons.Ribbon.checkBox_referencia.Checked == true) { link = true; save = false; }
                            InlineShape imagem = Globals.ThisAddIn.Application.Selection.InlineShapes.AddPicture(pathfile2[i], link, save);
                            imagem.LockAspectRatio = MsoTriState.msoTrue;
                            if (checkBox_largura.Checked)
                            {
                                string larg_string = Globals.Ribbons.Ribbon.editBox_largura.Text;
                                float.TryParse(larg_string, out float larg);
                                imagem.Width = Globals.ThisAddIn.Application.CentimetersToPoints(larg);
                            }

                            if (checkBox_altura.Checked)
                            {
                                string alt_string = Globals.Ribbons.Ribbon.editBox_altura.Text;
                                float.TryParse(alt_string, out float alt);
                                imagem.Height = Globals.ThisAddIn.Application.CentimetersToPoints(alt);
                            }

                            if (i != pathfile2.Length - 1) //Exceto última imagem
                            {

                                switch (dropDown_separador.SelectedItem.Label) //Insere separador
                                {
                                    case "Espaço":
                                        Globals.ThisAddIn.Application.Selection.InsertAfter(" ");
                                        Globals.ThisAddIn.Application.Selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                                        break;
                                    case "Parágrafo":
                                        Globals.ThisAddIn.Application.Selection.InsertAfter(System.Environment.NewLine);
                                        Globals.ThisAddIn.Application.Selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                                        break;
                                    case "Parágrafo + 3pt":
                                        Globals.ThisAddIn.Application.Selection.ParagraphFormat.SpaceAfter = 3;
                                        Globals.ThisAddIn.Application.Selection.InsertAfter(System.Environment.NewLine);
                                        Globals.ThisAddIn.Application.Selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                                        break;
                                }
                            }
                        }

                        // Seleção das imagens ao final da colagem
                        if (dropDown_separador.SelectedItem.Label == "Nenhum")
                        {
                            int L = pathfile2.Length;
                            Globals.ThisAddIn.Application.Selection.MoveEnd(WdUnits.wdCharacter, -L);
                            Globals.ThisAddIn.Application.Selection.MoveRight(WdUnits.wdCharacter, L, WdMovementType.wdExtend);
                        }
                        else
                        {
                            int L = pathfile2.Length;
                            Globals.ThisAddIn.Application.Selection.MoveEnd(WdUnits.wdCharacter, -(2 * L - 1));
                            Globals.ThisAddIn.Application.Selection.MoveRight(WdUnits.wdCharacter, 2 * L - 1, WdMovementType.wdExtend);
                        }
                    }
                    else
                    {
                        success = false;
                        msg_Falha = "Não há imagens no Clipboard.";
                    }
                }
                else
                {
                    success = false;
                    msg_Falha = "Não há imagens no Clipboard.";
                }
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();

            });

            // Mensagens da Thread
            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            if (!success && msg_Falha != "") MessageBox.Show(msg_Falha, RibbonButton.Label);

            // Configurações finais
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            RibbonButton.Image = Properties.Resources.image_icon;
            RibbonButton.Enabled = true;
        }

        public class Comparer_Windows_order : IComparer<string> /*implement an IComparer to get the same sort behavior as Windows Explorer*/
        {

            [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
            static extern int StrCmpLogicalW(String x, String y);

            public int Compare(string x, string y)
            {
                return StrCmpLogicalW(x, y);
            }
        }

        private void checkBox_largura_Click(object sender, RibbonControlEventArgs e)
        {
            if (!checkBox_largura.Checked) checkBox_largura.Checked = true;
            if (checkBox_largura.Checked)
            {
                checkBox_altura.Checked = false;
                editBox_altura.Enabled = false;
                editBox_altura.Text = "";
                editBox_largura.Enabled = true;
                editBox_largura.Text = Class_RibbonControls.GetPreference("largura");
            }
        }

        private void checkBox_altura_Click(object sender, RibbonControlEventArgs e)
        {
            if (!checkBox_altura.Checked) checkBox_altura.Checked = true;
            if (checkBox_altura.Checked)
            {
                checkBox_largura.Checked = false;
                editBox_largura.Enabled = false;
                editBox_largura.Text = "";
                editBox_altura.Enabled = true;
                editBox_altura.Text = Class_RibbonControls.GetPreference("altura");
            }
        }

        private void checkBox_referencia_Click(object sender, RibbonControlEventArgs e)
        {
            //if (checkBox_referencia.Checked)
            //{
            //    System.Windows.Forms.MessageBox.Show("Cuidado! Excluir/mover/renomear o arquivo da imagem causará perda de referência.","Referência");
            //}
        }

        private void editBox_largura_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (float.TryParse(editBox_largura.Text, out float larg) & larg.ToString() == editBox_largura.Text & larg >= 0.1 & larg < 100)
            {
                Class_RibbonControls.ChangePreference("largura", editBox_largura.Text);
            }
            else
            {
                editBox_largura.Text = Class_RibbonControls.GetPreference("largura");
            }
        }

        private void editBox_altura_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (float.TryParse(editBox_altura.Text, out float alt) & alt.ToString() == editBox_altura.Text & alt >= 0.1 & alt < 100)
            {
                Class_RibbonControls.ChangePreference("altura", editBox_altura.Text);
            }
            else
            {
                editBox_altura.Text = Class_RibbonControls.GetPreference("altura"); ;
            }
        }

        private async void button_redimensiona_imagem_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            // Configurações iniciais
            bool success = true;
            string msg_StatusBar = RibbonButton.Label + ": ";
            string msg_Falha = "";

            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                if (Globals.ThisAddIn.Application.Selection.InlineShapes.Count < 1)
                {
                    success = false;
                    msg_Falha = "Não há imagens selecionadas.";
                }

                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        InlineShape imagem = ishape;
                        imagem.LockAspectRatio = MsoTriState.msoTrue;

                        if (checkBox_largura.Checked)
                        {
                            string larg_string = Globals.Ribbons.Ribbon.editBox_largura.Text;
                            float.TryParse(larg_string, out float larg);
                            imagem.Width = Globals.ThisAddIn.Application.CentimetersToPoints(larg);
                        }

                        if (checkBox_altura.Checked)
                        {
                            string alt_string = Globals.Ribbons.Ribbon.editBox_altura.Text;
                            float.TryParse(alt_string, out float alt);
                            imagem.Height = Globals.ThisAddIn.Application.CentimetersToPoints(alt);
                        }
                    }
                }
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });
            // Mensagens da Thread
            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            if (!success && msg_Falha != "") MessageBox.Show(msg_Falha, RibbonButton.Label);

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            RibbonButton.Image = Properties.Resources.redimensionar2;
            RibbonButton.Enabled = true;
        }

        private async void button_autodimensiona_imagem_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            // Configurações iniciais
            bool success = true;
            string msg_StatusBar = RibbonButton.Label + ": ";
            string msg_Falha = "";

            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                if (Globals.ThisAddIn.Application.Selection.InlineShapes.Count < 1)
                {
                    success = false;
                    msg_Falha = "Não há imagens selecionadas.";
                }

                Dictionary<int, List<InlineShape>> dict_InlineShape_paragraph = new Dictionary<int, List<InlineShape>>();
                foreach (InlineShape iShape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    // Verifica se o parágrafo contém mais de uma InlineShape
                    if (iShape.Range.Paragraphs[1].Range.InlineShapes.Count > 1)
                    {
                        int num_Paragraph = 0;
                        if (iShape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | iShape.Type == WdInlineShapeType.wdInlineShapePicture)
                        {
                            Paragraph iParagraph = iShape.Range.Paragraphs.First;
                            for (int i = 1; i <= Globals.ThisAddIn.Application.ActiveDocument.Paragraphs.Count; i++)
                            {
                                if (Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[i].Range.Start == iParagraph.Range.Start)
                                {
                                    num_Paragraph = i;
                                    break;
                                }
                            }
                        }
                        // Verifica se o dicionário já contém o parágrafo
                        if (!dict_InlineShape_paragraph.ContainsKey(num_Paragraph))
                        {
                            // Se não contém, cria uma nova lista de InlineShapes para esse parágrafo
                            dict_InlineShape_paragraph[num_Paragraph] = new List<InlineShape>();
                        }
                        // Adiciona a InlineShape à lista correspondente ao parágrafo
                        dict_InlineShape_paragraph[num_Paragraph].Add(iShape);
                    }
                    else
                    {
                        if (!(iShape.Range.Paragraphs[1].Range.Information[WdInformation.wdWithInTable]))
                        {
                            int paginaInicial = iShape.Range.Information[WdInformation.wdActiveEndPageNumber];
                            float larguraPaginaPts = Globals.ThisAddIn.Application.ActiveDocument.PageSetup.PageWidth;
                            float margemEsquerdaPts = Globals.ThisAddIn.Application.ActiveDocument.PageSetup.LeftMargin;
                            float margemDireitaPts = Globals.ThisAddIn.Application.ActiveDocument.PageSetup.RightMargin;
                            float recuoEsquerdaPts = iShape.Range.Paragraphs[1].Format.LeftIndent;
                            float recuoDireitaPts = iShape.Range.Paragraphs[1].Format.RightIndent;
                            float primeiralinhaPts = iShape.Range.Paragraphs[1].Format.FirstLineIndent;
                            float espacoDigitavelPts = larguraPaginaPts - (margemEsquerdaPts + margemDireitaPts + recuoEsquerdaPts + recuoDireitaPts + primeiralinhaPts);
                            iShape.Width = espacoDigitavelPts;

                            // Reduzir imagem caso ultrapasse a página
                            float tamanhoOriginal = iShape.Width;
                            if (iShape.Range.Information[WdInformation.wdActiveEndPageNumber] > paginaInicial)
                            {
                                float minScale = 0.01f; // Escala mínima (1% do tamanho atual)
                                float maxScale = 1f;  // Escala máxima (100% do tamanho atual)
                                float tolerance = 0.001f; // Tolerância para encerrar a busca binária
                                while (maxScale - minScale > tolerance)
                                {
                                    float midScale = (minScale + maxScale) / 2;

                                    iShape.Width = tamanhoOriginal * midScale;

                                    if (iShape.Range.Information[WdInformation.wdActiveEndPageNumber] > paginaInicial)
                                    {
                                        maxScale = midScale;
                                    }
                                    else
                                    {
                                        minScale = midScale;
                                    }
                                }
                                iShape.Width = tamanhoOriginal * minScale;
                            }
                            else
                            {
                                iShape.Width = tamanhoOriginal;
                            }
                        }
                        else { success = false; }
                    }
                }
                // Itera por cada parágrafo que contém múltiplas InlineShapes
                foreach (var iParagraph in dict_InlineShape_paragraph.Keys)
                {
                    // Verifica se o parágrafo tem exatamente uma linha: caso de aumento das imagens
                    if (((dict_InlineShape_paragraph[iParagraph])[0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines)) == 0) { success = false; } //Se está dentro da tabela, o numero de linhas do paragrafo é zero
                    if (((dict_InlineShape_paragraph[iParagraph])[0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines)) == 1)
                    {
                        Redimenionar_imagens_por_busca_binaria(dict_InlineShape_paragraph[iParagraph], true);
                    }
                    // Verifica se o parágrafo tem mais de uma linha: caso de redução das imagens
                    if (((dict_InlineShape_paragraph[iParagraph])[0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines)) > 1)
                    {
                        // Salva os tamanhos originais das imagens
                        Dictionary<InlineShape, float> tamanhosOriginais = new Dictionary<InlineShape, float>();
                        foreach (InlineShape imagem in dict_InlineShape_paragraph[iParagraph])
                        {
                            tamanhosOriginais[imagem] = imagem.Width;
                        }

                        // Aplica o fator minScale a todas as imagens
                        foreach (InlineShape imagem in dict_InlineShape_paragraph[iParagraph])
                        {
                            imagem.Width = tamanhosOriginais[imagem] * 0.01f;
                        }

                        // Verifica o número de linhas do parágrafo
                        int numLinhas = dict_InlineShape_paragraph[iParagraph][0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines);

                        foreach (InlineShape imagem in dict_InlineShape_paragraph[iParagraph])
                        {
                            imagem.Width = tamanhosOriginais[imagem];
                        }

                        // Se o parágrafo ainda ocupa mais de uma linha, dar erro
                        if (numLinhas > 1)
                        {
                            success = false;
                            msg_Falha = "Alguma(s) imagem(ns) selecionada(s) não cabe(m) em uma única linha.";
                        }
                        else
                        {
                            Redimenionar_imagens_por_busca_binaria(dict_InlineShape_paragraph[iParagraph], false);
                        }
                    }
                }
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            // Mensagens da Thread
            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            if (!success && msg_Falha != "") MessageBox.Show(msg_Falha, RibbonButton.Label);

            // Configurações finais
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            RibbonButton.Image = Properties.Resources.redimensionar3;
            RibbonButton.Enabled = true;
        }
        void Redimenionar_imagens_por_busca_binaria(List<InlineShape> imagens, bool fit_to_page)
        {
            float minScale = 0.01f; // Escala mínima (1% do tamanho atual)
            float maxScale = 20f;  // Escala máxima (2000% do tamanho atual)
            float tolerance = 0.001f; // Tolerância para encerrar a busca binária

            // Salva os tamanhos originais para poder reverter caso necessário
            Dictionary<InlineShape, float> tamanhosOriginais = new Dictionary<InlineShape, float>();
            foreach (InlineShape imagem in imagens)
            {
                tamanhosOriginais[imagem] = imagem.Width;
            }

            // Variável para armazenar a página inicial (usada somente se fit_to_page for true)
            int paginaInicial = fit_to_page ? imagens[0].Range.Information[WdInformation.wdActiveEndPageNumber] : -1;

            while (maxScale - minScale > tolerance)
            {
                float midScale = (minScale + maxScale) / 2;

                // Aplica o fator de escala às imagens
                foreach (InlineShape imagem in imagens)
                {
                    imagem.Width = tamanhosOriginais[imagem] * midScale;
                }

                // Verifica o número de linhas ocupadas pelo parágrafo
                int numLinhas = imagens[0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines);

                // Verifica a página atual se fit_to_page for true
                bool mudouDePagina = fit_to_page && imagens[0].Range.Information[WdInformation.wdActiveEndPageNumber] != paginaInicial;

                if (numLinhas > 1 || mudouDePagina && midScale > 1)
                {
                    // Se ainda ocupa mais de uma linha ou mudou de página, diminui o tamanho
                    maxScale = midScale;
                }
                else
                {
                    // Se cabe em uma linha e está na mesma página, tenta aumentar o tamanho
                    minScale = midScale;
                }
            }

            // Ajuste final usando o menor fator encontrado
            foreach (InlineShape imagem in imagens)
            {
                imagem.Width = tamanhosOriginais[imagem] * minScale;
            }
        }


        private bool IsLastShapeInParagraph(InlineShape ishape)
        {
            InlineShape lastShape = null;
            foreach (InlineShape inlineShape in ishape.Range.Paragraphs[1].Range.InlineShapes)
            {
                if (inlineShape.Type == WdInlineShapeType.wdInlineShapePicture)
                {
                    lastShape = inlineShape;
                }
            }
            if (lastShape != null && ishape.Range.Information[WdInformation.wdHorizontalPositionRelativeToTextBoundary] >= lastShape.Range.Information[WdInformation.wdHorizontalPositionRelativeToTextBoundary] && ishape.Range.Information[WdInformation.wdVerticalPositionRelativeToTextBoundary] >= lastShape.Range.Information[WdInformation.wdVerticalPositionRelativeToTextBoundary])
            {
                return true;
            }
            return false;
        }

        private async void button_borda_preta_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            menu_inserir_imagem.Image = Properties.Resources.load_icon_png_7969;
            menu_inserir_imagem.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            bool success = true;
            string msg_StatusBar = ((RibbonButton)sender).Label + ": ";

            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.Line.Visible = MsoTriState.msoTrue;
                        ishape.Line.Weight = (float)0.5;
                        ishape.Line.ForeColor.RGB = Color.FromArgb(0, 0, 0).ToArgb();
                    }
                }
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            menu_inserir_imagem.Image = Properties.Resources._;
            menu_inserir_imagem.Enabled = true;
        }

        private async void button_borda_vermelha_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            menu_inserir_imagem.Image = Properties.Resources.load_icon_png_7969;
            menu_inserir_imagem.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            bool success = true;
            string msg_StatusBar = ((RibbonButton)sender).Label + ": ";

            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.Line.Visible = MsoTriState.msoTrue;
                        ishape.Line.Weight = 2;
                        ishape.Line.ForeColor.RGB = Color.FromArgb(0, 0, 255).ToArgb();
                    }
                }
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            menu_inserir_imagem.Image = Properties.Resources._;
            menu_inserir_imagem.Enabled = true;
        }

        private async void button_borda_amarela_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            menu_inserir_imagem.Image = Properties.Resources.load_icon_png_7969;
            menu_inserir_imagem.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            bool success = true;
            string msg_StatusBar = ((RibbonButton)sender).Label + ": ";

            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.Line.Visible = MsoTriState.msoTrue;
                        ishape.Line.Weight = 3;
                        ishape.Line.ForeColor.RGB = Color.FromArgb(0, 255, 255).ToArgb();
                    }
                }
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            menu_inserir_imagem.Image = Properties.Resources._;
            menu_inserir_imagem.Enabled = true;
        }

        private async void button_legenda_imagem_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            menu_inserir_imagem.Image = Properties.Resources.load_icon_png_7969;
            menu_inserir_imagem.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            bool success = true;
            string msg_StatusBar = ((RibbonButton)sender).Label + ": ";

            Range Selecao_inicial = Globals.ThisAddIn.Application.Selection.Range; //Salva a seleção inicial

            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                string estilo_nome_baseado = "Legenda";
                Globals.ThisAddIn.Application.OrganizerCopy(PeriTAB.Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);

                List<InlineShape> list_InlineShape = new List<InlineShape>();

                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        list_InlineShape.Add(ishape);
                    }
                }
                foreach (InlineShape ishape in list_InlineShape)
                {
                    ishape.Select();

                    if (Globals.ThisAddIn.Application.Selection.Paragraphs[1].Next() != null)
                    {
                        if (Globals.ThisAddIn.Application.Selection.Paragraphs[1].Next().Range.Characters.Count >= 7)
                        {
                            if (Globals.ThisAddIn.Application.Selection.Paragraphs[1].Next().Range.Text.Substring(0, 7) == "Figura ") continue;
                        }
                    }
                    if (IsLastShapeInParagraph(ishape))
                    {
                        bool label_existe = false;
                        foreach (CaptionLabel label in Globals.ThisAddIn.Application.CaptionLabels)
                        {
                            if (label.Name == "Figura") { label_existe = true; }
                        }
                        if (!label_existe) { Globals.ThisAddIn.Application.CaptionLabels.Add("Figura"); }

                        Globals.ThisAddIn.Application.Selection.InsertCaption(Label: "Figura", Title: " " + ((char)8211).ToString(), TitleAutoText: "", Position: WdCaptionPosition.wdCaptionPositionBelow, ExcludeLabel: 0);
                        Globals.ThisAddIn.Application.Selection.set_Style((object)"12 - Legendas de Figuras (PeriTAB)");
                        Globals.ThisAddIn.Application.Selection.InsertAfter(" ");
                        Globals.ThisAddIn.Application.Run("alinha_legenda");
                    }
                }
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;

            Selecao_inicial.Select(); // Restaura a seleção inicial
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            menu_inserir_imagem.Image = Properties.Resources._;
            menu_inserir_imagem.Enabled = true;
        }

        private async void button_remove_borda_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            menu_remover_imagem.Image = Properties.Resources.load_icon_png_7969;
            menu_remover_imagem.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            bool success = true;
            string msg_StatusBar = ((RibbonButton)sender).Label + ": ";

            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.Line.Visible = MsoTriState.msoFalse;
                    }
                }
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            menu_remover_imagem.Image = Properties.Resources.x;
            menu_remover_imagem.Enabled = true;
        }

        private async void button_remove_formatacao_Click_1(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            menu_remover_imagem.Image = Properties.Resources.load_icon_png_7969;
            menu_remover_imagem.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            bool success = true;
            string msg_StatusBar = ((RibbonButton)sender).Label + ": ";

            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.Reset();
                    }
                }
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            menu_remover_imagem.Image = Properties.Resources.x;
            menu_remover_imagem.Enabled = true;
        }

        private async void button_remove_forma_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            menu_remover_imagem.Image = Properties.Resources.load_icon_png_7969;
            menu_remover_imagem.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            bool success = true;
            string msg_StatusBar = ((RibbonButton)sender).Label + ": ";

            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                List<Microsoft.Office.Interop.Word.Shape> listaShapes = new List<Microsoft.Office.Interop.Word.Shape>();
                foreach (Microsoft.Office.Interop.Word.Shape ishape in Globals.ThisAddIn.Application.Selection.Range.ShapeRange)
                {
                    if (ishape.Type == MsoShapeType.msoAutoShape | ishape.Type == MsoShapeType.msoFreeform | ishape.Type == MsoShapeType.msoLine | ishape.Type == MsoShapeType.msoTextBox)
                    {
                        listaShapes.Add(ishape);
                    }
                }
                foreach (Microsoft.Office.Interop.Word.Shape ishape in listaShapes)
                {
                    ishape.Delete();
                }
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            menu_remover_imagem.Image = Properties.Resources.x;
            menu_remover_imagem.Enabled = true;
        }

        private async void button_remove_texto_alt_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            menu_remover_imagem.Image = Properties.Resources.load_icon_png_7969;
            menu_remover_imagem.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            bool success = true;
            string msg_StatusBar = ((RibbonButton)sender).Label + ": ";

            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.AlternativeText = "";
                    }
                }
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            menu_remover_imagem.Image = Properties.Resources.x;
            menu_remover_imagem.Enabled = true;
        }

        private void button_alinha_legenda_figuras_Click(object sender, RibbonControlEventArgs e)
        {
        }

        private async void button_remove_imagem_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            menu_remover_imagem.Image = Properties.Resources.load_icon_png_7969;
            menu_remover_imagem.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            bool success = true;
            string msg_StatusBar = ((RibbonButton)sender).Label + ": ";

            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                List<InlineShape> listaShapes = new List<InlineShape>();
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        listaShapes.Add(ishape);
                    }
                }
                foreach (InlineShape ishape in listaShapes)
                {
                    ishape.Delete();
                }
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            menu_remover_imagem.Image = Properties.Resources.x;
            menu_remover_imagem.Enabled = true;
        }
    }
}