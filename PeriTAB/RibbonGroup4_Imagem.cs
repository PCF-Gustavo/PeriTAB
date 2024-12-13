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
            object obj = System.Windows.Clipboard.GetData("FileDrop");

            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            await Tarefa.Run(() =>
            {
                // Configurações iniciais
                Stopwatch stopwatch = new Stopwatch(); if (Variables.debugging) { stopwatch.Start(); } // Inicia o cronômetro para medir o tempo de execução da Thread
                bool success = true;
                string msg_StatusBar = "";
                string msg_Falha = "";

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

                // Mensagens da Thread
                if (success) { msg_StatusBar = "Cola imagem: Sucesso"; } else { msg_StatusBar = "Cola imagem: Falha"; }
                if (Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
                {
                    stopwatch.Stop();
                    msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                }
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
                if (!success && msg_Falha != "") MessageBox.Show(msg_Falha, "Cola imagem");
            });
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

            await Tarefa.Run(() =>
            {
                // Configurações iniciais
                Stopwatch stopwatch = new Stopwatch(); if (Variables.debugging) { stopwatch.Start(); } // Inicia o cronômetro para medir o tempo de execução da Thread
                bool success = true;
                string msg_StatusBar = "";
                string msg_Falha = "";


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

                // Mensagens da Thread
                if (success) { msg_StatusBar = "Redimensiona: Sucesso"; } else { msg_StatusBar = "Redimensiona: Falha"; }
                if (Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
                {
                    stopwatch.Stop();
                    msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                }
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
                if (!success && msg_Falha != "") MessageBox.Show(msg_Falha, "Redimensiona");
            });
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

            await Tarefa.Run(() =>
            {
                // Configurações iniciais
                Stopwatch stopwatch = new Stopwatch(); if (Variables.debugging) { stopwatch.Start(); } // Inicia o cronômetro para medir o tempo de execução da Thread
                bool success = true;
                string msg_StatusBar = "";
                string msg_Falha = "";


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
                            float larguraPaginaPts = Globals.ThisAddIn.Application.ActiveDocument.PageSetup.PageWidth;
                            float margemEsquerdaPts = Globals.ThisAddIn.Application.ActiveDocument.PageSetup.LeftMargin;
                            float margemDireitaPts = Globals.ThisAddIn.Application.ActiveDocument.PageSetup.RightMargin;
                            float recuoEsquerdaPts = iShape.Range.Paragraphs[1].Format.LeftIndent;
                            float recuoDireitaPts = iShape.Range.Paragraphs[1].Format.RightIndent;
                            float primeiralinhaPts = iShape.Range.Paragraphs[1].Format.FirstLineIndent;
                            float espacoDigitavelPts = larguraPaginaPts - (margemEsquerdaPts + margemDireitaPts + recuoEsquerdaPts + recuoDireitaPts + primeiralinhaPts);
                            iShape.Width = espacoDigitavelPts;
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
                        while (((dict_InlineShape_paragraph[iParagraph])[0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines)) == 1)
                        {
                            foreach (InlineShape iShape in dict_InlineShape_paragraph[iParagraph])
                            {
                                iShape.Width *= 1.1f;
                            }
                        }
                        foreach (InlineShape iShape in dict_InlineShape_paragraph[iParagraph])
                        {
                            iShape.Width *= 0.9f;
                        }
                        while (((dict_InlineShape_paragraph[iParagraph])[0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines)) == 1)
                        {
                            foreach (InlineShape iShape in dict_InlineShape_paragraph[iParagraph])
                            {
                                iShape.Width *= 1.01f;
                            }
                        }
                        while (((dict_InlineShape_paragraph[iParagraph])[0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines)) == 2)
                        {
                            foreach (InlineShape iShape in dict_InlineShape_paragraph[iParagraph])
                            {
                                iShape.Width *= 0.99f;
                            }
                        }
                    }
                    // Verifica se o parágrafo tem mais de uma linha: caso de redução das imagens
                    if (((dict_InlineShape_paragraph[iParagraph])[0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines)) > 1)
                    {
                        // Dicionário para armazenar os tamanhos originais das imagens
                        Dictionary<InlineShape, float> tamanho_original = new Dictionary<InlineShape, float>();

                        // Armazena o tamanho original das imagens no parágrafo
                        foreach (InlineShape iShape in dict_InlineShape_paragraph[iParagraph])
                        {
                            tamanho_original[iShape] = iShape.Width;
                        }

                        // Primeira tentativa de ajustar todas as imagens
                        for (int iteration = 0; iteration < 50; iteration++)
                        {
                            // Verifica se as imagens já cabem em uma linha
                            if (((dict_InlineShape_paragraph[iParagraph])[0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines)) <= 1)
                            {
                                break; // Sai do loop se já estiver ajustado
                            }

                            // Reduz todas as imagens no parágrafo por 10% a cada iteração
                            foreach (InlineShape iShape in dict_InlineShape_paragraph[iParagraph])
                            {
                                iShape.Width *= 0.9f;
                            }
                        }

                        // Após as 50 iterações, verifica se alguma imagem ficou menor que 1 cm e se ainda ocupa mais de uma linha
                        bool algumMenorQue1cm = false;

                        foreach (InlineShape iShape in dict_InlineShape_paragraph[iParagraph])
                        {
                            if (iShape.Width < 28.35) // Verifica se a largura é menor que 1 cm em pontos
                            {
                                algumMenorQue1cm = true;
                                break; // Não precisa verificar mais
                            }
                        }

                        // Se depois de 50 tentativas ainda não couber em uma linha ou imagem ficar muito pequena, desiste de redimensionar.
                        if (((dict_InlineShape_paragraph[iParagraph])[0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines)) > 1 || algumMenorQue1cm)
                        {
                            success = false;
                            msg_Falha = "Alguma(s) imagem(ns) selecionada(s) não cabe(m) em uma única linha.";

                            // Restaura os tamanhos originais das imagens
                            foreach (InlineShape iShape in dict_InlineShape_paragraph[iParagraph])
                            {
                                iShape.Width = tamanho_original[iShape];
                            }
                        }
                        else
                        {
                            // Se as imagens couberem em uma linha, faz o ajuste fino
                            foreach (InlineShape iShape in dict_InlineShape_paragraph[iParagraph])
                            {
                                iShape.Width *= 1.1f; // Aumenta ligeiramente o tamanho

                                // Faz um ajuste final, caso precise
                                while (((dict_InlineShape_paragraph[iParagraph])[0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines)) > 1)
                                {
                                    iShape.Width *= 0.99f; // Ajusta em decrementos menores (1%)
                                }
                            }
                        }
                    }
                }

                // Mensagens da Thread
                if (success) { msg_StatusBar = "Autodimensiona: Sucesso"; } else { msg_StatusBar = "Autodimensiona: Falha"; }
                if (Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
                {
                    stopwatch.Stop();
                    msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                }
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
                if (!success && msg_Falha != "") MessageBox.Show(msg_Falha, "Autodimensiona");
            });

            // Configurações finais
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            RibbonButton.Image = Properties.Resources.redimensionar3;
            RibbonButton.Enabled = true;
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
            RibbonButton RibbonButton = (RibbonButton)sender;
            menu_inserir_imagem.Image = Properties.Resources.load_icon_png_7969;
            menu_inserir_imagem.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            await Tarefa.Run(() =>
            {
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.Line.Visible = MsoTriState.msoTrue;
                        ishape.Line.Weight = (float)0.5;
                        ishape.Line.ForeColor.RGB = Color.FromArgb(0, 0, 0).ToArgb();
                    }
                }
            });

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            menu_inserir_imagem.Image = Properties.Resources._;
            menu_inserir_imagem.Enabled = true;
        }

        private async void button_borda_vermelha_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            menu_inserir_imagem.Image = Properties.Resources.load_icon_png_7969;
            menu_inserir_imagem.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            await Tarefa.Run(() =>
            {
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.Line.Visible = MsoTriState.msoTrue;
                        ishape.Line.Weight = 2;
                        ishape.Line.ForeColor.RGB = Color.FromArgb(0, 0, 255).ToArgb();
                    }
                }
            });

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            menu_inserir_imagem.Image = Properties.Resources._;
            menu_inserir_imagem.Enabled = true;
        }

        private async void button_borda_amarela_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            menu_inserir_imagem.Image = Properties.Resources.load_icon_png_7969;
            menu_inserir_imagem.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            await Tarefa.Run(() =>
            {
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.Line.Visible = MsoTriState.msoTrue;
                        ishape.Line.Weight = 3;
                        ishape.Line.ForeColor.RGB = Color.FromArgb(0, 255, 255).ToArgb();
                    }
                }
            });

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            menu_inserir_imagem.Image = Properties.Resources._;
            menu_inserir_imagem.Enabled = true;
        }

        private async void button_legenda_imagem_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            menu_inserir_imagem.Image = Properties.Resources.load_icon_png_7969;
            menu_inserir_imagem.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            Range Selecao_inicial = Globals.ThisAddIn.Application.Selection.Range; //Salva a seleção inicial

            await Tarefa.Run(() =>
            {
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
            });

            Selecao_inicial.Select(); // Restaura a seleção inicial
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            menu_inserir_imagem.Image = Properties.Resources._;
            menu_inserir_imagem.Enabled = true;
        }

        private async void button_remove_borda_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            menu_remover_imagem.Image = Properties.Resources.load_icon_png_7969;
            menu_remover_imagem.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            await Tarefa.Run(() =>
            {
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.Line.Visible = MsoTriState.msoFalse;
                    }
                }
            });

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            menu_remover_imagem.Image = Properties.Resources.x;
            menu_remover_imagem.Enabled = true;
        }

        private async void button_remove_formatacao_Click_1(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            menu_remover_imagem.Image = Properties.Resources.load_icon_png_7969;
            menu_remover_imagem.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            await Tarefa.Run(() =>
            {
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.Reset();
                    }
                }
            });

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            menu_remover_imagem.Image = Properties.Resources.x;
            menu_remover_imagem.Enabled = true;
        }

        private async void button_remove_forma_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            menu_remover_imagem.Image = Properties.Resources.load_icon_png_7969;
            menu_remover_imagem.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            await Tarefa.Run(() =>
            {
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
            });

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            menu_remover_imagem.Image = Properties.Resources.x;
            menu_remover_imagem.Enabled = true;
        }

        private async void button_remove_texto_alt_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            menu_remover_imagem.Image = Properties.Resources.load_icon_png_7969;
            menu_remover_imagem.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            await Tarefa.Run(() =>
            {
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.AlternativeText = "";
                    }
                }
            });
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            menu_remover_imagem.Image = Properties.Resources.x;
            menu_remover_imagem.Enabled = true;
        }

        private void button_alinha_legenda_figuras_Click(object sender, RibbonControlEventArgs e)
        { }

        private async void button_remove_imagem_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            menu_remover_imagem.Image = Properties.Resources.load_icon_png_7969;
            menu_remover_imagem.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            await Tarefa.Run(() =>
            {
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
            });

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            menu_remover_imagem.Image = Properties.Resources.x;
            menu_remover_imagem.Enabled = true;
        }
    }
}