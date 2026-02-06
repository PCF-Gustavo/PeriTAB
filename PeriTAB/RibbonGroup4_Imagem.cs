using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;


namespace PeriTAB
{
    public partial class Ribbon
    {

        private async void Button_cola_imagem_Click(object sender, RibbonControlEventArgs e)
        {
            object obj = System.Windows.Clipboard.GetData("FileDrop");
            await Executar_Ribbon_com_UI_responsiva(sender, e, async progress =>
            {
                if (System.Windows.Clipboard.ContainsData("FileDrop"))
                {
                    string[] pathfile = (string[])obj;
                    string[] pathfile2 = { "" };
                    string[] pathfile3 = { "" };
                    int n = 0;
                    for (int i = 0; i <= pathfile.Length - 1; i++)
                    {
                        await progress.Tick_50ms((int)((i * 2) / pathfile.Length));

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
                        Array.Sort(pathfile2, new Comparer_Windows_order());

                        Microsoft.Office.Interop.Word.Shape Tela_de_desenho = Apenas_TelaDeDesenho_Selecionada();
                        if (Tela_de_desenho == null)
                        {
                            // Evita a exclusao de \r (Carrige Return) ao final da seleção
                            if (Globals.ThisAddIn.Application.Selection.Text != null)
                            {
                                if (Globals.ThisAddIn.Application.Selection.Text.EndsWith("\r") && Globals.ThisAddIn.Application.Selection.InlineShapes.Count > 0)
                                {
                                    Globals.ThisAddIn.Application.Selection.MoveEnd(WdUnits.wdCharacter, -1);
                                }
                            }
                            for (int i = 0; i <= pathfile2.Length - 1; i++)
                            {
                                await progress.Tick_50ms((int)((i * 10) / pathfile2.Length));

                                bool link = false; bool save = true;
                                if (Globals.Ribbons.Ribbon.checkBox_referencia.Checked == true) { link = true; save = false; }
                                InlineShape imagem = Globals.ThisAddIn.Application.Selection.InlineShapes.AddPicture(pathfile2[i], link, save);
                                imagem.LockAspectRatio = MsoTriState.msoTrue;
                                if (CheckBox_largura.Checked)
                                {
                                    string larg_string = Globals.Ribbons.Ribbon.EditBox_largura.Text;
                                    float.TryParse(larg_string, out float larg);
                                    imagem.Width = Globals.ThisAddIn.Application.CentimetersToPoints(larg);
                                }

                                if (CheckBox_altura.Checked)
                                {
                                    string alt_string = Globals.Ribbons.Ribbon.EditBox_altura.Text;
                                    float.TryParse(alt_string, out float alt);
                                    imagem.Height = Globals.ThisAddIn.Application.CentimetersToPoints(alt);
                                }

                                if (i != pathfile2.Length - 1) //Exceto última imagem
                                {

                                    switch (DropDown_separador.SelectedItem.Label) //Insere separador
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
                            if (DropDown_separador.SelectedItem.Label == "Nenhum")
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
                            // Para TELA DE DESENHO
                            for (int i = 0; i <= pathfile2.Length - 1; i++)
                            {
                                await progress.Tick_50ms((int)((i * 10) / pathfile2.Length));
                                using (Image imagem = Image.FromFile(pathfile2[i]))
                                {
                                    float largura = 0;
                                    float altura = 0;

                                    if (CheckBox_largura.Checked)
                                    {
                                        float.TryParse(Globals.Ribbons.Ribbon.EditBox_largura.Text, out largura);
                                        altura = largura * imagem.Height / imagem.Width;
                                    }
                                    if (CheckBox_altura.Checked)
                                    {
                                        float.TryParse(Globals.Ribbons.Ribbon.EditBox_altura.Text, out altura);
                                        largura = altura * imagem.Width / imagem.Height;
                                    }
                                    float largura_pontos = Globals.ThisAddIn.Application.CentimetersToPoints(largura);
                                    float altura_pontos = Globals.ThisAddIn.Application.CentimetersToPoints(altura);
                                    if (Tela_de_desenho.Width >= largura_pontos && Tela_de_desenho.Height >= altura_pontos)
                                    {
                                        Microsoft.Office.Interop.Word.Shape shape_image =
                                        Tela_de_desenho.CanvasItems.AddPicture(
                                            FileName: pathfile2[i],
                                            LinkToFile: false,
                                            SaveWithDocument: true,
                                            Width: largura_pontos,
                                            Height: altura_pontos
                                            );
                                        shape_image.Select(false); // Seleção das imagens ao final da colagem
                                    }
                                    else throw new Exception("As dimensões da imagem excedem o tamanho da tela de desenho.");
                                }
                            }
                        }
                    }
                    else throw new Exception("Não há imagens no Clipboard.");
                }
                else throw new Exception("Não há imagens no Clipboard.");
            }, barra_de_progresso: true, desabilitar_ScreenUpdating: true);
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

        private Microsoft.Office.Interop.Word.Shape Apenas_TelaDeDesenho_Selecionada()
        {
            Selection selecao = Globals.ThisAddIn.Application.Selection;

            // Verifica se a seleção é apenas de uma forma
            if (selecao.Type == WdSelectionType.wdSelectionInlineShape &&
                selecao.ShapeRange != null &&
                selecao.ShapeRange.Count == 1)
            {
                Microsoft.Office.Interop.Word.Shape shape = selecao.ShapeRange[1];

                // Verifica se a forma é uma tela de desenho (canvas)
                if (shape.Type == MsoShapeType.msoCanvas)
                {
                    return shape;
                }
            }

            // Se não for uma seleção exclusivamente de uma tela de desenho
            return null;
        }

        private void CheckBox_largura_Click(object sender, RibbonControlEventArgs e)
        {
            if (!CheckBox_largura.Checked) CheckBox_largura.Checked = true;
            if (CheckBox_largura.Checked)
            {
                CheckBox_altura.Checked = false;
                EditBox_altura.Enabled = false;
                EditBox_altura.Text = "";
                EditBox_largura.Enabled = true;
                EditBox_largura.Text = Class_RibbonControls.Retorna_preferencia("largura");
            }
        }

        private void CheckBox_altura_Click(object sender, RibbonControlEventArgs e)
        {
            if (!CheckBox_altura.Checked) CheckBox_altura.Checked = true;
            if (CheckBox_altura.Checked)
            {
                CheckBox_largura.Checked = false;
                EditBox_largura.Enabled = false;
                EditBox_largura.Text = "";
                EditBox_altura.Enabled = true;
                EditBox_altura.Text = Class_RibbonControls.Retorna_preferencia("altura");
            }
        }

        private void CheckBox_referencia_Click(object sender, RibbonControlEventArgs e)
        {
            //if (checkBox_referencia.Checked)
            //{
            //    System.Windows.Forms.MessageBox.Show("Cuidado! Excluir/mover/renomear o arquivo da imagem causará perda de referência.","Referência");
            //}
        }

        private void EditBox_largura_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (float.TryParse(EditBox_largura.Text, out float larg) & larg.ToString() == EditBox_largura.Text & larg >= 0.1 & larg < 100)
            {
                Class_RibbonControls.Muda_preferencia("largura", EditBox_largura.Text);
            }
            else
            {
                EditBox_largura.Text = Class_RibbonControls.Retorna_preferencia("largura");
            }
        }

        private void EditBox_altura_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (float.TryParse(EditBox_altura.Text, out float alt) & alt.ToString() == EditBox_altura.Text & alt >= 0.1 & alt < 100)
            {
                Class_RibbonControls.Muda_preferencia("altura", EditBox_altura.Text);
            }
            else
            {
                EditBox_altura.Text = Class_RibbonControls.Retorna_preferencia("altura"); ;
            }
        }

        private async void Button_redimensiona_imagem_Click(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon_com_UI_responsiva(sender, e, async progress =>
            {
                if (Globals.ThisAddIn.Application.Selection.InlineShapes.Count < 1) throw new Exception("Não há imagens selecionadas.");

                int i = 0;
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    await progress.Tick_50ms((int)((i * 10) / Globals.ThisAddIn.Application.Selection.InlineShapes.Count));
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        InlineShape imagem = ishape;
                        imagem.LockAspectRatio = MsoTriState.msoTrue;

                        if (CheckBox_largura.Checked)
                        {
                            string larg_string = Globals.Ribbons.Ribbon.EditBox_largura.Text;
                            float.TryParse(larg_string, out float larg);
                            imagem.Width = Globals.ThisAddIn.Application.CentimetersToPoints(larg);
                        }

                        if (CheckBox_altura.Checked)
                        {
                            string alt_string = Globals.Ribbons.Ribbon.EditBox_altura.Text;
                            float.TryParse(alt_string, out float alt);
                            imagem.Height = Globals.ThisAddIn.Application.CentimetersToPoints(alt);
                        }
                    }
                }
            }, desabilitar_ScreenUpdating: true);
        }

        private async void Button_autodimensiona_imagem_Click(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon_com_UI_responsiva(sender, e, async progress =>
            {
                if (Globals.ThisAddIn.Application.Selection.InlineShapes.Count < 1) throw new Exception("Não há imagens selecionadas.");

                Dictionary<int, List<InlineShape>> dict_InlineShape_paragraph = new Dictionary<int, List<InlineShape>>();
                int j = 0;
                foreach (InlineShape iShape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    await progress.Tick_50ms((int)((j * 10) / Globals.ThisAddIn.Application.Selection.InlineShapes.Count));

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
                        else { throw new Exception(""); }
                    }
                }
                // Itera por cada parágrafo que contém múltiplas InlineShapes
                j = 0;
                foreach (var iParagraph in dict_InlineShape_paragraph.Keys)
                {
                    await progress.Tick_50ms((int)((j * 10) / Globals.ThisAddIn.Application.Selection.InlineShapes.Count));
                    // Verifica se o parágrafo tem exatamente uma linha: caso de aumento das imagens
                    if (((dict_InlineShape_paragraph[iParagraph])[0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines)) == 0) { throw new Exception(""); } //Se está dentro da tabela, o numero de linhas do paragrafo é zero
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
                            throw new Exception("Alguma(s) imagem(ns) selecionada(s) não cabe(m) em uma única linha.");
                        }
                        else
                        {
                            Redimenionar_imagens_por_busca_binaria(dict_InlineShape_paragraph[iParagraph], false);
                        }
                    }
                }
            }, barra_de_progresso: true, desabilitar_ScreenUpdating: true);
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

            // Variável para armazenar a página inicial
            int paginaInicial = fit_to_page ? imagens[0].Range.Information[WdInformation.wdActiveEndPageNumber] : -1;

            while (maxScale - minScale > tolerance)
            {
                float midScale = (minScale + maxScale) / 2;

                // Aplica o fator de escala às imagens
                foreach (InlineShape imagem in imagens)
                {
                    imagem.Width = tamanhosOriginais[imagem] * midScale;

                    //Paragraph p_next = imagem.Range.Paragraphs[1].Next();
                    //if (p_next != null)
                    //{
                    //    Style estilo = p_next.get_Style() as Style;
                    //    if (estilo != null && estilo.NameLocal == "12 - Legendas de Figuras (PeriTAB)")
                    //    {
                    //        Globals.ThisAddIn.Dicionario_Window_e_UserControl[Globals.ThisAddIn.Application.ActiveWindow].Alinha_Legenda_de_Figura(p_next);
                    //    }
                    //}

                    //Paragraph p_next_next = p_next.Next();
                    //if (p_next_next != null)
                    //{
                    //    Style estilo = p_next_next.get_Style() as Style;
                    //    if (estilo != null && estilo.NameLocal == "13 - Texto de Figuras (PeriTAB)")
                    //    {
                    //        Globals.ThisAddIn.Dicionario_Window_e_UserControl[Globals.ThisAddIn.Application.ActiveWindow].Alinha_Texto_de_Figura(p_next_next);
                    //    }
                    //}

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

        private async void Button_borda_preta_Click(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon_com_UI_responsiva(sender, e, async progress =>
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
                await Task.CompletedTask;
            });
        }

        private async void Button_borda_vermelha_Click(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon_com_UI_responsiva(sender, e, async progress =>
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
                await Task.CompletedTask;
            });
        }

        private async void Button_borda_amarela_Click(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon_com_UI_responsiva(sender, e, async progress =>
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
                await Task.CompletedTask;
            });
        }

        private async void Button_legenda_imagem_Click(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon_com_UI_responsiva(sender, e, async progress =>
            {
                string estilo_nome_baseado = "Legenda";
                Globals.ThisAddIn.Application.OrganizerCopy(PeriTAB.Ribbon.Variables.Caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);

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
                        Globals.ThisAddIn.Dicionario_Window_e_UserControl.Values.First().Alinha_Legenda_de_Figura(Globals.ThisAddIn.Application.Selection.Paragraphs[1]);
                    }
                }
                await Task.CompletedTask;
            });
        }

        private async void Button_remove_borda_Click(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon_com_UI_responsiva(sender, e, async progress =>
            {
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.Line.Visible = MsoTriState.msoFalse;
                    }
                }
                await Task.CompletedTask;
            });
        }

        private async void Button_remove_formatacao_Click_1(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon_com_UI_responsiva(sender, e, async progress =>
            {
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.Reset();
                    }
                }
                await Task.CompletedTask;
            });
        }

        private async void Button_remove_forma_Click(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon_com_UI_responsiva(sender, e, async progress =>
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
                await Task.CompletedTask;
            });
        }

        private async void Button_remove_texto_alt_Click(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon_com_UI_responsiva(sender, e, async progress =>
            {
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.AlternativeText = "";
                    }
                }
                await Task.CompletedTask;
            });
        }

        private async void Button_remove_imagem_Click(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon_com_UI_responsiva(sender, e, async progress =>
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
                await Task.CompletedTask;
            });
        }
    }
}