using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Windows.Shapes;

namespace PeriTAB
{
    public partial class MyUserControl : UserControl
    {
        // Dicionário de instância para mapear o nome do estilo ao botão
        public readonly Dictionary<string, Button> dict_estilo_e_botao = new Dictionary<string, Button>();
        public readonly Dictionary<Button, string> dict_botao_e_estilo = new Dictionary<Button, string>();
        public MyUserControl()
        {
            InitializeComponent();

            // Define os estilos e botões associados
            var estilos_e_botoes = new (string, Button)[]
            {
                 ("01 - Sem Formatação (PeriTAB)", button_sem_formatacao)
                ,("02 - Corpo do Texto (PeriTAB)", button_corpo_do_texto)
                ,("03 - Parágrafo Numerado (PeriTAB)", button_paragrafo_numerado)
                ,("04 - Citações (PeriTAB)", button_citacoes)
                ,("05 - Seção_1 (PeriTAB)", button_secao_1)
                ,("06 - Seção_2 (PeriTAB)", button_secao_2)
                ,("07 - Seção_3 (PeriTAB)", button_secao_3)
                ,("08 - Seção_4 (PeriTAB)", button_secao_4)
                ,("09 - Seção_5 (PeriTAB)", button_secao_5)
                ,("10 - Enumerações (PeriTAB)", button_enumeracao)
                ,("11 - Figuras (PeriTAB)", button_figuras)
                ,("12 - Legendas de Figuras (PeriTAB)", button_legendas_de_figuras)
                ,("13 - Texto de Figuras (PeriTAB)", button_textos_de_figuras)
                ,("14 - Legendas de Tabelas (PeriTAB)", button_legendas_de_tabelas)
                ,("15 - Quesitos (PeriTAB)", button_quesitos)
                ,("16 - Fecho (PeriTAB)", button_fecho)
                ,("17 - Notas de rodapé (PeriTAB)", button_notas_de_rodape)
            };

            // Dicionário associando os botões aos seus estilos (Preenche o dicionário usando um loop)
            foreach (var item in estilos_e_botoes)
            {
                dict_estilo_e_botao.Add(item.Item1, item.Item2);
            }

            // Implementa o dicionario invertido
            dict_botao_e_estilo = dict_estilo_e_botao.ToDictionary(par => par.Value, par => par.Key);
        }

        private void MyUserControl_Load(object sender, EventArgs e)
        {
        }

        private /*async*/ void MyUserControl_Button_Click(object sender, EventArgs e)
        {
            Button Button = (Button)sender;
            Button.Invoke((Action)(() => Button.Enabled = false));

            string msg_StatusBar = Button.Name + ": ";
            bool success = true;

            Globals.ThisAddIn.Application.ScreenUpdating = false;

            Importa_todos_estilos();
            string estilo_nome = dict_botao_e_estilo[sender as Button];

            //await Tarefa.Run(() => DEU ERRO NA HORA DE PINTAR A TASKPANE - TEM QUE OBRIGAR A VERIFICAR A PINTURA DA TASKPANE DEPOIS DE RODAR ESSA TASK. COMO FAZ PARA ORDENAR ESSAS AÇÕES?
            //{
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
            List<Paragraph> list_Paragraph = new List<Paragraph>();
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
            {
                list_Paragraph.Add(p);
            }

            // Deletes
            foreach (Paragraph p in list_Paragraph)
            {
                if (new List<string>
                {
                    "04 - Citações (PeriTAB)",
                    "05 - Seção_1 (PeriTAB)",
                    "06 - Seção_2 (PeriTAB)",
                    "07 - Seção_3 (PeriTAB)",
                    "08 - Seção_4 (PeriTAB)",
                    "09 - Seção_5 (PeriTAB)",
                    "11 - Figuras (PeriTAB)",
                    "12 - Legendas de Figuras (PeriTAB)",
                    "13 - Texto de Figuras (PeriTAB)",
                    "14 - Legendas de Tabelas (PeriTAB)",
                    "15 - Quesitos (PeriTAB)",
                    "16 - Fecho (PeriTAB)"
                }.Contains(estilo_nome))
                {
                    Deleta_Paragrafos_Em_Branco(p, p.Previous());
                }

                if (new List<string>
                {
                    "04 - Citações (PeriTAB)",
                }.Contains(estilo_nome))
                {
                    Deleta_Paragrafos_Em_Branco(p, p.Next());
                }

                if (new List<string>
                {
                    "05 - Seção_1 (PeriTAB)",
                    "06 - Seção_2 (PeriTAB)",
                    "07 - Seção_3 (PeriTAB)",
                    "08 - Seção_4 (PeriTAB)",
                    "09 - Seção_5 (PeriTAB)"
                }.Contains(estilo_nome))
                {
                    Deleta_prefixo(p, new Regex(@"^\s*(M{0,3}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})(\.\d+)*\s*[-\u2013]\s*)")); // Expressão regular para identificar prefixos de números romanos + ponto + número arabico + espaços + hífen ou en-dash + espaços
                }
            }

            // Aplica Estilo
            foreach (Paragraph p in list_Paragraph)
            {
                try { p.set_Style((object)estilo_nome); }
                catch (System.Runtime.InteropServices.COMException) { success = false; } //Para impedir erro em parágrafos com modificação não permitida, como ContentControls
            }

            // Ajuste de formatação
            foreach (Paragraph p in list_Paragraph)
            {
                if (new List<string>
                {
                    "05 - Seção_1 (PeriTAB)",
                    "06 - Seção_2 (PeriTAB)",
                    "07 - Seção_3 (PeriTAB)",
                    "08 - Seção_4 (PeriTAB)",
                    "09 - Seção_5 (PeriTAB)",
                    "15 - Quesitos (PeriTAB)"
                }.Contains(estilo_nome))
                {
                    Zera_SpaceBefore_Se_paragrafo_anterior(p, new List<string> { "05 - Seção_1 (PeriTAB)", "06 - Seção_2 (PeriTAB)", "07 - Seção_3 (PeriTAB)", "08 - Seção_4 (PeriTAB)", "09 - Seção_5 (PeriTAB)" });
                    Zera_SpaceBefore_Se_paragrafo_anterior(p.Next(), new List<string> { "05 - Seção_1 (PeriTAB)", "06 - Seção_2 (PeriTAB)", "07 - Seção_3 (PeriTAB)", "08 - Seção_4 (PeriTAB)", "09 - Seção_5 (PeriTAB)" });
                }
                if (new List<string>
                {
                    "12 - Legendas de Figuras (PeriTAB)"
                }.Contains(estilo_nome))
                {
                    Alinha_Legenda_de_Figura(p);
                    if (p.Next() != null) Alinha_Texto_de_Figura(p.Next());
                }
                if (new List<string>
                {
                    "13 - Texto de Figuras (PeriTAB)"
                }.Contains(estilo_nome))
                {
                    Alinha_Texto_de_Figura(p);
                }
                if (new List<string>
                {
                    "14 - Legendas de Tabelas (PeriTAB)"
                }.Contains(estilo_nome))
                {
                    Alinha_Legenda_de_Tabela(p);
                }
                if (new List<string>
                {
                    "15 - Quesitos (PeriTAB)"
                }.Contains(estilo_nome))
                {
                    Ajusta_Quesito(p, new Regex(@"^\s*([a-zA-Z0-9]+\s*[-\u2013.)])\s*")); // Expressão regular para identificar numeração de quesitos
                }
            }
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            //});

            Globals.ThisAddIn.Application.ScreenUpdating = true;

            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;

            Button.Invoke((Action)(() => Button.Enabled = true));
        }

        private void Alinha_Legenda_de_Figura(Paragraph p)
        {
            PageSetup sectionPageSetup = p.Range.Sections[1].PageSetup;
            Paragraph previousParagraph = p.Previous();
            if (previousParagraph == null) return;

            float LeftMargin = sectionPageSetup.LeftMargin;
            float RightMargin = sectionPageSetup.RightMargin;
            float PageWidth = sectionPageSetup.PageWidth;

            float minLeftPosition = PageWidth;
            float maxRightPosition = 0;

            // FIGURA(S) NO PARÁGRAFO ANTERIOR
            if (previousParagraph.Range.InlineShapes.Count > 0)
            {
                foreach (InlineShape Imagem in previousParagraph.Range.InlineShapes)
                {
                    // ERRO NO CALCULO DA INFORMATION. NAO SEI POR QUE, MAS QUANDO CALCULA PELA SEGUNDA VEZ, FUNCIONA. - AINDA ASSIM HÁ UM ERRO INTERMITENTE QUE NÃO CONSEGUI SOLUCIONAR (adicionei um terceiro calculo de information para tentar solucionar o problema intermitente)
                    float LeftPosition = Imagem.Range.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary);
                    LeftPosition = Imagem.Range.get_Information(WdInformation.wdHorizontalPositionRelativeToPage);
                    LeftPosition = Imagem.Range.get_Information(WdInformation.wdHorizontalPositionRelativeToPage);
                    float RightPosition = LeftPosition + Imagem.Width;
                    //RightPosition = LeftPosition + Imagem.Width;

                    minLeftPosition = Math.Min(minLeftPosition, LeftPosition);
                    maxRightPosition = Math.Max(maxRightPosition, RightPosition);
                }
                p.Range.ParagraphFormat.LeftIndent = minLeftPosition - LeftMargin;
                p.Range.ParagraphFormat.RightIndent = PageWidth - maxRightPosition - RightMargin;

                // Ajuste para tabela
                if (previousParagraph.Range.Information[WdInformation.wdWithInTable] && p.Range.Information[WdInformation.wdWithInTable])
                {
                    // Teste para mesma tabela
                    Table Table1 = previousParagraph.Range.Tables[1];
                    Table Table2 = p.Range.Tables[1];
                    Table1.ID = "tabela 1";
                    if (Table2.ID == "tabela 1")
                    {
                        Table tabela = Table1;

                        // Teste para mesma celula
                        Cell Cell1 = previousParagraph.Range.Cells[1];
                        Cell Cell2 = p.Range.Cells[1];
                        Cell1.ID = "celula 1";
                        if (Cell2.ID == "celula 1")
                        {
                            Cell celula = Cell1;

                            float larguraEsquerda = 0;
                            for (int n = 1; n < celula.ColumnIndex; n++)
                            {
                                larguraEsquerda += tabela.Cell(celula.RowIndex, n).Width;
                            }
                            p.Range.ParagraphFormat.LeftIndent -= larguraEsquerda + tabela.LeftPadding;

                            float larguraDireita = 0;
                            for (int n = celula.ColumnIndex + 1; n <= tabela.Rows[1].Cells.Count; n++)
                            {
                                larguraDireita += tabela.Cell(celula.RowIndex, n).Width;
                            }

                            p.Range.ParagraphFormat.RightIndent -= larguraDireita + tabela.RightPadding;
                        }
                    }
                }
            }
            // UMA TELA DE DESENHO NO PARÁGRAFO ANTERIOR
            if (previousParagraph.Range.ShapeRange.Count == 1)
            {
                if (previousParagraph.Range.ShapeRange.Type == MsoShapeType.msoCanvas)
                {
                    // Obtém o recuo da tela de desenho
                    float Tela_LeftPosition = previousParagraph.Range.ShapeRange.Anchor.Information[WdInformation.wdHorizontalPositionRelativeToTextBoundary];
                    Tela_LeftPosition = previousParagraph.Range.ShapeRange.Anchor.Information[WdInformation.wdHorizontalPositionRelativeToPage];
                    Tela_LeftPosition = previousParagraph.Range.ShapeRange.Anchor.Information[WdInformation.wdHorizontalPositionRelativeToPage];
                    //Range Selecao_inicial = Globals.ThisAddIn.Application.Selection.Range; // Salva a seleção inicial
                    //previousParagraph.Range.ShapeRange.Select();
                    //float Tela_LeftPosition = Globals.ThisAddIn.Application.Selection.get_Information(WdInformation.wdHorizontalPositionRelativeToPage);
                    //Selecao_inicial.Select(); // Restaura a seleção inicial

                    if (previousParagraph.Range.ShapeRange.Line.Style == MsoLineStyle.msoLineStyleMixed)
                    {
                        // Tela de desenho sem borda
                        foreach (Microsoft.Office.Interop.Word.Shape Imagem in previousParagraph.Range.ShapeRange[1].CanvasItems)
                        {
                            if (Imagem.Type == MsoShapeType.msoPicture)
                            {
                                float LeftPosition = Imagem.Left * 20; //ERRO NO CALCULO DA .LEFT. CORRIGIDO APÓS MULTIPLICAR POR 20. NÃO SE SABE POR QUE.
                                float RightPosition = LeftPosition + Imagem.Width;

                                minLeftPosition = Math.Min(minLeftPosition, LeftPosition);
                                maxRightPosition = Math.Max(maxRightPosition, RightPosition);
                            }
                        }

                        p.Range.ParagraphFormat.LeftIndent = Tela_LeftPosition + minLeftPosition - LeftMargin;
                        p.Range.ParagraphFormat.RightIndent = PageWidth - Tela_LeftPosition - maxRightPosition - RightMargin ;
                    }
                    else
                    {
                        float Tela_Width = previousParagraph.Range.ShapeRange[1].Width;
                        // Tela de desenho com borda
                        p.Range.ParagraphFormat.LeftIndent = Tela_LeftPosition - LeftMargin;
                        p.Range.ParagraphFormat.RightIndent = PageWidth - Tela_LeftPosition - Tela_Width - RightMargin;
                    }
                }
            }
            // UMA TABELA NO PARÁGRAFO ANTERIOR
            if (!p.Range.Information[WdInformation.wdWithInTable] && previousParagraph.Range.Information[WdInformation.wdWithInTable])
            {
                Table tabela = previousParagraph.Range.Tables[1];

                // tabela sem borda
                if (tabela.Borders[WdBorderType.wdBorderTop].LineStyle == WdLineStyle.wdLineStyleNone &&
                    tabela.Borders[WdBorderType.wdBorderBottom].LineStyle == WdLineStyle.wdLineStyleNone)
                {
                    foreach (Cell Cell in tabela.Range.Cells)
                    {
                        foreach (InlineShape Imagem in Cell.Range.InlineShapes)
                        {
                            float LeftPosition = Imagem.Range.get_Information(WdInformation.wdHorizontalPositionRelativeToPage) + Imagem.Range.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary);
                            float RightPosition = LeftPosition + Imagem.Width;

                            minLeftPosition = Math.Min(minLeftPosition, LeftPosition);
                            maxRightPosition = Math.Max(maxRightPosition, RightPosition);
                        }
                    }
                    p.Range.ParagraphFormat.LeftIndent = minLeftPosition - LeftMargin;
                    p.Range.ParagraphFormat.RightIndent = PageWidth - maxRightPosition - RightMargin;
                }
                // tabela com borda
                else
                {
                        float TextLeftPosition = tabela.Range.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary);
                        float PageLeftPosition = tabela.Range.get_Information(WdInformation.wdHorizontalPositionRelativeToPage);
                        PageLeftPosition = tabela.Range.get_Information(WdInformation.wdHorizontalPositionRelativeToPage);

                        float maxTabela_Width = 0;
                        for (int i = 1; i <= Math.Min(10, tabela.Rows.Count); i++)
                        {
                            Row linha = tabela.Rows[i];

                            float larguraLinha = 0;
                            foreach (Cell celula in linha.Cells)
                            {
                                larguraLinha += celula.Width;
                            }

                            maxTabela_Width = Math.Max(maxTabela_Width, larguraLinha);
                        }

                        p.Range.ParagraphFormat.LeftIndent = PageLeftPosition - LeftMargin - tabela.LeftPadding - TextLeftPosition;
                        p.Range.ParagraphFormat.RightIndent = PageWidth - PageLeftPosition - maxTabela_Width - RightMargin + tabela.LeftPadding + TextLeftPosition;
                }
            }
        }

        private void Alinha_Texto_de_Figura(Paragraph p)
        {
            if ((((Microsoft.Office.Interop.Word.Style)p.get_Style()).NameLocal.ToString()) == "13 - Texto de Figuras (PeriTAB)")
            {
                if (p.Previous() != null)
                {
                    if ((((Microsoft.Office.Interop.Word.Style)p.Previous().get_Style()).NameLocal.ToString()) == "12 - Legendas de Figuras (PeriTAB)")
                    {
                        p.Range.ParagraphFormat.LeftIndent = p.Previous().Range.ParagraphFormat.LeftIndent;
                        p.Range.ParagraphFormat.RightIndent = p.Previous().Range.ParagraphFormat.RightIndent;
                        p.Previous().Range.ParagraphFormat.SpaceAfter = 0;
                        p.Previous().Range.ParagraphFormat.KeepWithNext = -1;
                    }
                }
            }
        }
        private void Alinha_Legenda_de_Tabela(Paragraph p)
        {
            PageSetup sectionPageSetup = p.Range.Sections[1].PageSetup;
            Paragraph nextParagraph = p.Next();
            if (nextParagraph == null) return;

            float LeftMargin = sectionPageSetup.LeftMargin;
            float RightMargin = sectionPageSetup.RightMargin;
            float PageWidth = sectionPageSetup.PageWidth;

            // UMA FIGURA NO PARÁGRAFO POSTERIOR
            if (nextParagraph.Range.InlineShapes.Count == 1)
            {
                InlineShape Imagem = nextParagraph.Range.InlineShapes[1];

                // ERRO NO CALCULO DA INFORMATION. NAO SEI POR QUE, MAS QUANDO CALCULA PELA SEGUNDA VEZ, FUNCIONA. - AINDA ASSIM HÁ UM ERRO INTERMITENTE QUE NÃO CONSEGUI SOLUCIONAR (adicionei um terceiro calculo de information para tentar solucionar o problema intermitente)
                float LeftPosition = Imagem.Range.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary);
                LeftPosition = Imagem.Range.get_Information(WdInformation.wdHorizontalPositionRelativeToPage);
                LeftPosition = Imagem.Range.get_Information(WdInformation.wdHorizontalPositionRelativeToPage);
                float RightPosition = LeftPosition + Imagem.Width;

                p.Range.ParagraphFormat.LeftIndent = LeftPosition - LeftMargin;
                p.Range.ParagraphFormat.RightIndent = PageWidth - RightPosition - RightMargin;
            }
            // UMA TABELA NO PARÁGRAFO POSTERIOR
            if (nextParagraph.Range.Tables.Count == 1)
            {
                Table tabela = nextParagraph.Range.Tables[1];
                
                float TextLeftPosition = tabela.Range.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary);
                float PageLeftPosition = tabela.Range.get_Information(WdInformation.wdHorizontalPositionRelativeToPage);
                PageLeftPosition = tabela.Range.get_Information(WdInformation.wdHorizontalPositionRelativeToPage);

                float maxTabela_Width = 0;
                for (int i = 1; i <= Math.Min(10, tabela.Rows.Count); i++)
                {
                    Row linha = tabela.Rows[i];

                    float larguraLinha = 0;
                    foreach (Cell celula in linha.Cells)
                    {
                        larguraLinha += celula.Width;
                    }

                    maxTabela_Width = Math.Max(maxTabela_Width, larguraLinha);
                }

                p.Range.ParagraphFormat.LeftIndent = PageLeftPosition - LeftMargin - tabela.LeftPadding - TextLeftPosition;
                p.Range.ParagraphFormat.RightIndent = PageWidth - PageLeftPosition - maxTabela_Width - RightMargin + tabela.LeftPadding + TextLeftPosition;
            }
        }

        private void Zera_SpaceBefore_Se_paragrafo_anterior(Paragraph p, List<string> list)
        {
            Paragraph PreviousParagraph = p.Previous();
            if (PreviousParagraph != null)
            {
                if (list.Contains(p.Previous().get_Style().NameLocal))
                {
                    p.Range.ParagraphFormat.SpaceBefore = 0;
                }
            }
        }

        public void Importa_todos_estilos()
        {
            List<string> listaEstilos = dict_estilo_e_botao.Keys.ToList();
            foreach (string estilo in listaEstilos)
            {
                Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo, WdOrganizerObject.wdOrganizerObjectStyles);
            }
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "Normal", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "Legenda", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "Texto de nota de rodapé", WdOrganizerObject.wdOrganizerObjectStyles);
        }

        private void Deleta_Paragrafos_Em_Branco(Paragraph paragrafoInicial, Paragraph paragrafoDirecao) // pode ser p.Previous() ou p.Next()
        {
            Paragraph paragrafoAtual = paragrafoDirecao;

            while (paragrafoAtual != null && string.IsNullOrWhiteSpace(paragrafoAtual.Range.Text))
            {
                paragrafoAtual.Range.Delete();
                paragrafoAtual = (paragrafoAtual == paragrafoInicial.Next()) ? paragrafoAtual.Next() : paragrafoAtual.Previous();
            }
        }
        private void Deleta_prefixo(Paragraph p, Regex prefixo)  //Deleta parágrafos anteriores em branco ou apenas com espaços
        {
            Match match = prefixo.Match(p.Range.Text);
            if (match.Success)
            {
                Range prefixRange = p.Range.Duplicate; // Cria um range duplicado para não afetar o original
                prefixRange.Start = p.Range.Start;     // Início do range = início do parágrafo
                prefixRange.End = p.Range.Start + match.Length; // Final do range = tamanho do prefixo identificado

                prefixRange.Delete();
            }
        }
        private void Ajusta_Quesito(Paragraph p, Regex prefixo)  //Deleta parágrafos anteriores em branco ou apenas com espaços
        {
            Match match = prefixo.Match(p.Range.Text);
            if (match.Success)
            {
                p.Range.Font.Bold = 0; // Remove o negrito de todo o texto
                int prefixLength = match.Length; // Calcula o comprimento do prefixo
                for (int i = 1; i <= prefixLength; i++)
                {
                    p.Range.Characters[i].Bold = 1;
                }
            }
        }

        //private void button_textos_de_figuras_Click(object sender, EventArgs e)
        //{
        //    Importa_todos_estilos();
        //    string estilo_nome = dict_botao_e_estilo[sender as Button];

        //    Globals.ThisAddIn.Application.ScreenUpdating = false;

        //    List<Paragraph> list_Paragraph = new List<Paragraph>();
        //    foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
        //    {
        //        list_Paragraph.Add(p);
        //    }

        //    // Aplica Estilo
        //    foreach (Paragraph p in list_Paragraph)
        //    {
        //        p.set_Style((object)estilo_nome);
        //    }

        //    // Ajuste de formatação
        //    foreach (Paragraph p in list_Paragraph)
        //    {
        //        if (p.Previous() != null)
        //        {
        //            if ((((Microsoft.Office.Interop.Word.Style)p.Previous().get_Style()).NameLocal.ToString()) == "12 - Legendas de Figuras (PeriTAB)")
        //            {
        //                p.Range.ParagraphFormat.LeftIndent = p.Previous().Range.ParagraphFormat.LeftIndent;
        //                p.Range.ParagraphFormat.RightIndent = p.Previous().Range.ParagraphFormat.RightIndent;
        //                p.Previous().Range.ParagraphFormat.SpaceAfter = 0;
        //                p.Previous().Range.ParagraphFormat.KeepWithNext = -1;
        //            }
        //        }
        //    }

        //    Globals.ThisAddIn.Application.ScreenUpdating = true;

        //}

        public void Habilita_Destaca(Button b, bool habilita, bool destaca = false)
        {
            b.Enabled = habilita;
            if (destaca) { b.BackColor = SystemColors.Highlight; b.ForeColor = SystemColors.HighlightText; }
        }

        public void Remove_Destaque_Botoes(MyUserControl UserControl)
        {
            foreach (var botao in UserControl.Controls.OfType<Button>())
            {
                botao.BackColor = SystemColors.Control;
                botao.ForeColor = SystemColors.ControlText;
            }
        }
    }
}
