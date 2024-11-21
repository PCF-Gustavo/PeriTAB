using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

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

        private void MyUserControl_Button_Click(object sender, EventArgs e)
        {
            Importa_todos_estilos();
            string estilo_nome = dict_botao_e_estilo[sender as Button];

            Globals.ThisAddIn.Application.ScreenUpdating = false;

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
                p.set_Style((object)estilo_nome);
            }

            Range Selecao_inicial = Globals.ThisAddIn.Application.Selection.Range; //Salva a seleção inicial
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
                }

                if (new List<string>
                {
                    "15 - Quesitos (PeriTAB)"
                }.Contains(estilo_nome))
                {
                    Ajusta_Quesito(p, new Regex(@"\s*([a-zA-Z0-9]+\s*[-\u2013.)])\s*")); // Expressão regular para identificar numeração de quesitos
                }

                if (new List<string>
                {
                    "12 - Legendas de Figuras (PeriTAB)",
                    "14 - Legendas de Tabelas (PeriTAB)"
                }.Contains(estilo_nome))
                {
                    p.Range.Select();
                    Globals.ThisAddIn.Application.Run("alinha_legenda");
                }

            }
            Selecao_inicial.Select(); // Restaura a seleção inicial

            Globals.ThisAddIn.Application.ScreenUpdating = true;
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
                int prefixLength = match.Length; // Calcula o comprimento do prefixo
                for (int i = 1; i <= prefixLength; i++)
                {
                    if (p.Range.Characters[1].Fields.Count > 0)
                    {
                        p.Range.Characters[1].Fields.Unlink();
                    }
                    p.Range.Characters[1].Delete();
                }
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


        private void button_textos_de_figuras_Click(object sender, EventArgs e)
        {
            Importa_todos_estilos();
            string estilo_nome = dict_botao_e_estilo[sender as Button];

            Globals.ThisAddIn.Application.ScreenUpdating = false;

            List<Paragraph> list_Paragraph = new List<Paragraph>();
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
            {
                list_Paragraph.Add(p);
            }

            // Aplica Estilo
            foreach (Paragraph p in list_Paragraph)
            {
                p.set_Style((object)estilo_nome);
            }

            // Ajuste de formatação
            foreach (Paragraph p in list_Paragraph)
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

            Globals.ThisAddIn.Application.ScreenUpdating = true;

        }

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
