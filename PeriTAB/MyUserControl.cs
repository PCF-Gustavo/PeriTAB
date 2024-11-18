using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using System.Text.RegularExpressions;

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

            // Inicializa o dicionário associando os botões aos seus estilos
            //dict_estilo_e_botao2.Add("01 - Sem Formatação (PeriTAB)", button_sem_formatacao);
            //dict_estilo_e_botao2.Add("02 - Corpo do Texto (PeriTAB)", button_corpo_do_texto);
            //dict_estilo_e_botao2.Add("03 - Citações (PeriTAB)", button_citacoes);
            //dict_estilo_e_botao2.Add("04a - Seção_1 (PeriTAB)", button_secao_1);
            //dict_estilo_e_botao2.Add("04b - Seção_2 (PeriTAB)", button_secao_2);
            //dict_estilo_e_botao2.Add("04c - Seção_3 (PeriTAB)", button_secao_3);
            //dict_estilo_e_botao2.Add("04d - Seção_4 (PeriTAB)", button_secao_4);
            //dict_estilo_e_botao2.Add("05 - Enumerações (PeriTAB)", button_enumeracao);
            //dict_estilo_e_botao2.Add("06 - Figuras (PeriTAB)", button_figuras);
            //dict_estilo_e_botao2.Add("07 - Legendas de Figuras (PeriTAB)", button_legendas_de_figuras);
            //dict_estilo_e_botao2.Add("08a - Texto de Figuras (PeriTAB)", button_textos_de_figuras);
            //dict_estilo_e_botao2.Add("08 - Legendas de Tabelas (PeriTAB)", button_legendas_de_tabelas);
            //dict_estilo_e_botao2.Add("09 - Quesitos (PeriTAB)", button_quesitos);
            //dict_estilo_e_botao2.Add("10 - Fecho (PeriTAB)", button_fecho);
            //dict_estilo_e_botao2.Add("11 - Parágrafo Numerado (PeriTAB)", button_paragrafo_numerado);

            // Definindo os estilos e botões associados
            var estilos_e_botoes = new (string, Button)[]
            {
                // ("01 - Sem Formatação (PeriTAB)", button_sem_formatacao)
                //,("02 - Corpo do Texto (PeriTAB)", button_corpo_do_texto)
                //,("03 - Citações (PeriTAB)", button_citacoes)
                //,("04a - Seção_1 (PeriTAB)", button_secao_1)
                //,("04b - Seção_2 (PeriTAB)", button_secao_2)
                //,("04c - Seção_3 (PeriTAB)", button_secao_3)
                //,("04d - Seção_4 (PeriTAB)", button_secao_4)
                //,("10 - Seção_5 (PeriTAB)", button_secao_5)
                //,("05 - Enumerações (PeriTAB)", button_enumeracao)
                //,("06 - Figuras (PeriTAB)", button_figuras)
                //,("07 - Legendas de Figuras (PeriTAB)", button_legendas_de_figuras)
                //,("08a - Texto de Figuras (PeriTAB)", button_textos_de_figuras)
                //,("08 - Legendas de Tabelas (PeriTAB)", button_legendas_de_tabelas)
                //,("09 - Quesitos (PeriTAB)", button_quesitos)
                //,("10 - Fecho (PeriTAB)", button_fecho)
                //,("11 - Parágrafo Numerado (PeriTAB)", button_paragrafo_numerado)
                //,("17 - Notas de rodapé (PeriTAB)", button_notas_de_rodape)
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

        private void button_sem_formatacao_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon.Variables.debugging) { stopwatch.Start(); }
            //Button botao = sender as Button;
            string estilo_nome = dict_botao_e_estilo[sender as Button];
            //MessageBox.Show(botao.Name);
            //MessageBox.Show(estilo_nome);
            string estilo_nome_baseado = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_BaseStyle().NameLocal;
            //MessageBox.Show(estilo_nome_baseado);
            string estilo_nome_seguinte = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_NextParagraphStyle().NameLocal;
            //MessageBox.Show(estilo_nome_seguinte);
            //string estilo_nome_baseado = "Normal";
            //string estilo_nome = "01 - Sem Formatação (PeriTAB)";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); }
            if (Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Estilo Sem Formatação: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        //private void button_sem_formatacao_Click(object sender, EventArgs e)
        //{
        //    Stopwatch stopwatch = new Stopwatch(); if (Ribbon.Variables.debugging) { stopwatch.Start(); }
        //    string estilo_nome_baseado = "Normal";
        //    string estilo_nome = "01 - Sem Formatação (PeriTAB)";
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.ScreenUpdating = false;
        //    foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs){ p.Range.set_Style((object)estilo_nome); }
        //    if (Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
        //    {
        //        string msg_StatusBar = "Estilo Sem Formatação: Sucesso";
        //        stopwatch.Stop();
        //        msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
        //        Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
        //    }
        //    Globals.ThisAddIn.Application.ScreenUpdating = true;
        //}

        private void button_corpo_do_texto_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon.Variables.debugging) { stopwatch.Start(); }
            //string estilo_nome_baseado = "Normal";
            //string estilo_nome = "02 - Corpo do Texto (PeriTAB)";
            string estilo_nome = dict_botao_e_estilo[sender as Button];
            string estilo_nome_baseado = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_BaseStyle().NameLocal;
            string estilo_nome_seguinte = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_NextParagraphStyle().NameLocal;
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); }
            if (Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Estilo Corpo do Texto: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        private void button_paragrafo_numerado_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon.Variables.debugging) { stopwatch.Start(); }
            //string estilo_nome_baseado = "Normal";
            //string estilo_nome = "11 - Parágrafo Numerado (PeriTAB)";
            string estilo_nome = dict_botao_e_estilo[sender as Button];
            string estilo_nome_baseado = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_BaseStyle().NameLocal;
            string estilo_nome_seguinte = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_NextParagraphStyle().NameLocal;
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); }
            if (Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Estilo Parágrafo Numerado: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        private void button_citacoes_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon.Variables.debugging) { stopwatch.Start(); }
            //string estilo_nome_baseado = "Normal";
            //string estilo_nome = "03 - Citações (PeriTAB)";
            string estilo_nome = dict_botao_e_estilo[sender as Button];
            string estilo_nome_baseado = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_BaseStyle().NameLocal;
            string estilo_nome_seguinte = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_NextParagraphStyle().NameLocal;
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); }
            Range r1 = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r1 != null) { if (r1.Text == ((char)13).ToString()) { r1.Delete(); } } //Deleta parágrafo anterior em branco
            Range r2 = Globals.ThisAddIn.Application.Selection.Next(WdUnits.wdParagraph, 1); if (r2 != null) { if (r2.Text == ((char)13).ToString()) { r2.Delete(); } } //Deleta parágrafo seguinte em branco
            if (Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Estilo Citações: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        //private void button_secao_1_Click(object sender, EventArgs e)
        //{
        //string estilo_nome = "04 - Seções (PeriTAB)";
        //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
        //Globals.ThisAddIn.Application.ScreenUpdating = false;
        //foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); p.Range.SetListLevel(1); }
        //Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco
        //Globals.ThisAddIn.Application.ScreenUpdating = true;
        //}
        //private void button_secao_2_Click(object sender, EventArgs e)
        //{
        //    string estilo_nome = "04 - Seções (PeriTAB)";
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.ScreenUpdating = false;
        //    foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); p.Range.SetListLevel(2); p.Range.Font.AllCaps = 0; }
        //    Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco
        //    Globals.ThisAddIn.Application.ScreenUpdating = true;
        //}

        //private void button_secao_3_Click(object sender, EventArgs e)
        //{
        //    string estilo_nome = "04 - Seções (PeriTAB)";
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.ScreenUpdating = false;
        //    foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); p.Range.SetListLevel(3); p.Range.Font.AllCaps = 0; p.Range.Font.Bold = 0; p.Range.Font.Italic = -1; }
        //    Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco
        //    Globals.ThisAddIn.Application.ScreenUpdating = true;
        //}

        //private void button_secao_4_Click(object sender, EventArgs e)
        //{
        //    string estilo_nome = "04 - Seções (PeriTAB)";
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.ScreenUpdating = false;
        //    foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); p.Range.SetListLevel(4); p.Range.Font.AllCaps = 0; p.Range.Font.Bold = 0; p.Range.Font.Underline = WdUnderline.wdUnderlineSingle; }
        //    Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco
        //    Globals.ThisAddIn.Application.ScreenUpdating = true;
        //}
        //private void button_secao_1_Click(object sender, EventArgs e)
        //{
        //    Stopwatch stopwatch = new Stopwatch(); if (Ribbon.Variables.debugging) { stopwatch.Start(); }
        //    //string estilo_nome_baseado = "04 - Seções (PeriTAB)";
        //    //string estilo_nome_seguinte = "02 - Corpo do Texto (PeriTAB)";
        //    //string estilo_nome = "04a - Seção_1 (PeriTAB)";
        //    string estilo_nome = dict_botao_e_estilo[sender as Button];
        //    string estilo_nome_baseado = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_BaseStyle().NameLocal;
        //    string estilo_nome_seguinte = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_NextParagraphStyle().NameLocal;
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.ScreenUpdating = false;

        //    foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) 
        //    {


        //        p.Range.set_Style((object)estilo_nome); 
        //        Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco

        //        int s;
        //        if (p.Range.Text.Length >= 7) { s = 7; } else { s = p.Range.Text.Length; }
        //        if (s == 0) break;
        //        if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count == 1) 
        //        {
        //            string a = p.Range.Text.Substring(0, s).Replace(((char)8211).ToString(), "-");
        //            try
        //            {
        //                if (a.Replace(" ", "").Substring(0, 2) == "I-" | a.Replace(" ", "").Substring(0, 3) == "II-" | a.Replace(" ", "").Substring(0, 4) == "III-" | a.Replace(" ", "").Substring(0, 3) == "IV-" | a.Replace(" ", "").Substring(0, 2) == "V-" | a.Replace(" ", "").Substring(0, 3) == "VI-" | a.Replace(" ", "").Substring(0, 4) == "VII-" | a.Replace(" ", "").Substring(0, 5) == "VIII-" | a.Replace(" ", "").Substring(0, 3) == "IX-" | a.Replace(" ", "").Substring(0, 2) == "X-" | a.Substring(0, 2) == "I." | a.Substring(0, 3) == "II." | a.Substring(0, 4) == "III." | a.Substring(0, 3) == "IV." | a.Substring(0, 2) == "V." | a.Substring(0, 3) == "VI." | a.Substring(0, 4) == "VII." | a.Substring(0, 5) == "VIII." | a.Substring(0, 3) == "IX." | a.Substring(0, 2) == "X.")
        //                {
        //                    //MessageBox.Show(a);
        //                    int loc_hifen = a.IndexOf("-");
        //                    //MessageBox.Show(loc_hifen.ToString());
        //                    for (int i = 1; i <= loc_hifen; i++)
        //                    {
        //                        //MessageBox.Show(p.Range.Characters[1].Text);
        //                        if (p.Range.Characters[1].Fields.Count > 0) { p.Range.Characters[1].Fields.Unlink(); }
        //                        p.Range.Characters[1].Delete();
        //                    }
        //                    //break;
        //                }
        //            }
        //            catch (System.ArgumentOutOfRangeException) { }
        //        }
        //    }
        //    if (Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
        //    {
        //        string msg_StatusBar = "Estilo Seção Primária: Sucesso";
        //        stopwatch.Stop();
        //        msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
        //        Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
        //    }
        //    Globals.ThisAddIn.Application.ScreenUpdating = true;
        //}

        private void button_secoes_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon.Variables.debugging) { stopwatch.Start(); }

            string estilo_nome = dict_botao_e_estilo[sender as Button];
            string estilo_nome_baseado = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_BaseStyle().NameLocal;
            string estilo_nome_seguinte = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_NextParagraphStyle().NameLocal;

            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);

            Globals.ThisAddIn.Application.ScreenUpdating = false;

            // Expressão regular para identificar prefixos de números romanos com espaços, hífen/en dash e mais espaços
            Regex Regex_prefixo_secoes = new Regex(@"^\s*(M{0,3}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})(\.\d+)*\s*[-\u2013]\s*)");

            List<Paragraph> list_Paragraph = new List<Paragraph>();
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) 
            {
                list_Paragraph.Add(p);
            }

            foreach (Paragraph p in list_Paragraph)
            {
                // Verifica e apaga todos os parágrafos anteriores em branco ou apenas com espaços
                //Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1);
                //while (r != null && string.IsNullOrWhiteSpace(r.Text)) // Verifica se o parágrafo está em branco ou apenas com espaços
                //{
                //    r.Delete(); // Deleta o parágrafo
                //    r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); // Avança para o parágrafo anterior
                //}
                Paragraph previousParagraph = p.Previous();
                while (previousParagraph != null && string.IsNullOrWhiteSpace(previousParagraph.Range.Text)) // Verifica se o parágrafo está em branco ou apenas com espaços
                {
                    previousParagraph.Range.Delete(); // Deleta o parágrafo
                    previousParagraph = p.Previous(); // Avança para o parágrafo anterior
                }

                // Verifica se o texto começa com o prefixo
                string paragraphText = p.Range.Text;
                Match match = Regex_prefixo_secoes.Match(paragraphText);

                if (match.Success)
                {
                    // Calcula o comprimento do prefixo e remove-o do parágrafo
                    int prefixLength = match.Length;
                    for (int i = 1; i <= prefixLength; i++)
                    {
                        if (p.Range.Characters[1].Fields.Count > 0)
                        {
                            p.Range.Characters[1].Fields.Unlink();
                        }
                        p.Range.Characters[1].Delete();
                    }
                }
                MessageBox.Show(" ");
                //p.Range.set_Style((object)estilo_nome);
            }

            foreach (Paragraph p in list_Paragraph)
            {
                p.set_Style((object)estilo_nome);
            }

            // Exibe o tempo de execução se estiver no modo debugging
            if (Ribbon.Variables.debugging)
            {
                string msg_StatusBar = "Estilo Seção: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }

            Globals.ThisAddIn.Application.ScreenUpdating = true;
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

        private void Importa_todos_estilos()
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


        //private void button_secao_2_Click(object sender, EventArgs e)
        //{
        //    Stopwatch stopwatch = new Stopwatch(); if (Ribbon.Variables.debugging) { stopwatch.Start(); }
        //    //string estilo_nome_baseado = "04 - Seções (PeriTAB)";
        //    //string estilo_nome_seguinte = "02 - Corpo do Texto (PeriTAB)";
        //    //string estilo_nome = "04b - Seção_2 (PeriTAB)";
        //    //Globals.ThisAddIn.Application.ScreenUpdating = false;
        //    //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
        //    //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
        //    //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
        //    string estilo_nome = dict_botao_e_estilo[sender as Button];
        //    string estilo_nome_baseado = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_BaseStyle().NameLocal;
        //    string estilo_nome_seguinte = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_NextParagraphStyle().NameLocal;
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.ScreenUpdating = false;

        //    foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
        //    {
        //        p.Range.set_Style((object)estilo_nome);
        //        Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); } } //Deleta parágrafo anterior em branco
        //        if (r != null & Globals.ThisAddIn.Application.Selection.Paragraphs.Count == 1)
        //        {
        //            Microsoft.Office.Interop.Word.Style r_estilo = (Microsoft.Office.Interop.Word.Style)r.get_Style();
        //            if (r_estilo != null) //Ao que parece, paragráfos com o estilo "revisado" perdem o parâmetro de estilo. Esta linha evita este erro.          
        //            {
        //                if (r_estilo.NameLocal.ToString() == "04a - Seção_1 (PeriTAB)" | r_estilo.NameLocal.ToString() == "04b - Seção_2 (PeriTAB)" | r_estilo.NameLocal.ToString() == "04c - Seção_3 (PeriTAB)" | r_estilo.NameLocal.ToString() == "04d - Seção_4 (PeriTAB)")
        //                {
        //                    p.Range.ParagraphFormat.SpaceBefore = 0;
        //                }
        //            }
        //        }

        //        int s;
        //        if (p.Range.Text.Length >= 9) { s = 9; } else { s = p.Range.Text.Length; }
        //        if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count == 1)
        //        {
        //            string a = p.Range.Text.Substring(0, s).Replace(((char)8211).ToString(), "-");
        //            try
        //            {
        //                if (a.Replace(" ", "").Substring(0, 2) == "I-" | a.Replace(" ", "").Substring(0, 3) == "II-" | a.Replace(" ", "").Substring(0, 4) == "III-" | a.Replace(" ", "").Substring(0, 3) == "IV-" | a.Replace(" ", "").Substring(0, 2) == "V-" | a.Replace(" ", "").Substring(0, 3) == "VI-" | a.Replace(" ", "").Substring(0, 4) == "VII-" | a.Replace(" ", "").Substring(0, 5) == "VIII-" | a.Replace(" ", "").Substring(0, 3) == "IX-" | a.Replace(" ", "").Substring(0, 2) == "X-" | a.Substring(0, 2) == "I." | a.Substring(0, 3) == "II." | a.Substring(0, 4) == "III." | a.Substring(0, 3) == "IV." | a.Substring(0, 2) == "V." | a.Substring(0, 3) == "VI." | a.Substring(0, 4) == "VII." | a.Substring(0, 5) == "VIII." | a.Substring(0, 3) == "IX." | a.Substring(0, 2) == "X.")
        //            {
        //                int loc_hifen = a.IndexOf("-");
        //                for (int i = 1; i <= loc_hifen; i++)
        //                {
        //                    if (p.Range.Characters[1].Fields.Count > 0) { p.Range.Characters[1].Fields.Unlink(); }
        //                    p.Range.Characters[1].Delete();
        //                }
        //            }
        //            }
        //            catch (System.ArgumentOutOfRangeException) { }
        //        }
        //    }
        //    if (Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
        //    {
        //        string msg_StatusBar = "Estilo Seção Secundária: Sucesso";
        //        stopwatch.Stop();
        //        msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
        //        Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
        //    }
        //    Globals.ThisAddIn.Application.ScreenUpdating = true;
        //}

        //private void button_secao_3_Click(object sender, EventArgs e)
        //{
        //    Stopwatch stopwatch = new Stopwatch(); if (Ribbon.Variables.debugging) { stopwatch.Start(); }
        //    //string estilo_nome_baseado = "04 - Seções (PeriTAB)";
        //    //string estilo_nome_seguinte = "02 - Corpo do Texto (PeriTAB)";
        //    //string estilo_nome = "04c - Seção_3 (PeriTAB)";
        //    //Globals.ThisAddIn.Application.ScreenUpdating = false;
        //    //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
        //    //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
        //    //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
        //    string estilo_nome = dict_botao_e_estilo[sender as Button];
        //    string estilo_nome_baseado = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_BaseStyle().NameLocal;
        //    string estilo_nome_seguinte = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_NextParagraphStyle().NameLocal;
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.ScreenUpdating = false;

        //    foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
        //    {
        //        p.Range.set_Style((object)estilo_nome);
        //        Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); } } //Deleta parágrafo anterior em branco
        //        if (r != null & Globals.ThisAddIn.Application.Selection.Paragraphs.Count == 1)
        //        {
        //            Microsoft.Office.Interop.Word.Style r_estilo = (Microsoft.Office.Interop.Word.Style)r.get_Style();
        //            if (r_estilo != null) //Ao que parece, paragráfos com o estilo "revisado" perdem o parâmetro de estilo. Esta linha evita este erro.          
        //            {
        //                if (r_estilo.NameLocal.ToString() == "04a - Seção_1 (PeriTAB)" | r_estilo.NameLocal.ToString() == "04b - Seção_2 (PeriTAB)" | r_estilo.NameLocal.ToString() == "04c - Seção_3 (PeriTAB)" | r_estilo.NameLocal.ToString() == "04d - Seção_4 (PeriTAB)")
        //                {
        //                    p.Range.ParagraphFormat.SpaceBefore = 0;
        //                }
        //            }
        //        }

        //        int s;
        //        if (p.Range.Text.Length >= 11) { s = 11; } else { s = p.Range.Text.Length; }
        //        if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count == 1)
        //        {
        //            string a = p.Range.Text.Substring(0, s).Replace(((char)8211).ToString(), "-");
        //            try
        //            {
        //                if (a.Replace(" ", "").Substring(0, 2) == "I-" | a.Replace(" ", "").Substring(0, 3) == "II-" | a.Replace(" ", "").Substring(0, 4) == "III-" | a.Replace(" ", "").Substring(0, 3) == "IV-" | a.Replace(" ", "").Substring(0, 2) == "V-" | a.Replace(" ", "").Substring(0, 3) == "VI-" | a.Replace(" ", "").Substring(0, 4) == "VII-" | a.Replace(" ", "").Substring(0, 5) == "VIII-" | a.Replace(" ", "").Substring(0, 3) == "IX-" | a.Replace(" ", "").Substring(0, 2) == "X-" | a.Substring(0, 2) == "I." | a.Substring(0, 3) == "II." | a.Substring(0, 4) == "III." | a.Substring(0, 3) == "IV." | a.Substring(0, 2) == "V." | a.Substring(0, 3) == "VI." | a.Substring(0, 4) == "VII." | a.Substring(0, 5) == "VIII." | a.Substring(0, 3) == "IX." | a.Substring(0, 2) == "X.")
        //                {
        //                    int loc_hifen = a.IndexOf("-");
        //                    for (int i = 1; i <= loc_hifen; i++)
        //                    {
        //                        if (p.Range.Characters[1].Fields.Count > 0) { p.Range.Characters[1].Fields.Unlink(); }
        //                        p.Range.Characters[1].Delete();
        //                    }
        //                }
        //            }
        //            catch (System.ArgumentOutOfRangeException) { }
        //        }
        //    }
        //    if (Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
        //    {
        //        string msg_StatusBar = "Estilo Seção Terciária: Sucesso";
        //        stopwatch.Stop();
        //        msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
        //        Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
        //    }
        //    Globals.ThisAddIn.Application.ScreenUpdating = true;
        //}

        //private void button_secao_4_Click(object sender, EventArgs e)
        //{
        //    Stopwatch stopwatch = new Stopwatch(); if (Ribbon.Variables.debugging) { stopwatch.Start(); }
        //    //string estilo_nome_baseado = "04 - Seções (PeriTAB)";
        //    //string estilo_nome_seguinte = "02 - Corpo do Texto (PeriTAB)";
        //    //string estilo_nome = "04d - Seção_4 (PeriTAB)";
        //    //Globals.ThisAddIn.Application.ScreenUpdating = false;
        //    //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
        //    //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
        //    //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
        //    string estilo_nome = dict_botao_e_estilo[sender as Button];
        //    string estilo_nome_baseado = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_BaseStyle().NameLocal;
        //    string estilo_nome_seguinte = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_NextParagraphStyle().NameLocal;
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.ScreenUpdating = false;

        //    foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
        //    {
        //        p.Range.set_Style((object)estilo_nome);
        //        Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); } } //Deleta parágrafo anterior em branco
        //        if (r != null & Globals.ThisAddIn.Application.Selection.Paragraphs.Count == 1)
        //        {
        //            Microsoft.Office.Interop.Word.Style r_estilo = (Microsoft.Office.Interop.Word.Style)r.get_Style();
        //            if (r_estilo != null) //Ao que parece, paragráfos com o estilo "revisado" perdem o parâmetro de estilo. Esta linha evita este erro.          
        //            {
        //                if (r_estilo.NameLocal.ToString() == "04a - Seção_1 (PeriTAB)" | r_estilo.NameLocal.ToString() == "04b - Seção_2 (PeriTAB)" | r_estilo.NameLocal.ToString() == "04c - Seção_3 (PeriTAB)" | r_estilo.NameLocal.ToString() == "04d - Seção_4 (PeriTAB)")
        //                {
        //                    p.Range.ParagraphFormat.SpaceBefore = 0;
        //                }
        //            }
        //        }

        //        int s;
        //        if (p.Range.Text.Length >= 13) { s = 13; } else { s = p.Range.Text.Length; }
        //        if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count == 1)
        //        {
        //            string a = p.Range.Text.Substring(0, s).Replace(((char)8211).ToString(), "-");
        //                    try
        //                    {
        //                        if (a.Replace(" ", "").Substring(0, 2) == "I-" | a.Replace(" ", "").Substring(0, 3) == "II-" | a.Replace(" ", "").Substring(0, 4) == "III-" | a.Replace(" ", "").Substring(0, 3) == "IV-" | a.Replace(" ", "").Substring(0, 2) == "V-" | a.Replace(" ", "").Substring(0, 3) == "VI-" | a.Replace(" ", "").Substring(0, 4) == "VII-" | a.Replace(" ", "").Substring(0, 5) == "VIII-" | a.Replace(" ", "").Substring(0, 3) == "IX-" | a.Replace(" ", "").Substring(0, 2) == "X-" | a.Substring(0, 2) == "I." | a.Substring(0, 3) == "II." | a.Substring(0, 4) == "III." | a.Substring(0, 3) == "IV." | a.Substring(0, 2) == "V." | a.Substring(0, 3) == "VI." | a.Substring(0, 4) == "VII." | a.Substring(0, 5) == "VIII." | a.Substring(0, 3) == "IX." | a.Substring(0, 2) == "X.")
        //            {
        //                int loc_hifen = a.IndexOf("-");
        //                for (int i = 1; i <= loc_hifen; i++)
        //                {
        //                    if (p.Range.Characters[1].Fields.Count > 0) { p.Range.Characters[1].Fields.Unlink(); }
        //                    p.Range.Characters[1].Delete();
        //                }
        //            }
        //            }
        //            catch (System.ArgumentOutOfRangeException) { }
        //        }
        //    }
        //    if (Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
        //    {
        //        string msg_StatusBar = "Estilo Seção Quaternária: Sucesso";
        //        stopwatch.Stop();
        //        msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
        //        Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
        //    }
        //    Globals.ThisAddIn.Application.ScreenUpdating = true;
        //}

        //private void button_secao_5_Click(object sender, EventArgs e)
        //{
        //    Stopwatch stopwatch = new Stopwatch(); if (Ribbon.Variables.debugging) { stopwatch.Start(); }
        //    //string estilo_nome_baseado = "04 - Seções (PeriTAB)";
        //    //string estilo_nome_seguinte = "02 - Corpo do Texto (PeriTAB)";
        //    //string estilo_nome = "04d - Seção_4 (PeriTAB)";
        //    //Globals.ThisAddIn.Application.ScreenUpdating = false;
        //    //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
        //    //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
        //    //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
        //    string estilo_nome = dict_botao_e_estilo[sender as Button];
        //    string estilo_nome_baseado = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_BaseStyle().NameLocal;
        //    string estilo_nome_seguinte = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_NextParagraphStyle().NameLocal;
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.ScreenUpdating = false;

        //    foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
        //    {
        //        p.Range.set_Style((object)estilo_nome);
        //        Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); } } //Deleta parágrafo anterior em branco
        //        if (r != null & Globals.ThisAddIn.Application.Selection.Paragraphs.Count == 1)
        //        {
        //            Microsoft.Office.Interop.Word.Style r_estilo = (Microsoft.Office.Interop.Word.Style)r.get_Style();
        //            if (r_estilo != null) //Ao que parece, paragráfos com o estilo "revisado" perdem o parâmetro de estilo. Esta linha evita este erro.          
        //            {
        //                if (r_estilo.NameLocal.ToString() == "04a - Seção_1 (PeriTAB)" | r_estilo.NameLocal.ToString() == "04b - Seção_2 (PeriTAB)" | r_estilo.NameLocal.ToString() == "04c - Seção_3 (PeriTAB)" | r_estilo.NameLocal.ToString() == "04d - Seção_4 (PeriTAB)")
        //                {
        //                    p.Range.ParagraphFormat.SpaceBefore = 0;
        //                }
        //            }
        //        }

        //        int s;
        //        if (p.Range.Text.Length >= 13) { s = 13; } else { s = p.Range.Text.Length; }
        //        if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count == 1)
        //        {
        //            string a = p.Range.Text.Substring(0, s).Replace(((char)8211).ToString(), "-");
        //            try
        //            {
        //                if (a.Replace(" ", "").Substring(0, 2) == "I-" | a.Replace(" ", "").Substring(0, 3) == "II-" | a.Replace(" ", "").Substring(0, 4) == "III-" | a.Replace(" ", "").Substring(0, 3) == "IV-" | a.Replace(" ", "").Substring(0, 2) == "V-" | a.Replace(" ", "").Substring(0, 3) == "VI-" | a.Replace(" ", "").Substring(0, 4) == "VII-" | a.Replace(" ", "").Substring(0, 5) == "VIII-" | a.Replace(" ", "").Substring(0, 3) == "IX-" | a.Replace(" ", "").Substring(0, 2) == "X-" | a.Substring(0, 2) == "I." | a.Substring(0, 3) == "II." | a.Substring(0, 4) == "III." | a.Substring(0, 3) == "IV." | a.Substring(0, 2) == "V." | a.Substring(0, 3) == "VI." | a.Substring(0, 4) == "VII." | a.Substring(0, 5) == "VIII." | a.Substring(0, 3) == "IX." | a.Substring(0, 2) == "X.")
        //                {
        //                    int loc_hifen = a.IndexOf("-");
        //                    for (int i = 1; i <= loc_hifen; i++)
        //                    {
        //                        if (p.Range.Characters[1].Fields.Count > 0) { p.Range.Characters[1].Fields.Unlink(); }
        //                        p.Range.Characters[1].Delete();
        //                    }
        //                }
        //            }
        //            catch (System.ArgumentOutOfRangeException) { }
        //        }
        //    }
        //    if (Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
        //    {
        //        string msg_StatusBar = "Estilo Seção Quinária: Sucesso";
        //        stopwatch.Stop();
        //        msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
        //        Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
        //    }
        //    Globals.ThisAddIn.Application.ScreenUpdating = true;
        //}

        private void button_enumeracao_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon.Variables.debugging) { stopwatch.Start(); }
            //string estilo_nome_baseado = "Normal";
            //string estilo_nome = "05 - Enumerações (PeriTAB)";
            //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            //Globals.ThisAddIn.Application.ScreenUpdating = false;
            string estilo_nome = dict_botao_e_estilo[sender as Button];
            string estilo_nome_baseado = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_BaseStyle().NameLocal;
            string estilo_nome_seguinte = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_NextParagraphStyle().NameLocal;
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); }
            if (Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Estilo Enumeração: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        //private void button_reinicia_lista_Click(object sender, EventArgs e)
        //{
        //    Stopwatch stopwatch = new Stopwatch(); if (Ribbon.Variables.debugging) { stopwatch.Start(); }
        //    Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplate(Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListTemplate,(object)false);
        //    //if (Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListValue == 1) { Habilita_button_reinicia_lista(false); }
        //    if (Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListValue == 1) { Habilita_Destaca(MyButton("button_reinicia_lista"), false); }
        //    if (Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
        //    {
        //        string msg_StatusBar = "Reinicia Lista: Sucesso";
        //        stopwatch.Stop();
        //        msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
        //        Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
        //    }
        //}

        private void button_figuras_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon.Variables.debugging) { stopwatch.Start(); }
            //string estilo_nome_baseado = "Normal";
            //string estilo_nome_seguinte = "07 - Legendas de Figuras (PeriTAB)";
            //string estilo_nome = "06 - Figuras (PeriTAB)";
            //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            //Globals.ThisAddIn.Application.ScreenUpdating = false;
            string estilo_nome = dict_botao_e_estilo[sender as Button];
            string estilo_nome_baseado = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_BaseStyle().NameLocal;
            string estilo_nome_seguinte = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_NextParagraphStyle().NameLocal;
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
            {
                p.Range.set_Style((object)estilo_nome);
                Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco
            }
            if (Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Estilo Figuras: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        private void button_legendas_de_figuras_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon.Variables.debugging) { stopwatch.Start(); }
            //string estilo_nome_baseado = "Legenda";
            //string estilo_nome_seguinte = "02 - Corpo do Texto (PeriTAB)";
            //string estilo_nome = "07 - Legendas de Figuras (PeriTAB)";
            //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            //Globals.ThisAddIn.Application.ScreenUpdating = false;
            string estilo_nome = dict_botao_e_estilo[sender as Button];
            string estilo_nome_baseado = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_BaseStyle().NameLocal;
            string estilo_nome_seguinte = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_NextParagraphStyle().NameLocal;
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
            {
                //List<int> positions = new List<int>();
                //for (int i = 1; i <= p.Range.Characters.Count; i++)
                //{
                //    Range characterRange = p.Range.Characters[i];
                //    if (characterRange.Font.Name == "Courier New")
                //    {
                //        positions.Add(i);
                //    }
                //}
                //string positionsOutput = string.Join(", ", positions);
                //MessageBox.Show("Posições dos caracteres com fonte 'Courier New': " + positionsOutput);

                p.Range.set_Style((object)estilo_nome);
                Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco

                //foreach (int pos in positions)
                //{

                //    //MessageBox.Show(pos.ToString());
                //    p.Range.Characters[pos].Font.Size = 30;
                //    //p.Range.Characters[pos].Font.Name = "Courier New";
                //    MessageBox.Show(p.Range.Characters[pos].Font.Name);
                //    //Range rangeToChange = Globals.ThisAddIn.Application.ActiveDocument.Range(pos, pos + 1);
                //    //rangeToChange.Font.Name = "Courier New";

                //    //p.Range.Characters(1).
                //    //    (pos, pos + 1).Font.Name = "Courier New";
                //}
            }
            if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count == 1) { Globals.ThisAddIn.Application.Run("alinha_legenda"); }
            if (Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Estilo Legendas de Figuras: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        //private void button_textos_de_figuras_Click(object sender, EventArgs e)
        //{
        //    Stopwatch stopwatch = new Stopwatch(); if (Ribbon.Variables.debugging) { stopwatch.Start(); }
        //    //string estilo_nome_baseado = "07 - Legendas de Figuras (PeriTAB)";
        //    //string estilo_nome_seguinte = "02 - Corpo do Texto (PeriTAB)";
        //    //string estilo_nome = "08a - Texto de Figuras (PeriTAB)";
        //    //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
        //    //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
        //    //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
        //    //Globals.ThisAddIn.Application.ScreenUpdating = false;
        //    string estilo_nome = dict_botao_e_estilo[sender as Button];
        //    string estilo_nome_baseado = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_BaseStyle().NameLocal;
        //    string estilo_nome_seguinte = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_NextParagraphStyle().NameLocal;
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.ScreenUpdating = false;

        //    foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
        //    {
        //        p.Range.set_Style((object)estilo_nome);
        //        //*****************************************************************************************************************
        //        if (p.Previous() != null)
        //        {
        //            if ((((Microsoft.Office.Interop.Word.Style)p.Previous().get_Style()).NameLocal.ToString()) == "07 - Legendas de Figuras (PeriTAB)")
        //            {
        //                p.Range.ParagraphFormat.LeftIndent = p.Previous().Range.ParagraphFormat.LeftIndent;
        //                p.Range.ParagraphFormat.RightIndent = p.Previous().Range.ParagraphFormat.RightIndent;
        //                p.Previous().Range.ParagraphFormat.SpaceAfter = 0;
        //                //MessageBox.Show(p.Previous().Range.ParagraphFormat.KeepWithNext.ToString());
        //                p.Previous().Range.ParagraphFormat.KeepWithNext = -1;
        //            }
        //        }
        //    }
        //    if (Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
        //    {
        //        string msg_StatusBar = "Estilo Texto de Figuras: Sucesso";
        //        stopwatch.Stop();
        //        msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
        //        Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
        //    }
        //    Globals.ThisAddIn.Application.ScreenUpdating = true;
        //}


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

        private void button_legendas_de_tabelas_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon.Variables.debugging) { stopwatch.Start(); }
            //string estilo_nome_baseado = "Legenda";
            //string estilo_nome_seguinte = "01 - Sem Formatação (PeriTAB)";
            //string estilo_nome = "08 - Legendas de Tabelas (PeriTAB)";
            //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            //Globals.ThisAddIn.Application.ScreenUpdating = false;
            string estilo_nome = dict_botao_e_estilo[sender as Button];
            string estilo_nome_baseado = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_BaseStyle().NameLocal;
            string estilo_nome_seguinte = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_NextParagraphStyle().NameLocal;
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) 
            { 
                p.Range.set_Style((object)estilo_nome);
                Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco
            }
            if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count == 1) { Globals.ThisAddIn.Application.Run("alinha_legenda"); }
            if (Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Estilo Legendas de Tabelas: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }        

        private void button_quesitos_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon.Variables.debugging) { stopwatch.Start(); }
            //string estilo_nome_baseado = "02 - Corpo do Texto (PeriTAB)";
            //string estilo_nome = "09 - Quesitos (PeriTAB)";
            //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            //Globals.ThisAddIn.Application.ScreenUpdating = false;
            string estilo_nome = dict_botao_e_estilo[sender as Button];
            string estilo_nome_baseado = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_BaseStyle().NameLocal;
            string estilo_nome_seguinte = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_NextParagraphStyle().NameLocal;
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) 
            {
                p.Range.set_Style((object)estilo_nome);
                Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); } } //Deleta parágrafo anterior em branco
                if (r != null & Globals.ThisAddIn.Application.Selection.Paragraphs.Count == 1)
                {
                    Microsoft.Office.Interop.Word.Style r_estilo = (Microsoft.Office.Interop.Word.Style)r.get_Style();                   
                    if (r_estilo != null) //Ao que parece, paragráfos com o estilo "revisado" perdem o parâmetro de estilo. Esta linha evita este erro.          
                    {
                        if (r_estilo.NameLocal.ToString() == "04a - Seção_1 (PeriTAB)" | r_estilo.NameLocal.ToString() == "04b - Seção_2 (PeriTAB)" | r_estilo.NameLocal.ToString() == "04c - Seção_3 (PeriTAB)" | r_estilo.NameLocal.ToString() == "04d - Seção_4 (PeriTAB)")
                            {
                            p.Range.ParagraphFormat.SpaceBefore = 0;
                        }               
                    }
                }
            }
            //Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco
            if (Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Estilo Quesitos: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void button_fecho_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon.Variables.debugging) { stopwatch.Start(); }
            //string estilo_nome_baseado = "02 - Corpo do Texto (PeriTAB)";
            //string estilo_nome = "10 - Fecho (PeriTAB)";
            //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            //Globals.ThisAddIn.Application.ScreenUpdating = false;
            string estilo_nome = dict_botao_e_estilo[sender as Button];
            string estilo_nome_baseado = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_BaseStyle().NameLocal;
            string estilo_nome_seguinte = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_NextParagraphStyle().NameLocal;
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
            {
                p.Range.set_Style((object)estilo_nome);
                Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); } } //Deleta parágrafo anterior em branco
            }
            if (Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Estilo Fecho: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void button_notas_de_rodape_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon.Variables.debugging) { stopwatch.Start(); }
            //string estilo_nome_baseado = "Texto de nota de rodapé";
            //string estilo_nome = "17 - Notas de rodapé (PeriTAB)";
            //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            //Globals.ThisAddIn.Application.ScreenUpdating = false;
            string estilo_nome = dict_botao_e_estilo[sender as Button];
            string estilo_nome_baseado = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_BaseStyle().NameLocal;
            string estilo_nome_seguinte = Globals.ThisAddIn.Application.ActiveDocument.Styles[estilo_nome].get_NextParagraphStyle().NameLocal;
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
            {
                p.Range.set_Style((object)estilo_nome);
            }
            if (Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Estilo Notas de rodapé: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        //public Button MyButton(string nome_botao)
        //{
        //    foreach (Button botao in Controls)
        //    {
        //        if (botao.Name == nome_botao) return botao;
        //    }
        //    MessageBox.Show("não achou o botao: " + nome_botao);
        //    return null;
        //}

        // Método para buscar botão pelo nome
        public Button MyButton(string nomeBotao)
        {
            var botao = Controls.OfType<Button>().FirstOrDefault(b => b.Name == nomeBotao);
            //if (botao == null)
            //{
            //    MessageBox.Show($"Botão '{nomeBotao}' não encontrado.");
            //}
            return botao;
        }

        public void Habilita_Destaca(Button b, bool habilita, bool destaca = false)
        {
            //if (b == null) MessageBox.Show("possivel erro de pintar o painel");
            b.Enabled = habilita;
            if (destaca) { b.BackColor = SystemColors.Highlight; b.ForeColor = SystemColors.HighlightText; }
        }
        //internal void Remove_Destaque_Botoes(MyUserControl UCs)
        //{
        //    foreach (Button b in UCs.Controls)
        //    {
        //        if ((b.GetType()).Name == "Button")
        //        {
        //            b.BackColor = SystemColors.Control;
        //            b.ForeColor = SystemColors.ControlText;
        //        }

        //    }
        //}

        public void Remove_Destaque_Botoes(MyUserControl UserControl)
        {
            foreach (var botao in UserControl.Controls.OfType<Button>())
            {
                botao.BackColor = SystemColors.Control;
                botao.ForeColor = SystemColors.ControlText;
            }
        }

        private void MyUserControl_Load(object sender, EventArgs e)
        {

        }



        //// Dicionário estático, inicializado uma vez para todos os usos.
        //public static readonly Dictionary<string, string> dict_estilo_e_botao = new Dictionary<string, string>
        //{
        //    { "01 - Sem Formatação (PeriTAB)", "button_sem_formatacao" },
        //    { "02 - Corpo do Texto (PeriTAB)", "button_corpo_do_texto" },
        //    { "03 - Citações (PeriTAB)", "button_citacoes" },
        //    { "04a - Seção_1 (PeriTAB)", "button_secao_1" },
        //    { "04b - Seção_2 (PeriTAB)", "button_secao_2" },
        //    { "04c - Seção_3 (PeriTAB)", "button_secao_3" },
        //    { "04d - Seção_4 (PeriTAB)", "button_secao_4" },
        //    { "05 - Enumerações (PeriTAB)", "button_enumeracao" },
        //    { "06 - Figuras (PeriTAB)", "button_figuras" },
        //    { "07 - Legendas de Figuras (PeriTAB)", "button_legendas_de_figuras" },
        //    { "08a - Texto de Figuras (PeriTAB)", "button_textos_de_figuras" },
        //    { "08 - Legendas de Tabelas (PeriTAB)", "button_legendas_de_tabelas" },
        //    { "09 - Quesitos (PeriTAB)", "button_quesitos" },
        //    { "10 - Fecho (PeriTAB)", "button_fecho" },
        //    { "11 - Parágrafo Numerado (PeriTAB)", "button_paragrafo_numerado" }
        //};






    }
}
