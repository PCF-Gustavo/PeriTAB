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

namespace PeriTAB
{
    public partial class MyUserControl : UserControl
    {

        public MyUserControl()
        {
            InitializeComponent();
        }
        private void button_sem_formatacao_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon1.Variables.debugging) { stopwatch.Start(); }
            string estilo_nome_baseado = "Normal";
            string estilo_nome = "01 - Sem Formatação (PeriTAB)";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs){ p.Range.set_Style((object)estilo_nome); }
            if (Ribbon1.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Estilo Sem Formatação: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void button_corpo_do_texto_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon1.Variables.debugging) { stopwatch.Start(); }
            string estilo_nome_baseado = "Normal";
            string estilo_nome = "02 - Corpo do Texto (PeriTAB)";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); }
            if (Ribbon1.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
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
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon1.Variables.debugging) { stopwatch.Start(); }
            string estilo_nome_baseado = "Normal";
            string estilo_nome = "11 - Parágrafo Numerado (PeriTAB)";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); }
            if (Ribbon1.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
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
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon1.Variables.debugging) { stopwatch.Start(); }
            string estilo_nome_baseado = "Normal";
            string estilo_nome = "03 - Citações (PeriTAB)";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); }
            Range r1 = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r1 != null) { if (r1.Text == ((char)13).ToString()) { r1.Delete(); } } //Deleta parágrafo anterior em branco
            Range r2 = Globals.ThisAddIn.Application.Selection.Next(WdUnits.wdParagraph, 1); if (r2 != null) { if (r2.Text == ((char)13).ToString()) { r2.Delete(); } } //Deleta parágrafo seguinte em branco
            if (Ribbon1.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
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
        //Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
        //Globals.ThisAddIn.Application.ScreenUpdating = false;
        //foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); p.Range.SetListLevel(1); }
        //Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco
        //Globals.ThisAddIn.Application.ScreenUpdating = true;
        //}
        //private void button_secao_2_Click(object sender, EventArgs e)
        //{
        //    string estilo_nome = "04 - Seções (PeriTAB)";
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.ScreenUpdating = false;
        //    foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); p.Range.SetListLevel(2); p.Range.Font.AllCaps = 0; }
        //    Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco
        //    Globals.ThisAddIn.Application.ScreenUpdating = true;
        //}

        //private void button_secao_3_Click(object sender, EventArgs e)
        //{
        //    string estilo_nome = "04 - Seções (PeriTAB)";
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.ScreenUpdating = false;
        //    foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); p.Range.SetListLevel(3); p.Range.Font.AllCaps = 0; p.Range.Font.Bold = 0; p.Range.Font.Italic = -1; }
        //    Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco
        //    Globals.ThisAddIn.Application.ScreenUpdating = true;
        //}

        //private void button_secao_4_Click(object sender, EventArgs e)
        //{
        //    string estilo_nome = "04 - Seções (PeriTAB)";
        //    Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
        //    Globals.ThisAddIn.Application.ScreenUpdating = false;
        //    foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); p.Range.SetListLevel(4); p.Range.Font.AllCaps = 0; p.Range.Font.Bold = 0; p.Range.Font.Underline = WdUnderline.wdUnderlineSingle; }
        //    Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco
        //    Globals.ThisAddIn.Application.ScreenUpdating = true;
        //}
        private void button_secao_1_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon1.Variables.debugging) { stopwatch.Start(); }
            string estilo_nome_baseado = "04 - Seções (PeriTAB)";
            string estilo_nome_seguinte = "02 - Corpo do Texto (PeriTAB)";
            string estilo_nome = "04a - Seção_1 (PeriTAB)";
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) 
            {


                p.Range.set_Style((object)estilo_nome); 
                Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco

                int s;
                if (p.Range.Text.Length >= 7) { s = 7; } else { s = p.Range.Text.Length; }
                if (s == 0) break;
                if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count == 1) 
                {
                    string a = p.Range.Text.Substring(0, s).Replace(((char)8211).ToString(), "-");
                    try
                    {
                        if (a.Replace(" ", "").Substring(0, 2) == "I-" | a.Replace(" ", "").Substring(0, 3) == "II-" | a.Replace(" ", "").Substring(0, 4) == "III-" | a.Replace(" ", "").Substring(0, 3) == "IV-" | a.Replace(" ", "").Substring(0, 2) == "V-" | a.Replace(" ", "").Substring(0, 3) == "VI-" | a.Replace(" ", "").Substring(0, 4) == "VII-" | a.Replace(" ", "").Substring(0, 5) == "VIII-" | a.Replace(" ", "").Substring(0, 3) == "IX-" | a.Replace(" ", "").Substring(0, 2) == "X-" | a.Substring(0, 2) == "I." | a.Substring(0, 3) == "II." | a.Substring(0, 4) == "III." | a.Substring(0, 3) == "IV." | a.Substring(0, 2) == "V." | a.Substring(0, 3) == "VI." | a.Substring(0, 4) == "VII." | a.Substring(0, 5) == "VIII." | a.Substring(0, 3) == "IX." | a.Substring(0, 2) == "X.")
                        {
                            //MessageBox.Show(a);
                            int loc_hifen = a.IndexOf("-");
                            //MessageBox.Show(loc_hifen.ToString());
                            for (int i = 1; i <= loc_hifen; i++)
                            {
                                //MessageBox.Show(p.Range.Characters[1].Text);
                                if (p.Range.Characters[1].Fields.Count > 0) { p.Range.Characters[1].Fields.Unlink(); }
                                p.Range.Characters[1].Delete();
                            }
                            //break;
                        }
                    }
                    catch (System.ArgumentOutOfRangeException) { }
                }


                ////string a = p.Range.Text.Substring(0, s).Replace(((char)8211).ToString(), "-").Replace(".", "").Replace("0", "").Replace("1", "").Replace("2", "").Replace("3", "").Replace("4", "").Replace("5", "").Replace("6", "").Replace("7", "").Replace("8", "").Replace("9", "");
                //int loc_hifen = a.IndexOf("-");
                //MessageBox.Show(a);
                //for (int i = 1; i <= loc_hifen; i++)
                //{
                //    MessageBox.Show(p.Range.Characters[1].Text);
                //    p.Range.Characters[1].Delete();
                //}
                //if (loc_hifen != -1)
                //{
                //    //MessageBox.Show(a.Substring(0, 2).Replace(" ", ""));
                //    if (a.Replace(" ", "").Substring(0, 2) == "I-" | a.Substring(0, 2) == "II" | a.Substring(0, 2) == "IV" | a.Replace(" ", "").Substring(0, 2) == "V-" | a.Substring(0, 2) == "VI")
                //    {
                //        for (int i = 1; i <= loc_hifen; i++)
                //        {
                //            //MessageBox.Show(p.Range.Characters[1].Text);
                //            p.Range.Characters[1].Delete();
                //        }
                //    }
                //}
                //Replace(((char)8211).ToString(), "-")
                //if ((p.Range.Text.Substring(0, 12)).IndexOf(Convert.ToChar(150)) != -1)
                //{
                //    for (int i = 0; i <= (p.Range.Text.Substring(0, 12)).IndexOf(Convert.ToChar(150)) - 1; i++)
                //    {
                //        MessageBox.Show(p.Range.Characters[i].ToString());
                //        p.Range.Characters[i].Delete();
                //    }
                //}
            }
            if (Ribbon1.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Estilo Seção Primária: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void button_secao_2_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon1.Variables.debugging) { stopwatch.Start(); }
            string estilo_nome_baseado = "04 - Seções (PeriTAB)";
            string estilo_nome_seguinte = "02 - Corpo do Texto (PeriTAB)";
            string estilo_nome = "04b - Seção_2 (PeriTAB)";
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            
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

                int s;
                if (p.Range.Text.Length >= 9) { s = 9; } else { s = p.Range.Text.Length; }
                if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count == 1)
                {
                    string a = p.Range.Text.Substring(0, s).Replace(((char)8211).ToString(), "-");
                    try
                    {
                        if (a.Replace(" ", "").Substring(0, 2) == "I-" | a.Replace(" ", "").Substring(0, 3) == "II-" | a.Replace(" ", "").Substring(0, 4) == "III-" | a.Replace(" ", "").Substring(0, 3) == "IV-" | a.Replace(" ", "").Substring(0, 2) == "V-" | a.Replace(" ", "").Substring(0, 3) == "VI-" | a.Replace(" ", "").Substring(0, 4) == "VII-" | a.Replace(" ", "").Substring(0, 5) == "VIII-" | a.Replace(" ", "").Substring(0, 3) == "IX-" | a.Replace(" ", "").Substring(0, 2) == "X-" | a.Substring(0, 2) == "I." | a.Substring(0, 3) == "II." | a.Substring(0, 4) == "III." | a.Substring(0, 3) == "IV." | a.Substring(0, 2) == "V." | a.Substring(0, 3) == "VI." | a.Substring(0, 4) == "VII." | a.Substring(0, 5) == "VIII." | a.Substring(0, 3) == "IX." | a.Substring(0, 2) == "X.")
                    {
                        int loc_hifen = a.IndexOf("-");
                        for (int i = 1; i <= loc_hifen; i++)
                        {
                            if (p.Range.Characters[1].Fields.Count > 0) { p.Range.Characters[1].Fields.Unlink(); }
                            p.Range.Characters[1].Delete();
                        }
                    }
                    }
                    catch (System.ArgumentOutOfRangeException) { }
                }
            }
            if (Ribbon1.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Estilo Seção Secundária: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void button_secao_3_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon1.Variables.debugging) { stopwatch.Start(); }
            string estilo_nome_baseado = "04 - Seções (PeriTAB)";
            string estilo_nome_seguinte = "02 - Corpo do Texto (PeriTAB)";
            string estilo_nome = "04c - Seção_3 (PeriTAB)";
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            
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

                int s;
                if (p.Range.Text.Length >= 11) { s = 11; } else { s = p.Range.Text.Length; }
                if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count == 1)
                {
                    string a = p.Range.Text.Substring(0, s).Replace(((char)8211).ToString(), "-");
                        try
                        {
                            if (a.Replace(" ", "").Substring(0, 2) == "I-" | a.Replace(" ", "").Substring(0, 3) == "II-" | a.Replace(" ", "").Substring(0, 4) == "III-" | a.Replace(" ", "").Substring(0, 3) == "IV-" | a.Replace(" ", "").Substring(0, 2) == "V-" | a.Replace(" ", "").Substring(0, 3) == "VI-" | a.Replace(" ", "").Substring(0, 4) == "VII-" | a.Replace(" ", "").Substring(0, 5) == "VIII-" | a.Replace(" ", "").Substring(0, 3) == "IX-" | a.Replace(" ", "").Substring(0, 2) == "X-" | a.Substring(0, 2) == "I." | a.Substring(0, 3) == "II." | a.Substring(0, 4) == "III." | a.Substring(0, 3) == "IV." | a.Substring(0, 2) == "V." | a.Substring(0, 3) == "VI." | a.Substring(0, 4) == "VII." | a.Substring(0, 5) == "VIII." | a.Substring(0, 3) == "IX." | a.Substring(0, 2) == "X.")
                    {
                        int loc_hifen = a.IndexOf("-");
                        for (int i = 1; i <= loc_hifen; i++)
                        {
                            if (p.Range.Characters[1].Fields.Count > 0) { p.Range.Characters[1].Fields.Unlink(); }
                            p.Range.Characters[1].Delete();
                        }
                    }
                    }
                    catch (System.ArgumentOutOfRangeException) { }
                }
            }
            if (Ribbon1.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Estilo Seção Terciária: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void button_secao_4_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon1.Variables.debugging) { stopwatch.Start(); }
            string estilo_nome_baseado = "04 - Seções (PeriTAB)";
            string estilo_nome_seguinte = "02 - Corpo do Texto (PeriTAB)";
            string estilo_nome = "04d - Seção_4 (PeriTAB)";
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            
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

                int s;
                if (p.Range.Text.Length >= 13) { s = 13; } else { s = p.Range.Text.Length; }
                if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count == 1)
                {
                    string a = p.Range.Text.Substring(0, s).Replace(((char)8211).ToString(), "-");
                            try
                            {
                                if (a.Replace(" ", "").Substring(0, 2) == "I-" | a.Replace(" ", "").Substring(0, 3) == "II-" | a.Replace(" ", "").Substring(0, 4) == "III-" | a.Replace(" ", "").Substring(0, 3) == "IV-" | a.Replace(" ", "").Substring(0, 2) == "V-" | a.Replace(" ", "").Substring(0, 3) == "VI-" | a.Replace(" ", "").Substring(0, 4) == "VII-" | a.Replace(" ", "").Substring(0, 5) == "VIII-" | a.Replace(" ", "").Substring(0, 3) == "IX-" | a.Replace(" ", "").Substring(0, 2) == "X-" | a.Substring(0, 2) == "I." | a.Substring(0, 3) == "II." | a.Substring(0, 4) == "III." | a.Substring(0, 3) == "IV." | a.Substring(0, 2) == "V." | a.Substring(0, 3) == "VI." | a.Substring(0, 4) == "VII." | a.Substring(0, 5) == "VIII." | a.Substring(0, 3) == "IX." | a.Substring(0, 2) == "X.")
                    {
                        int loc_hifen = a.IndexOf("-");
                        for (int i = 1; i <= loc_hifen; i++)
                        {
                            if (p.Range.Characters[1].Fields.Count > 0) { p.Range.Characters[1].Fields.Unlink(); }
                            p.Range.Characters[1].Delete();
                        }
                    }
                    }
                    catch (System.ArgumentOutOfRangeException) { }
                }
            }
            if (Ribbon1.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Estilo Seção Quaternária: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void button_enumeracao_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon1.Variables.debugging) { stopwatch.Start(); }
            string estilo_nome_baseado = "Normal";
            string estilo_nome = "05 - Enumerações (PeriTAB)";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); }
            if (Ribbon1.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Estilo Enumeração: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void button_reinicia_lista_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon1.Variables.debugging) { stopwatch.Start(); }
            Globals.ThisAddIn.Application.Selection.Range.ListFormat.ApplyListTemplate(Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListTemplate,(object)false);
            //if (Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListValue == 1) { Habilita_button_reinicia_lista(false); }
            if (Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListValue == 1) { Habilita_Destaca(MyButton("button_reinicia_lista"), false); }
            if (Ribbon1.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Reinicia Lista: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
        }

        private void button_figuras_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon1.Variables.debugging) { stopwatch.Start(); }
            string estilo_nome_baseado = "Normal";
            string estilo_nome_seguinte = "07 - Legendas de Figuras (PeriTAB)";
            string estilo_nome = "06 - Figuras (PeriTAB)";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
            {
                p.Range.set_Style((object)estilo_nome);
                Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco
            }
            if (Ribbon1.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
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
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon1.Variables.debugging) { stopwatch.Start(); }
            string estilo_nome_baseado = "Legenda";
            string estilo_nome_seguinte = "02 - Corpo do Texto (PeriTAB)";
            string estilo_nome = "07 - Legendas de Figuras (PeriTAB)";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
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
            if (Ribbon1.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Estilo Legendas de Figuras: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void button_textos_de_figuras_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon1.Variables.debugging) { stopwatch.Start(); }
            string estilo_nome_baseado = "07 - Legendas de Figuras (PeriTAB)";
            string estilo_nome_seguinte = "02 - Corpo do Texto (PeriTAB)";
            string estilo_nome = "08a - Texto de Figuras (PeriTAB)";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
            {
                p.Range.set_Style((object)estilo_nome);

                if (p.Previous() != null)
                {
                    if ((((Microsoft.Office.Interop.Word.Style)p.Previous().get_Style()).NameLocal.ToString()) == "07 - Legendas de Figuras (PeriTAB)")
                    {
                        p.Range.ParagraphFormat.LeftIndent = p.Previous().Range.ParagraphFormat.LeftIndent;
                        p.Range.ParagraphFormat.RightIndent = p.Previous().Range.ParagraphFormat.RightIndent;
                        p.Previous().Range.ParagraphFormat.SpaceAfter = 0;
                        //MessageBox.Show(p.Previous().Range.ParagraphFormat.KeepWithNext.ToString());
                        p.Previous().Range.ParagraphFormat.KeepWithNext = -1;
                    }
                }
            }
            if (Ribbon1.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Estilo Texto de Figuras: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void button_legendas_de_tabelas_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon1.Variables.debugging) { stopwatch.Start(); }
            string estilo_nome_baseado = "Legenda";
            string estilo_nome_seguinte = "01 - Sem Formatação (PeriTAB)";
            string estilo_nome = "08 - Legendas de Tabelas (PeriTAB)";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_seguinte, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) 
            { 
                p.Range.set_Style((object)estilo_nome);
                Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco
            }
            if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count == 1) { Globals.ThisAddIn.Application.Run("alinha_legenda"); }
            if (Ribbon1.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
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
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon1.Variables.debugging) { stopwatch.Start(); }
            string estilo_nome_baseado = "02 - Corpo do Texto (PeriTAB)";
            string estilo_nome = "09 - Quesitos (PeriTAB)";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
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
            if (Ribbon1.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
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
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon1.Variables.debugging) { stopwatch.Start(); }
            string estilo_nome_baseado = "02 - Corpo do Texto (PeriTAB)";
            string estilo_nome = "10 - Fecho (PeriTAB)";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
            {
                p.Range.set_Style((object)estilo_nome);
                Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); } } //Deleta parágrafo anterior em branco
            }
            if (Ribbon1.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                string msg_StatusBar = "Estilo Fecho: Sucesso";
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void button_DockRight_Click(object sender, EventArgs e)
        {
            //    Globals.ThisAddIn.TaskPane1.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
            //    Globals.ThisAddIn.TaskPane1.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            //    Globals.ThisAddIn.TaskPane1.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
            //    this.Size = new System.Drawing.Size(100, 900);
            //    this.button_DockRight.Visible = false;
            //    this.button_sem_formatacao.Location = new System.Drawing.Point(5, 5);
            //    this.button_corpo_do_texto.Location = new System.Drawing.Point(5, 50);
            //    this.button_citacoes.Location = new System.Drawing.Point(5, 95);
            //    this.button_secao_1.Location = new System.Drawing.Point(5, 140);
            //    this.button_secao_2.Location = new System.Drawing.Point(5, 185);
            //    this.button_secao_3.Location = new System.Drawing.Point(5, 230);
            //    this.button_secao_4.Location = new System.Drawing.Point(5, 275);
            //    this.button_enumeracao.Location = new System.Drawing.Point(5, 320);
            //    this.button_reinicia_lista.Size = new System.Drawing.Size(90, 24);
            //    this.button_reinicia_lista.Location = new System.Drawing.Point(5, 359);
            //    this.button_figuras.Location = new System.Drawing.Point(5, 388);
            //    this.button_legendas_de_figuras.Location = new System.Drawing.Point(5, 433);
            //    this.button_legendas_de_tabelas.Location = new System.Drawing.Point(5, 478);
            //    this.button_quesitos.Location = new System.Drawing.Point(5, 523);
            //    this.button_DockBottom.Location = new System.Drawing.Point(30, 568);
            //    this.button_DockBottom.Visible = true;
            //    Globals.ThisAddIn.TaskPane1.Width = 120;
            //    this.Size = new System.Drawing.Size(100, 900);
        }

    private void button_DockBottom_Click(object sender, EventArgs e)
        {
            //    Globals.ThisAddIn.TaskPane1.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
            //    Globals.ThisAddIn.TaskPane1.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
            //    Globals.ThisAddIn.TaskPane1.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
            //    this.Size = new System.Drawing.Size(1400, 100);
            //    this.button_DockBottom.Visible = false;
            //    this.button_DockRight.Visible = true;
            //    this.button_DockRight.Location = new System.Drawing.Point(1204, 5);
            //    this.button_sem_formatacao.Location = new System.Drawing.Point(5, 5);
            //    this.button_corpo_do_texto.Location = new System.Drawing.Point(100, 5);
            //    this.button_citacoes.Location = new System.Drawing.Point(195, 5);
            //    this.button_secao_1.Location = new System.Drawing.Point(290, 5);
            //    this.button_secao_2.Location = new System.Drawing.Point(385, 5);
            //    this.button_secao_3.Location = new System.Drawing.Point(480, 5);
            //    this.button_secao_4.Location = new System.Drawing.Point(575, 5);
            //    this.button_enumeracao.Location = new System.Drawing.Point(670, 5);
            //    this.button_reinicia_lista.Size = new System.Drawing.Size(60, 40);
            //    this.button_reinicia_lista.Location = new System.Drawing.Point(759, 5);
            //    this.button_figuras.Location = new System.Drawing.Point(824, 5);
            //    this.button_legendas_de_figuras.Location = new System.Drawing.Point(919, 5);
            //    this.button_legendas_de_tabelas.Location = new System.Drawing.Point(1014, 5);
            //    this.button_quesitos.Location = new System.Drawing.Point(1109, 5);                  
            //    this.Size = new System.Drawing.Size(1400, 100);
            //    Globals.ThisAddIn.TaskPane1.Height = 90;
        }

        public Button MyButton(string nome_botao)
        {
            foreach (Button botao in Controls)
            {
                if (botao.Name == nome_botao) return botao;
            }
            MessageBox.Show("não achou o botao: " + nome_botao);
            return null;
        }

        public void Habilita_Destaca(Button b, bool habilita, bool destaca = false)
        {
            //if (b == null) MessageBox.Show("possivel erro de pintar o painel");
            b.Enabled = habilita;
            if (destaca) { b.BackColor = SystemColors.Highlight; b.ForeColor = SystemColors.HighlightText; }
        }
        internal void Remove_Destaque_Botoes(MyUserControl UCs)
        {
            foreach (Button b in UCs.Controls)
            {
                if ((b.GetType()).Name == "Button")
                {
                    b.BackColor = SystemColors.Control;
                    b.ForeColor = SystemColors.ControlText;
                }

            }
        }





        //public void Habilita_Destaca_button1(bool habilita, bool destaca = false)
        //{
        //    button_sem_formatacao.Enabled = habilita;
        //    if (destaca){button_sem_formatacao.BackColor = SystemColors.Highlight;button_sem_formatacao.ForeColor = SystemColors.HighlightText;}
        //}
        //public void Habilita_Destaca_button2(bool habilita, bool destaca = false)
        //{
        //    button_corpo_do_texto.Enabled = habilita;
        //    if (destaca) { button_corpo_do_texto.BackColor = SystemColors.Highlight; button_corpo_do_texto.ForeColor = SystemColors.HighlightText; }
        //}
        //public void Habilita_Destaca_button3(bool habilita, bool destaca = false)
        //{
        //    button_citacoes.Enabled = habilita;
        //    if (destaca) { button_citacoes.BackColor = SystemColors.Highlight; button_citacoes.ForeColor = SystemColors.HighlightText; }
        //}
        //public void Habilita_Destaca_button4(bool habilita, bool destaca = false)
        //{
        //    button_secao_1.Enabled = habilita;
        //    if (destaca) { button_secao_1.BackColor = SystemColors.Highlight; button_secao_1.ForeColor = SystemColors.HighlightText; }
        //}
        //public void Habilita_Destaca_button5(bool habilita, bool destaca = false)
        //{
        //    button_secao_2.Enabled = habilita;
        //    if (destaca) { button_secao_2.BackColor = SystemColors.Highlight; button_secao_2.ForeColor = SystemColors.HighlightText; }
        //}
        //public void Habilita_Destaca_button6(bool habilita, bool destaca = false)
        //{
        //    button_secao_3.Enabled = habilita;
        //    if (destaca) { button_secao_3.BackColor = SystemColors.Highlight; button_secao_3.ForeColor = SystemColors.HighlightText; }
        //}
        //public void Habilita_Destaca_button7(bool habilita, bool destaca = false)
        //{
        //    button_secao_4.Enabled = habilita;
        //    if (destaca) { button_secao_4.BackColor = SystemColors.Highlight; button_secao_4.ForeColor = SystemColors.HighlightText; }
        //}
        //public void Habilita_Destaca_button8(bool habilita, bool destaca = false)
        //{
        //    button_enumeracao.Enabled = habilita;
        //    if (destaca) { button_enumeracao.BackColor = SystemColors.Highlight; button_enumeracao.ForeColor = SystemColors.HighlightText; }
        //}
        //public void Habilita_button9(bool habilita)
        //{
        //    button_reinicia_lista.Enabled = habilita;
        //}
        //public void Habilita_Destaca_button10(bool habilita, bool destaca = false)
        //{
        //    button_figuras.Enabled = habilita;
        //    if (destaca) { button_figuras.BackColor = SystemColors.Highlight; button_figuras.ForeColor = SystemColors.HighlightText; }
        //}
        //public void Habilita_Destaca_button11(bool habilita, bool destaca = false)
        //{
        //    button_legendas_de_figuras.Enabled = habilita;
        //    if (destaca) { button_legendas_de_figuras.BackColor = SystemColors.Highlight; button_legendas_de_figuras.ForeColor = SystemColors.HighlightText; }
        //}
        //public void Habilita_Destaca_button12(bool habilita, bool destaca = false)
        //{
        //    button_legendas_de_tabelas.Enabled = habilita;
        //    if (destaca) { button_legendas_de_tabelas.BackColor = SystemColors.Highlight; button_legendas_de_tabelas.ForeColor = SystemColors.HighlightText; }
        //}
        //public void Habilita_Destaca_button13(bool habilita, bool destaca = false)
        //{
        //    button_quesitos.Enabled = habilita;
        //    if (destaca) { button_quesitos.BackColor = SystemColors.Highlight; button_quesitos.ForeColor = SystemColors.HighlightText; }
        //}
        //public void Habilita_Destaca_button14(bool habilita, bool destaca = false)
        //{
        //    button_fecho.Enabled = habilita;
        //    if (destaca) { button_fecho.BackColor = SystemColors.Highlight; button_fecho.ForeColor = SystemColors.HighlightText; }
        //}

        //public void Habilita_button_reinicia_lista(bool habilita)
        //{
        //    button_reinicia_lista.Enabled = habilita;
        //}




    }
}
