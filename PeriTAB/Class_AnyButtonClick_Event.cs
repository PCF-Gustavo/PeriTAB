using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;

namespace PeriTAB
{
    internal class Class_AnyButtonClick_Event
    {
        public void Evento_AnyButtonClick(MyUserControl UC)
        {            
            //foreach (RibbonGroup g in Globals.Ribbons.Ribbon1.tab.Groups) //Loop botoes do Ribbon
            //{
            //    foreach (RibbonControl c in g.Items)
            //    {
            //        if ((c.GetType()).Name == "RibbonButtonImpl")
            //        {
            //            RibbonButton b = (RibbonButton)c;
            //            b.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Metodo_AnyButtonClick_Ribbon);
            //        }
            //    }
            //}
            if (Globals.ThisAddIn.Application.Documents.Count > 0 & Globals.ThisAddIn.iMyUserControl != null)
            {
                //MyUserControl MUC = Globals.ThisAddIn.Dicionario_Doc_e_UserControl[Globals.ThisAddIn.Application.ActiveDocument];
                foreach (System.Windows.Forms.Button b1 in UC.Controls) //Loop botoes do Taskpane
                {
                    if ((b1.GetType()).Name == "Button")
                    {
                        b1.Click += new System.EventHandler(Metodo_AnyButtonClick_TaskPane);
                    }
                }
            }
        }

        //private void Metodo_AnyButtonClick_Ribbon(object sender, RibbonControlEventArgs e)
        //{
        //}
        
        private void Metodo_AnyButtonClick_TaskPane(object sender, EventArgs e)
        {
            //MessageBox.Show("AnyButtonClick_TaskPane");
            //Revisa a habilitação do botao "Reinicia Lista" do TaskPane
            if (Globals.ThisAddIn.CustomTaskPanes.Count > 0)
            {
                Globals.ThisAddIn.iMyUserControl.Habilita_button9(true);
                if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count > 1 | Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListType == WdListType.wdListNoNumbering | Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListValue == 1) { Globals.ThisAddIn.iMyUserControl.Habilita_button9(false); }
            }

            //Revisa o destaque dos botoes do TaskPane
            if (Globals.ThisAddIn.CustomTaskPanes.Count > 0){
                Stopwatch stopWatch = new Stopwatch(); stopWatch.Start(); //inicia cronometro
                MyUserControl MUC = Globals.ThisAddIn.Dicionario_Doc_e_UserControl[Globals.ThisAddIn.Application.ActiveDocument];
                Globals.ThisAddIn.iMyUserControl.Remove_Destaque_Botoes(MUC);
                foreach (Microsoft.Office.Interop.Word.Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
                {
                    if (stopWatch.Elapsed.TotalSeconds > 0.2) break; //limita tempo de processamento
                    Microsoft.Office.Interop.Word.Style s = null;
                    try { s = p.Range.get_Style(); } catch (System.Runtime.InteropServices.COMException ex) { }
                    if (s != null)
                    {
                        if (s.NameLocal == "01 - Sem Formatação (PeriTAB)") MUC.Habilita_Destaca_button1(true, true);
                        if (s.NameLocal == "02 - Corpo do Texto (PeriTAB)") MUC.Habilita_Destaca_button2(true, true);
                        if (s.NameLocal == "03 - Citações (PeriTAB)") MUC.Habilita_Destaca_button3(true, true);
                        if (s.NameLocal == "04a - Seção_1 (PeriTAB)") MUC.Habilita_Destaca_button4(true, true);
                        if (s.NameLocal == "04b - Seção_2 (PeriTAB)") MUC.Habilita_Destaca_button5(true, true);
                        if (s.NameLocal == "04c - Seção_3 (PeriTAB)") MUC.Habilita_Destaca_button6(true, true);
                        if (s.NameLocal == "04d - Seção_4 (PeriTAB)") MUC.Habilita_Destaca_button7(true, true);
                        if (s.NameLocal == "05 - Enumerações (PeriTAB)") MUC.Habilita_Destaca_button8(true, true);
                        if (s.NameLocal == "06 - Figuras (PeriTAB)") MUC.Habilita_Destaca_button10(true, true);
                        if (s.NameLocal == "07 - Legendas de Figuras (PeriTAB)") MUC.Habilita_Destaca_button11(true, true);
                        if (s.NameLocal == "08 - Legendas de Tabelas (PeriTAB)") MUC.Habilita_Destaca_button12(true, true);
                        if (s.NameLocal == "09 - Quesitos (PeriTAB)") MUC.Habilita_Destaca_button13(true, true);
                        //if (s.NameLocal == "01 - Sem Formatação (PeriTAB)") Globals.ThisAddIn.iMyUserControl.Habilita_Destaca_button1(true, true);
                        //if (s.NameLocal == "02 - Corpo do Texto (PeriTAB)") Globals.ThisAddIn.iMyUserControl.Habilita_Destaca_button2(true, true);
                        //if (s.NameLocal == "03 - Citações (PeriTAB)") Globals.ThisAddIn.iMyUserControl.Habilita_Destaca_button3(true, true);
                        //if (s.NameLocal == "04a - Seção_1 (PeriTAB)") Globals.ThisAddIn.iMyUserControl.Habilita_Destaca_button4(true, true);
                        //if (s.NameLocal == "04b - Seção_2 (PeriTAB)") Globals.ThisAddIn.iMyUserControl.Habilita_Destaca_button5(true, true);
                        //if (s.NameLocal == "04c - Seção_3 (PeriTAB)") Globals.ThisAddIn.iMyUserControl.Habilita_Destaca_button6(true, true);
                        //if (s.NameLocal == "04d - Seção_4 (PeriTAB)") Globals.ThisAddIn.iMyUserControl.Habilita_Destaca_button7(true, true);
                        //if (s.NameLocal == "05 - Enumerações (PeriTAB)") Globals.ThisAddIn.iMyUserControl.Habilita_Destaca_button8(true, true);
                        //if (s.NameLocal == "06 - Figuras (PeriTAB)") Globals.ThisAddIn.iMyUserControl.Habilita_Destaca_button10(true, true);
                        //if (s.NameLocal == "07 - Legendas de Figuras (PeriTAB)") Globals.ThisAddIn.iMyUserControl.Habilita_Destaca_button11(true, true);
                        //if (s.NameLocal == "08 - Legendas de Tabelas (PeriTAB)") Globals.ThisAddIn.iMyUserControl.Habilita_Destaca_button12(true, true);
                        //if (s.NameLocal == "09 - Quesitos (PeriTAB)") Globals.ThisAddIn.iMyUserControl.Habilita_Destaca_button13(true, true);
                    }
                }
            }
        }
    }
}
