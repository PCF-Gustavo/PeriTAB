using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Diagnostics;

namespace PeriTAB
{
    internal class Class_AnyButtonClick_Event
    {
        public void Evento_AnyButtonClick(MyUserControl UC)
        {
            foreach (RibbonGroup g in Globals.Ribbons.Ribbon.tab.Groups) //Loop botoes do Ribbon
            {
                foreach (RibbonControl c in g.Items)
                {
                    //MessageBox.Show((c.GetType()).Name);
                    if ((c.GetType()).Name == "RibbonButtonImpl")
                    {
                        RibbonButton b = (RibbonButton)c;
                        b.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Metodo_AnyButtonClick_Ribbon);
                    }
                    if ((c.GetType()).Name == "RibbonToggleButtonImpl")
                    {
                        RibbonToggleButton tb = (RibbonToggleButton)c;
                        tb.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Metodo_AnyButtonClick_Ribbon);
                    }
                    if ((c.GetType()).Name == "RibbonCheckBoxImpl")
                    {
                        RibbonCheckBox cb = (RibbonCheckBox)c;
                        cb.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Metodo_AnyButtonClick_Ribbon);
                    }
                }
            }
            if (UC != null)
            {
                foreach (System.Windows.Forms.Button b1 in UC.Controls) //Loop botoes do Taskpane
                {
                    if ((b1.GetType()).Name == "Button")
                    {
                        b1.Click += new System.EventHandler(Metodo_AnyButtonClick_TaskPane);
                    }
                }
            }
        }

        private void Metodo_AnyButtonClick_Ribbon(object sender, RibbonControlEventArgs e)
        {
            //MessageBox.Show("AnyButtonClick_Ribbon");

            //Class_Buttons iClass_Buttons = new Class_Buttons();
            Class_RibbonControls iClass_RibbonControls = new Class_RibbonControls();

            ////Revisa a habilitação do botao "Gera PDF" do Ribbon e Sessão de token
            iClass_RibbonControls.button_gera_pdf_valorinicial();
            if (Globals.ThisAddIn.Application.ActiveDocument.Path == "") { Globals.Ribbons.Ribbon.button_gera_pdf.Enabled = false; Globals.Ribbons.Ribbon.button_gera_pdf.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon.button_gera_pdf.SuperTip = "Este documento ainda não foi salvo."; }

            //Revisa o destaque dos botoes do TaskPane
            if (Globals.Ribbons.Ribbon.toggleButton_painel_de_estilos.Checked)
            {
                if (Globals.ThisAddIn.CustomTaskPanes.Count > 0)
                {
                    Stopwatch stopWatch = new Stopwatch();
                    stopWatch.Start(); // Inicia o cronômetro

                    MyUserControl UserControl_ActiveDocument = Globals.ThisAddIn.Dicionario_Doc_e_UserControl[Globals.ThisAddIn.Application.ActiveDocument];
                    Globals.ThisAddIn.iMyUserControl.Remove_Destaque_Botoes(UserControl_ActiveDocument);

                    foreach (Microsoft.Office.Interop.Word.Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
                    {
                        // Limita o tempo de processamento a 0.2 segundos
                        if (stopWatch.Elapsed.TotalSeconds > 0.2)
                            break;

                        Microsoft.Office.Interop.Word.Style estilo = null;
                        if (p.Range.StoryType == WdStoryType.wdMainTextStory)
                        {
                            try { estilo = p.Range.get_Style(); } catch (System.Runtime.InteropServices.COMException) { }

                            if (estilo != null && UserControl_ActiveDocument.dict_estilo_e_botao.ContainsKey(estilo.NameLocal))
                            {
                                System.Windows.Forms.Button botao = UserControl_ActiveDocument.dict_estilo_e_botao[estilo.NameLocal];
                                UserControl_ActiveDocument.Habilita_Destaca(botao, true, true);
                            }
                        }
                        if (p.Range.StoryType == WdStoryType.wdFootnotesStory)
                        {
                            try { estilo = p.Range.ParagraphFormat.get_Style(); } catch (System.Runtime.InteropServices.COMException) { }
                            if (estilo != null && UserControl_ActiveDocument.dict_estilo_e_botao.ContainsKey(estilo.NameLocal))
                            {
                                System.Windows.Forms.Button botao = UserControl_ActiveDocument.dict_estilo_e_botao[estilo.NameLocal];
                                UserControl_ActiveDocument.Habilita_Destaca(botao, true, true);
                            }
                        }
                    }
                }
            }
        }
        private void Metodo_AnyButtonClick_TaskPane(object sender, EventArgs e)
        {
            //System.Windows.MessageBox.Show("AnyButtonClick_TaskPane");

            //Declara instacias das classes
            MyUserControl UserControl_ActiveDocument = Globals.ThisAddIn.Dicionario_Doc_e_UserControl[Globals.ThisAddIn.Application.ActiveDocument];

            //Revisa o destaque dos botoes do TaskPane
            if (Globals.Ribbons.Ribbon.toggleButton_painel_de_estilos.Checked)
            {
                if (Globals.ThisAddIn.CustomTaskPanes.Count > 0)
                {
                    Stopwatch stopWatch = new Stopwatch();
                    stopWatch.Start(); // Inicia o cronômetro

                    Globals.ThisAddIn.iMyUserControl.Remove_Destaque_Botoes(UserControl_ActiveDocument);

                    foreach (Microsoft.Office.Interop.Word.Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
                    {
                        // Limita o tempo de processamento a 0.2 segundos
                        if (stopWatch.Elapsed.TotalSeconds > 0.2)
                            break;

                        Microsoft.Office.Interop.Word.Style estilo = null;
                        if (p.Range.StoryType == WdStoryType.wdMainTextStory)
                        {
                            try { estilo = p.Range.get_Style(); } catch (System.Runtime.InteropServices.COMException) { }

                            if (estilo != null && UserControl_ActiveDocument.dict_estilo_e_botao.ContainsKey(estilo.NameLocal))
                            {
                                System.Windows.Forms.Button botao = UserControl_ActiveDocument.dict_estilo_e_botao[estilo.NameLocal];
                                UserControl_ActiveDocument.Habilita_Destaca(botao, true, true);
                            }
                        }
                        if (p.Range.StoryType == WdStoryType.wdFootnotesStory)
                        {
                            try { estilo = p.Range.ParagraphFormat.get_Style(); } catch (System.Runtime.InteropServices.COMException) { }
                            if (estilo != null && UserControl_ActiveDocument.dict_estilo_e_botao.ContainsKey(estilo.NameLocal))
                            {
                                System.Windows.Forms.Button botao = UserControl_ActiveDocument.dict_estilo_e_botao[estilo.NameLocal];
                                UserControl_ActiveDocument.Habilita_Destaca(botao, true, true);
                            }
                        }
                    }
                }
            }
        }
    }
}
