using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PeriTAB
{
    internal class Class_SelectionChange_Event
    {
        public void Evento_SelectionChange()
        {
            Globals.ThisAddIn.Application.WindowSelectionChange += new ApplicationEvents4_WindowSelectionChangeEventHandler(Metodo_SelectionChange);
        }

        private void Metodo_SelectionChange(Selection Sel)
        {
            //Declara instacias das classes
            Class_Buttons iClass_Buttons = new Class_Buttons(); 
            
            //iClass_Buttons.button_renomeia_documento_Default();

            ////Revisa a habilitação do botao "Cola Figura" do Ribbon
            //iClass_Buttons.button_cola_imagem_Default();
            //if (!System.Windows.Clipboard.ContainsData("FileDrop")) { Globals.Ribbons.Ribbon1.button_cola_imagem.Enabled = false; Globals.Ribbons.Ribbon1.button_cola_imagem.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon1.button_cola_imagem.SuperTip = "Não há imagem no Clipboard."; }
            //if (Globals.ThisAddIn.Application.Language != MsoLanguageID.msoLanguageIDBrazilianPortuguese) { Globals.Ribbons.Ribbon1.button_cola_imagem.Enabled = false; Globals.Ribbons.Ribbon1.button_cola_imagem.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon1.button_cola_imagem.SuperTip = "Este botão apenas funciona no Word em Português Brasileiro."; }

            //Revisa a habilitação do ToggleButton "Painel de Estilos" do Ribbon
            if (Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked) Class_New_or_Open_Event.Metodo_TaskPanes_Visible(true);

            //Revisa a habilitação do CheckBox "Destacar campos" do Ribbon
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)1) { Globals.Ribbons.Ribbon1.checkBox_destaca_campos.Checked = true; }
                if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)0 | Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)2) { Globals.Ribbons.Ribbon1.checkBox_destaca_campos.Checked = false; }
            } catch (System.Runtime.InteropServices.COMException ex) { }

            //Revisa a habilitação do CheckBox "Mostrar indicadores" do Ribbon
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowBookmarks == true) { Globals.Ribbons.Ribbon1.checkBox_mostra_indicadores.Checked = true; }
                if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowBookmarks == false) { Globals.Ribbons.Ribbon1.checkBox_mostra_indicadores.Checked = false; }
            }
            catch (System.Runtime.InteropServices.COMException ex) { }


            //Revisa a habilitação do CheckBox "Ver código" do Ribbon
            if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes == true) { Globals.Ribbons.Ribbon1.checkBox_vercodigo_campos.Checked = true; }
            if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes == false) { Globals.Ribbons.Ribbon1.checkBox_vercodigo_campos.Checked = false; }

            //Revisa a habilitação do CheckBox "Atualizar antes de imprimir" do Ribbon
            if (Globals.ThisAddIn.Application.Options.UpdateFieldsAtPrint == true) { Globals.Ribbons.Ribbon1.checkBox_atualizar_antes_de_imprimir_campos.Checked = true; }
            if (Globals.ThisAddIn.Application.Options.UpdateFieldsAtPrint == false) { Globals.Ribbons.Ribbon1.checkBox_atualizar_antes_de_imprimir_campos.Checked = false; }

            //Revisa a habilitação do botao "Reinicia Lista" do TaskPane
            if (Globals.ThisAddIn.CustomTaskPanes.Count > 0)
            {
                //Globals.ThisAddIn.iMyUserControl.Habilita_button_reinicia_lista(true);
                Globals.ThisAddIn.iMyUserControl.Habilita_Destaca(Globals.ThisAddIn.iMyUserControl.MyButton("button_reinicia_lista"), true);
                if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count > 1 | Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListType == WdListType.wdListNoNumbering | Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListValue == 1) { Globals.ThisAddIn.iMyUserControl.Habilita_Destaca(Globals.ThisAddIn.iMyUserControl.MyButton("button_reinicia_lista"), false); }
            }

            //Revisa o destaque dos botoes do TaskPane
            if (Globals.ThisAddIn.CustomTaskPanes.Count > 0)
            {
                Stopwatch stopWatch = new Stopwatch(); stopWatch.Start(); //inicia cronometro
                MyUserControl MUC = Globals.ThisAddIn.Dicionario_Doc_e_UserControl[Globals.ThisAddIn.Application.ActiveDocument];
                Globals.ThisAddIn.iMyUserControl.Remove_Destaque_Botoes(MUC);
                foreach (Microsoft.Office.Interop.Word.Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
                {
                    if (stopWatch.Elapsed.TotalSeconds > 0.2) break; //limita tempo de processamento
                    if (p.Range.StoryType == WdStoryType.wdMainTextStory)
                    {
                        Microsoft.Office.Interop.Word.Style s = null;
                        try { s = p.Range.get_Style(); } catch (System.Runtime.InteropServices.COMException ex) { }
                        if (s != null)
                        {
                            if (s.NameLocal == "01 - Sem Formatação (PeriTAB)") MUC.Habilita_Destaca(MUC.MyButton("button_sem_formatacao"), true, true);
                            if (s.NameLocal == "02 - Corpo do Texto (PeriTAB)") MUC.Habilita_Destaca(MUC.MyButton("button_corpo_do_texto"), true, true);
                            if (s.NameLocal == "03 - Citações (PeriTAB)") MUC.Habilita_Destaca(MUC.MyButton("button_citacoes"), true, true);
                            if (s.NameLocal == "04a - Seção_1 (PeriTAB)") MUC.Habilita_Destaca(MUC.MyButton("button_secao_1"), true, true);
                            if (s.NameLocal == "04b - Seção_2 (PeriTAB)") MUC.Habilita_Destaca(MUC.MyButton("button_secao_2"), true, true);
                            if (s.NameLocal == "04c - Seção_3 (PeriTAB)") MUC.Habilita_Destaca(MUC.MyButton("button_secao_3"), true, true);
                            if (s.NameLocal == "04d - Seção_4 (PeriTAB)") MUC.Habilita_Destaca(MUC.MyButton("button_secao_4"), true, true);
                            if (s.NameLocal == "05 - Enumerações (PeriTAB)") MUC.Habilita_Destaca(MUC.MyButton("button_enumeracao"), true, true);
                            if (s.NameLocal == "06 - Figuras (PeriTAB)") MUC.Habilita_Destaca(MUC.MyButton("button_figuras"), true, true);
                            if (s.NameLocal == "07 - Legendas de Figuras (PeriTAB)") MUC.Habilita_Destaca(MUC.MyButton("button_legendas_de_figuras"), true, true);
                            if (s.NameLocal == "08a - Texto de Figuras (PeriTAB)") MUC.Habilita_Destaca(MUC.MyButton("button_textos_de_figuras"), true, true);
                            if (s.NameLocal == "08 - Legendas de Tabelas (PeriTAB)") MUC.Habilita_Destaca(MUC.MyButton("button_legendas_de_tabelas"), true, true);
                            if (s.NameLocal == "09 - Quesitos (PeriTAB)") MUC.Habilita_Destaca(MUC.MyButton("button_quesitos"), true, true);
                            if (s.NameLocal == "10 - Fecho (PeriTAB)") MUC.Habilita_Destaca(MUC.MyButton("button_fecho"), true, true);
                            //if (s.NameLocal == "01 - Sem Formatação (PeriTAB)") MUC.Habilita_Destaca_button1(true, true);
                            //if (s.NameLocal == "02 - Corpo do Texto (PeriTAB)") MUC.Habilita_Destaca_button2(true, true);
                            //if (s.NameLocal == "03 - Citações (PeriTAB)") MUC.Habilita_Destaca_button3(true, true);
                            //if (s.NameLocal == "04a - Seção_1 (PeriTAB)") MUC.Habilita_Destaca_button4(true, true);
                            //if (s.NameLocal == "04b - Seção_2 (PeriTAB)") MUC.Habilita_Destaca_button5(true, true);
                            //if (s.NameLocal == "04c - Seção_3 (PeriTAB)") MUC.Habilita_Destaca_button6(true, true);
                            //if (s.NameLocal == "04d - Seção_4 (PeriTAB)") MUC.Habilita_Destaca_button7(true, true);
                            //if (s.NameLocal == "05 - Enumerações (PeriTAB)") MUC.Habilita_Destaca_button8(true, true);
                            //if (s.NameLocal == "06 - Figuras (PeriTAB)") MUC.Habilita_Destaca_button10(true, true);
                            //if (s.NameLocal == "07 - Legendas de Figuras (PeriTAB)") MUC.Habilita_Destaca_button11(true, true);
                            //if (s.NameLocal == "08 - Legendas de Tabelas (PeriTAB)") MUC.Habilita_Destaca_button12(true, true);
                            //if (s.NameLocal == "09 - Quesitos (PeriTAB)") MUC.Habilita_Destaca_button13(true, true);
                            //if (s.NameLocal == "10 - Fecho (PeriTAB)") MUC.Habilita_Destaca_button14(true, true);
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
}
