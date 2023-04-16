using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
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
            Class_Buttons iClass_Buttons = new Class_Buttons(); iClass_Buttons.button_renomeia_documento_Default();

            ////Revisa a habilitação do botao "Cola Figura" do Ribbon
            //iClass_Buttons.button_cola_imagem_Default();
            //if (!System.Windows.Clipboard.ContainsData("FileDrop")) { Globals.Ribbons.Ribbon1.button_cola_imagem.Enabled = false; Globals.Ribbons.Ribbon1.button_cola_imagem.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon1.button_cola_imagem.SuperTip = "Não há imagem no Clipboard."; }
            //if (Globals.ThisAddIn.Application.Language != MsoLanguageID.msoLanguageIDBrazilianPortuguese) { Globals.Ribbons.Ribbon1.button_cola_imagem.Enabled = false; Globals.Ribbons.Ribbon1.button_cola_imagem.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon1.button_cola_imagem.SuperTip = "Este botão apenas funciona no Word em Português Brasileiro."; }

            //Revisa a habilitação do CheckBox "Destacar" do Ribbon            
            if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)1) Globals.Ribbons.Ribbon1.checkBox_destaca_campos.Checked = true;
            if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)0 | Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)2) Globals.Ribbons.Ribbon1.checkBox_destaca_campos.Checked = false;

            //Revisa a habilitação do CheckBox "Ver código" do Ribbon
            if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes == true) Globals.Ribbons.Ribbon1.checkBox_vercodigo_campos.Checked = true;
            if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes == false) Globals.Ribbons.Ribbon1.checkBox_vercodigo_campos.Checked = false;

            //Revisa a habilitação do botao "Reinicia Lista" do TaskPane
            Globals.ThisAddIn.iUserControl1.Habilita_button9(true);
            if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count>1 | Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListType == 0 | Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListValue == 1) { Globals.ThisAddIn.iUserControl1.Habilita_button9(false); }

            //Revisa o destaque dos botoes do TaskPane
            Globals.ThisAddIn.iUserControl1.Remove_Destaque_Botoes();
            //try{
                foreach (Microsoft.Office.Interop.Word.Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
                {
                    Microsoft.Office.Interop.Word.Style s = p.Range.get_Style();
                    if (s != null)
                    {
                        if (s.NameLocal == "01 - Sem Formatação (PeriTAB)") Globals.ThisAddIn.iUserControl1.Habilita_Destaca_button1(true, true);
                        if (s.NameLocal == "02 - Corpo do Texto (PeriTAB)") Globals.ThisAddIn.iUserControl1.Habilita_Destaca_button2(true, true);
                        if (s.NameLocal == "03 - Citações (PeriTAB)") Globals.ThisAddIn.iUserControl1.Habilita_Destaca_button3(true, true);
                        if (s.NameLocal == "04 - Seções (PeriTAB)" & p.Range.ListFormat.ListLevelNumber == 1) Globals.ThisAddIn.iUserControl1.Habilita_Destaca_button4(true, true);
                        if (s.NameLocal == "04 - Seções (PeriTAB)" & p.Range.ListFormat.ListLevelNumber == 2) Globals.ThisAddIn.iUserControl1.Habilita_Destaca_button5(true, true);
                        if (s.NameLocal == "04 - Seções (PeriTAB)" & p.Range.ListFormat.ListLevelNumber == 3) Globals.ThisAddIn.iUserControl1.Habilita_Destaca_button6(true, true);
                        if (s.NameLocal == "04 - Seções (PeriTAB)" & p.Range.ListFormat.ListLevelNumber == 4) Globals.ThisAddIn.iUserControl1.Habilita_Destaca_button7(true, true);
                        if (s.NameLocal == "05 - Enumerações (PeriTAB)") Globals.ThisAddIn.iUserControl1.Habilita_Destaca_button8(true, true);
                        if (s.NameLocal == "06 - Figuras (PeriTAB)") Globals.ThisAddIn.iUserControl1.Habilita_Destaca_button10(true, true);
                        if (s.NameLocal == "07 - Legendas de Figuras (PeriTAB)") Globals.ThisAddIn.iUserControl1.Habilita_Destaca_button11(true, true);
                        if (s.NameLocal == "08 - Legendas de Tabelas (PeriTAB)") Globals.ThisAddIn.iUserControl1.Habilita_Destaca_button12(true, true);
                        if (s.NameLocal == "09 - Quesitos (PeriTAB)") Globals.ThisAddIn.iUserControl1.Habilita_Destaca_button13(true, true);
                    }
                }
            //} catch { }
        }
    }
}
