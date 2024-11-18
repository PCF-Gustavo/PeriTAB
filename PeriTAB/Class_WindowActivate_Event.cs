using Microsoft.Office.Interop.Word;

namespace PeriTAB
{
    internal class Class_WindowActivate_Event
    {
        public void Evento_WindowActivate()
        {
            Globals.ThisAddIn.Application.WindowActivate += new ApplicationEvents4_WindowActivateEventHandler(Metodo_WindowActivate);
        }
        private void Metodo_WindowActivate(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn)
        {
            //MessageBox.Show("Winact");
            //Declara instacias das classes
            //Class_Buttons iClass_Buttons = new Class_Buttons();
            Class_Controls iClass_Controls = new Class_Controls();
            //Class_ValueChanged_Event iClass_ValueChanged_Event = new Class_ValueChanged_Event();           

            //Revisa a habilitação do CheckBox "Destacar campos" do Ribbon            
            //iClass_ValueChanged_Event.FieldShading(); *** Não está funcinando bem. A call "var = Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading;" impede a inserção de formas
            //if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)1) { Globals.Ribbons.Ribbon.checkBox_destaca_campos.Checked = true; }
            //if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)0 | Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)2) { Globals.Ribbons.Ribbon.checkBox_destaca_campos.Checked = false; }

            //Revisa a habilitação do CheckBox "Destacar campos" do Ribbon
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)1) { Globals.Ribbons.Ribbon.checkBox_destaca_campos.Checked = true; }
                if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)0 | Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)2) { Globals.Ribbons.Ribbon.checkBox_destaca_campos.Checked = false; }
            }
            catch (System.Runtime.InteropServices.COMException) { }

            //Revisa a habilitação do CheckBox "Mostrar indicadores" do Ribbon
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowBookmarks == true) { Globals.Ribbons.Ribbon.checkBox_mostra_indicadores.Checked = true; }
                if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowBookmarks == false) { Globals.Ribbons.Ribbon.checkBox_mostra_indicadores.Checked = false; }
            }
            catch (System.Runtime.InteropServices.COMException) { }

            //Revisa a habilitação do CheckBox "Ver código" do Ribbon
            //iClass_ValueChanged_Event.ShowFieldCodes();
            if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes == true) { Globals.Ribbons.Ribbon.checkBox_vercodigo_campos.Checked = true; }
            if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes == false) { Globals.Ribbons.Ribbon.checkBox_vercodigo_campos.Checked = false; }

            //Revisa a habilitação do CheckBox "Atualizar antes de imprimir" do Ribbon
            if (Globals.ThisAddIn.Application.Options.UpdateFieldsAtPrint == true) { Globals.Ribbons.Ribbon.checkBox_atualizar_antes_de_imprimir_campos.Checked = true; }
            if (Globals.ThisAddIn.Application.Options.UpdateFieldsAtPrint == false) { Globals.Ribbons.Ribbon.checkBox_atualizar_antes_de_imprimir_campos.Checked = false; }

            //Revisa a habilitação do botao "Abre SISCRIM" do Ribbon
            iClass_Controls.button_abre_SISCRIM_valorinicial();
            //iClass_Buttons.button_abre_SISCRIM_Default();
            if (Globals.ThisAddIn.Application.ActiveDocument.Path == "") { Globals.Ribbons.Ribbon.button_abre_SISCRIM.Enabled = false; Globals.Ribbons.Ribbon.button_abre_SISCRIM.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon.button_abre_SISCRIM.SuperTip = "Este documento ainda não foi salvo."; }

            //Revisa a habilitação do botao "Renomeia Documento" do Ribbon
            iClass_Controls.button_renomeia_documento_valorinicial();
            //iClass_Buttons.button_renomeia_documento_Default();
            if (Globals.ThisAddIn.Application.ActiveDocument.Path == "") { Globals.Ribbons.Ribbon.button_renomeia_documento.Enabled = false; Globals.Ribbons.Ribbon.button_renomeia_documento.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon.button_renomeia_documento.SuperTip = "Este documento ainda não foi salvo."; }

            //Revisa a habilitação do botao "Gera PDF" do Ribbon
            iClass_Controls.button_gera_pdf_valorinicial();
            //iClass_Buttons.button_gera_pdf_Default();
            if (Globals.ThisAddIn.Application.ActiveDocument.Path == "") { Globals.Ribbons.Ribbon.button_gera_pdf.Enabled = false; Globals.Ribbons.Ribbon.button_gera_pdf.Image = Properties.Resources.icone_pdf; Globals.Ribbons.Ribbon.button_gera_pdf.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon.button_gera_pdf.SuperTip = "Este documento ainda não foi salvo."; }

            ////Revisa a habilitação do botao "Reinicia Lista" do TaskPane
            //if (Globals.ThisAddIn.CustomTaskPanes.Count > 0)
            //{
            //    //Globals.ThisAddIn.iMyUserControl.Habilita_button_reinicia_lista(true);
            //    Globals.ThisAddIn.iMyUserControl.Habilita_Destaca(Globals.ThisAddIn.iMyUserControl.MyButton("button_reinicia_lista"), true);
            //    if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count > 1 | Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListType == WdListType.wdListNoNumbering | Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListValue == 1) { Globals.ThisAddIn.iMyUserControl.Habilita_Destaca(Globals.ThisAddIn.iMyUserControl.MyButton("button_reinicia_lista"), false); }
            //}

            //Revisa a habilitação do botao "Reinicia Lista" do TaskPane
            //if (Globals.ThisAddIn.CustomTaskPanes.Count > 0)
            //{

            //if (Globals.ThisAddIn.Dicionario_Doc_e_UserControl.ContainsKey(Globals.ThisAddIn.Application.ActiveDocument))
            //{
            //    MyUserControl MUC = Globals.ThisAddIn.Dicionario_Doc_e_UserControl[Globals.ThisAddIn.Application.ActiveDocument];
            //    MUC.Habilita_Destaca(MUC.MyButton("button_reinicia_lista"), true);
            //    if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count > 1 | Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListType == WdListType.wdListNoNumbering | Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListValue == 1) { MUC.Habilita_Destaca(MUC.MyButton("button_reinicia_lista"), false); }
            //    else
            //    {
            //        Microsoft.Office.Interop.Word.Style s = null;
            //        try { s = Globals.ThisAddIn.Application.Selection.Range.get_Style(); } catch (System.Runtime.InteropServices.COMException) { }
            //        if (s != null)
            //        {
            //            if (!(s.NameLocal == "05 - Enumerações (PeriTAB)")) MUC.Habilita_Destaca(MUC.MyButton("button_reinicia_lista"), false);
            //        }
            //    }
            //}


            //}

            //if (Class_New_or_Open_Event.Dicionario_Doc_e_TaskPane.ContainsKey(Doc) == false) 
            //{
            //    Class_New_or_Open_Event iClass_New_or_Open_Event = new Class_New_or_Open_Event(); iClass_New_or_Open_Event.Metodo_New_or_Open(null);




            //}

        }
    }
}
