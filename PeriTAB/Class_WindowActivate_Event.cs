using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.IO;

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
            Class_Buttons iClass_Buttons = new Class_Buttons();
            //Class_ValueChanged_Event iClass_ValueChanged_Event = new Class_ValueChanged_Event();           

            //Revisa a habilitação do CheckBox "Destacar campos" do Ribbon            
            //iClass_ValueChanged_Event.FieldShading(); *** Não está funcinando bem. A call "var = Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading;" impede a inserção de formas
            //if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)1) { Globals.Ribbons.Ribbon1.checkBox_destaca_campos.Checked = true; }
            //if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)0 | Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)2) { Globals.Ribbons.Ribbon1.checkBox_destaca_campos.Checked = false; }

            //Revisa a habilitação do CheckBox "Destacar campos" do Ribbon
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)1) { Globals.Ribbons.Ribbon1.checkBox_destaca_campos.Checked = true; }
                if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)0 | Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)2) { Globals.Ribbons.Ribbon1.checkBox_destaca_campos.Checked = false; }
            }
            catch (System.Runtime.InteropServices.COMException) { }

            //Revisa a habilitação do CheckBox "Mostrar indicadores" do Ribbon
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowBookmarks == true) { Globals.Ribbons.Ribbon1.checkBox_mostra_indicadores.Checked = true; }
                if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowBookmarks == false) { Globals.Ribbons.Ribbon1.checkBox_mostra_indicadores.Checked = false; }
            }
            catch (System.Runtime.InteropServices.COMException) { }

            //Revisa a habilitação do CheckBox "Ver código" do Ribbon
            //iClass_ValueChanged_Event.ShowFieldCodes();
            if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes == true) { Globals.Ribbons.Ribbon1.checkBox_vercodigo_campos.Checked = true; }
            if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes == false) { Globals.Ribbons.Ribbon1.checkBox_vercodigo_campos.Checked = false; }

            //Revisa a habilitação do CheckBox "Atualizar antes de imprimir" do Ribbon
            if (Globals.ThisAddIn.Application.Options.UpdateFieldsAtPrint == true) { Globals.Ribbons.Ribbon1.checkBox_atualizar_antes_de_imprimir_campos.Checked = true; }
            if (Globals.ThisAddIn.Application.Options.UpdateFieldsAtPrint == false) { Globals.Ribbons.Ribbon1.checkBox_atualizar_antes_de_imprimir_campos.Checked = false; }

            ////Revisa a habilitação do botao "Cola Figura" do Ribbon
            //iClass_Buttons.button_cola_imagem_Default();
            //if (!System.Windows.Clipboard.ContainsData("FileDrop")) { Globals.Ribbons.Ribbon1.button_cola_imagem.Enabled = false; Globals.Ribbons.Ribbon1.button_cola_imagem.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon1.button_cola_imagem.SuperTip = "Não há imagem no Clipboard."; }
            //if (Globals.ThisAddIn.Application.Language != MsoLanguageID.msoLanguageIDBrazilianPortuguese) { Globals.Ribbons.Ribbon1.button_cola_imagem.Enabled = false; Globals.Ribbons.Ribbon1.button_cola_imagem.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon1.button_cola_imagem.SuperTip = "Este botão apenas funciona no Word em Português Brasileiro."; }

            //Revisa a habilitação do botao "Renomeia Documento" do Ribbon
            iClass_Buttons.button_renomeia_documento_Default();
            if (Globals.ThisAddIn.Application.ActiveDocument.Path == "") { Globals.Ribbons.Ribbon1.button_renomeia_documento.Enabled = false; Globals.Ribbons.Ribbon1.button_renomeia_documento.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon1.button_renomeia_documento.SuperTip = "Este documento ainda não foi salvo."; }
            //else if ((Globals.ThisAddIn.Application.ActiveDocument.Path).Substring(0, 4) == "http") { Globals.Ribbons.Ribbon1.button_renomeia_documento.Enabled = false; Globals.Ribbons.Ribbon1.button_renomeia_documento.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon1.button_renomeia_documento.SuperTip = "Este documento não pode ser renomeado porque está salvo online."; }

            //Revisa a habilitação do botao "Gera PDF" do Ribbon
            iClass_Buttons.button_gera_pdf_Default();
            if (Globals.ThisAddIn.Application.ActiveDocument.Path == "") { Globals.Ribbons.Ribbon1.button_gera_pdf.Enabled = false; Globals.Ribbons.Ribbon1.button_gera_pdf.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon1.button_gera_pdf.SuperTip = "Este documento ainda não foi salvo."; }

            ////Revisa a habilitação do botao "Abre SISCRIM" do Ribbon
            //iClass_Buttons.button_abre_SISCRIM_Default();
            //string localpath = Globals.Ribbons.Ribbon1.GetLocalPath(Globals.ThisAddIn.Application.ActiveDocument.FullName);
            //if (File.Exists(localpath.Substring(0, localpath.LastIndexOf(".")) + ".pdf") | File.Exists(localpath.Substring(0, localpath.LastIndexOf(".")) + "_assinado.pdf"))
            //{
            //    Globals.Ribbons.Ribbon1.button_abre_SISCRIM.Enabled = true; Globals.Ribbons.Ribbon1.button_abre_SISCRIM.ScreenTip = ""; Globals.Ribbons.Ribbon1.button_abre_SISCRIM.SuperTip = "Abre SISCRIM na página do Laudo.";
            //}
            
            //if (Globals.ThisAddIn.Application.ActiveDocument.Path == "") 
            //{ 
            //    Globals.Ribbons.Ribbon1.button_abre_SISCRIM.Enabled = false; Globals.Ribbons.Ribbon1.button_abre_SISCRIM.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon1.button_abre_SISCRIM.SuperTip = "O PDF do laudo ainda não foi gerado."; }

            //Revisa a habilitação do botao "Reinicia Lista" do TaskPane

            if (Globals.ThisAddIn.CustomTaskPanes.Count > 0)
            {
                //Globals.ThisAddIn.iMyUserControl.Habilita_button_reinicia_lista(true);
                Globals.ThisAddIn.iMyUserControl.Habilita_Destaca(Globals.ThisAddIn.iMyUserControl.MyButton("button_reinicia_lista"), true);
                if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count > 1 | Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListType == WdListType.wdListNoNumbering | Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListValue == 1) { Globals.ThisAddIn.iMyUserControl.Habilita_Destaca(Globals.ThisAddIn.iMyUserControl.MyButton("button_reinicia_lista"), false); }
            }

            //if (Class_New_or_Open_Event.Dicionario_Doc_e_TaskPane.ContainsKey(Doc) == false) 
            //{
            //    Class_New_or_Open_Event iClass_New_or_Open_Event = new Class_New_or_Open_Event(); iClass_New_or_Open_Event.Metodo_New_or_Open(null);




            //}

        }
    }
}
