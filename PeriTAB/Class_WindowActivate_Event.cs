using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

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
            //Revisa a habilitação do botao "Renomeia Documento" do Ribbon
            Class_Buttons iClass_Buttons = new Class_Buttons(); iClass_Buttons.button_renomeia_documento_Default();
            if (Globals.ThisAddIn.Application.ActiveDocument.Path == "") { Globals.Ribbons.Ribbon1.button_renomeia_documento.Enabled = false; Globals.Ribbons.Ribbon1.button_renomeia_documento.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon1.button_renomeia_documento.SuperTip = "Este documento ainda não foi salvo."; }
            if (Globals.ThisAddIn.Application.ActiveDocument.FullName == "http") { Globals.Ribbons.Ribbon1.button_renomeia_documento.Enabled = false; Globals.Ribbons.Ribbon1.button_renomeia_documento.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon1.button_renomeia_documento.SuperTip = "Este documento não pode ser renomeado porque está salvo online."; }

            //Revisa a habilitação do CheckBox "Destacar" do Ribbon
            Class_ValueChanged_Event iClass_ValueChanged_Event = new Class_ValueChanged_Event();
            iClass_ValueChanged_Event.FieldShading();

            //Revisa a habilitação do CheckBox "Ver código" do Ribbon
            iClass_ValueChanged_Event.ShowFieldCodes();
                        
            //Revisa a habilitação do botao "Reinicia Lista" do TaskPane
            Globals.ThisAddIn.iUserControl1.Habilita_button9(true);
            if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count > 1 | Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListType == 0 | Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListValue == 1) { Globals.ThisAddIn.iUserControl1.Habilita_button9(false); }
        }
    }
}
