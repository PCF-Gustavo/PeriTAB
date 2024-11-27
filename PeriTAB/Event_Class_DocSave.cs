using Microsoft.Office.Interop.Word;
using System.Threading;
using Tarefa = System.Threading.Tasks.Task;

namespace PeriTAB
{
    internal class Class_DocSave_Event
    {
        public void Evento_DocSave()
        {
            Globals.ThisAddIn.Application.DocumentBeforeSave += new ApplicationEvents4_DocumentBeforeSaveEventHandler(Metodo_DocumentBeforeSave);
        }

        public void Metodo_DocumentBeforeSave(Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            espera_salvar(); //Cria a thread  que chama o Metodo_DocumentAfterSave()
        }

        private void espera_salvar()
        {
            /*new Thread(() =>*/
            Tarefa.Run(() =>
{
    while (true)
    {
        try
        {
            while (Globals.ThisAddIn.Application.BackgroundSavingStatus > 0) // Wait until the save operation is complete. (Globals.ThisAddIn.Application will throw exceptions while the save file dialog is open)
                Thread.Sleep(1000);
            break;
        }
        catch
        {
            Thread.Sleep(1000);
        }
    }
    // If we get to here, the user either saved the document or canceled the saving process. To distinguish between the two, we check the value of document.Saved.
    while (true)
    {
        try
        {
            if (Globals.ThisAddIn.Application.ActiveDocument.Saved) Metodo_DocumentAfterSave();
            break;
        }
        catch
        {
            Thread.Sleep(1000);
        }
    }
    /*}).Start();*/
});
        }

        public void Metodo_DocumentAfterSave()
        {
            //Declara instacias das classes
            //Class_Buttons iClass_Buttons = new Class_Buttons();
            Class_RibbonControls iClass_RibbonControls = new Class_RibbonControls();

            //Revisa a habilitação do botao "Abre SISCRIM" do Ribbon
            iClass_RibbonControls.button_abre_SISCRIM_valorinicial();
            //iClass_Buttons.button_abre_SISCRIM_Default();
            if (Globals.ThisAddIn.Application.ActiveDocument.Path == "") { Globals.Ribbons.Ribbon.button_abre_SISCRIM.Enabled = false; Globals.Ribbons.Ribbon.button_abre_SISCRIM.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon.button_abre_SISCRIM.SuperTip = "Este documento ainda não foi salvo."; }

            //Revisa a habilitação do botao "Renomeia Documento" do Ribbon
            iClass_RibbonControls.button_renomeia_documento_valorinicial();
            //iClass_Buttons.button_renomeia_documento_Default();
            if (Globals.ThisAddIn.Application.ActiveDocument.Path == "") { Globals.Ribbons.Ribbon.button_renomeia_documento.Enabled = false; Globals.Ribbons.Ribbon.button_renomeia_documento.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon.button_renomeia_documento.SuperTip = "Este documento ainda não foi salvo."; }

            //Revisa a habilitação do botao "Gera PDF" do Ribbon
            iClass_RibbonControls.button_gera_pdf_valorinicial();
            //iClass_Buttons.button_gera_pdf_Default();
            if (Globals.ThisAddIn.Application.ActiveDocument.Path == "") { Globals.Ribbons.Ribbon.button_gera_pdf.Enabled = false; Globals.Ribbons.Ribbon.button_gera_pdf.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon.button_gera_pdf.SuperTip = "Este documento ainda não foi salvo."; }
        }








    }
}
