using Microsoft.Office.Interop.Word;
using System;
using System.Windows.Forms;

namespace PeriTAB
{
    internal class Class_DocSave_Event
    {
        private readonly Class_RibbonControls iClass_RibbonControls = new Class_RibbonControls();
        private Timer timer;
        private Document Doc_Esperando_Primeiro_Salvamento;

        public void Evento_DocSave()
        {
            Globals.ThisAddIn.Application.DocumentBeforeSave += Metodo_DocumentBeforeSave;
        }

        private void Metodo_DocumentBeforeSave(Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            if (!string.IsNullOrEmpty(Doc.Path)) return;
            Doc_Esperando_Primeiro_Salvamento = Doc;
            Espera_primeiro_salvamento();
        }

        private void Espera_primeiro_salvamento()
        {
            // Evita múltiplos timers
            timer?.Stop();
            timer?.Dispose();

            timer = new Timer();
            timer.Interval = 500; // ms
            timer.Tick += Testa_Se_Salvou;
            timer.Start();
        }

        private void Testa_Se_Salvou(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(Doc_Esperando_Primeiro_Salvamento.Path)) return;

                timer.Stop();
                timer.Dispose();
                timer = null;

                Metodo_DocumentAfterFirstSave();
                Doc_Esperando_Primeiro_Salvamento = null;
            }
            catch (System.Runtime.InteropServices.COMException) {}
        }

        private void Metodo_DocumentAfterFirstSave()
        {
            iClass_RibbonControls.Atualiza_Habilitacao(Globals.Ribbons.Ribbon.Button_renomeia_documento);
            iClass_RibbonControls.Atualiza_Habilitacao(Globals.Ribbons.Ribbon.Button_gera_pdf);
        }
    }
}
