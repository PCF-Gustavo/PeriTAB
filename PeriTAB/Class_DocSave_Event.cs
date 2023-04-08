﻿using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

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
            new Thread(() =>
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
            }).Start();
        }

        public void Metodo_DocumentAfterSave()
        {
            //Revisa a habilitação do botao "Renomeia Documento" do Ribbon
            Class_Buttons iClass_Buttons = new Class_Buttons(); iClass_Buttons.button_renomeia_documento_Default();
            if (Globals.ThisAddIn.Application.ActiveDocument.Path == "") { Globals.Ribbons.Ribbon1.button_renomeia_documento.Enabled = false; Globals.Ribbons.Ribbon1.button_renomeia_documento.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon1.button_renomeia_documento.SuperTip = "Este documento ainda não foi salvo."; }
            if (Globals.ThisAddIn.Application.ActiveDocument.FullName == "http") { Globals.Ribbons.Ribbon1.button_renomeia_documento.Enabled = false; Globals.Ribbons.Ribbon1.button_renomeia_documento.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon1.button_renomeia_documento.SuperTip = "Este documento não pode ser renomeado porque está salvo online."; }                     
        }








    }
}
