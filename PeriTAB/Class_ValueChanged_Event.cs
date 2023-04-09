using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace PeriTAB
{
    public class Class_ValueChanged_Event
    {

        public void FieldShading()
        {
            new Thread(() =>
            {
                WdFieldShading var = Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading;
                while (true)
                {
                    Thread.Sleep(1000);

                    if (var != Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading) break;
                }
                Metodo_FieldShading();
            }).Start();
        }
        private void Metodo_FieldShading()
        {
            if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)1) Globals.Ribbons.Ribbon1.checkBox_destaca_campos.Checked = true;
            if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)0 | Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)2) Globals.Ribbons.Ribbon1.checkBox_destaca_campos.Checked = false;   
            FieldShading();
        }


        public void ShowFieldCodes()
        {
            new Thread(() =>
            {
                bool b = Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes;
                while (true)
                {
                    Thread.Sleep(1000);
                    if (b != Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes) break;
                }
                Metodo_ShowFieldCodesChanged();                
            }).Start();
        }
        private void Metodo_ShowFieldCodesChanged()
        {
            if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes == true) Globals.Ribbons.Ribbon1.checkBox_vercodigo_campos.Checked = true;
            if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes == false) Globals.Ribbons.Ribbon1.checkBox_vercodigo_campos.Checked = false;
            ShowFieldCodes();
        }

    }
}
