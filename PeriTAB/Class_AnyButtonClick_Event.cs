using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Markup;

namespace PeriTAB
{
    internal class Class_AnyButtonClick_Event
    {
        public void Evento_AnyButtonClick()
        {            
            foreach (RibbonGroup g in Globals.Ribbons.Ribbon1.tab.Groups) //Loop botoes do Ribbon
            {
                foreach (RibbonControl c in g.Items)
                {
                    if ((c.GetType()).Name == "RibbonButtonImpl")
                    {
                        RibbonButton b = (RibbonButton)c;
                        b.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Metodo_AnyButtonClick_Ribbon);
                    }
                }
            }
            foreach (System.Windows.Forms.Button b1 in Globals.ThisAddIn.iUserControl1.Controls) //Loop botoes do Taskpane
            {
                if ((b1.GetType()).Name == "Button")
                {
                    b1.Click += new System.EventHandler(Metodo_AnyButtonClick_TaskPane);
                }
            }     
        }

        private void Metodo_AnyButtonClick_Ribbon(object sender, RibbonControlEventArgs e)
        {
        }
        
        private void Metodo_AnyButtonClick_TaskPane(object sender, EventArgs e)
        {
            Globals.ThisAddIn.iUserControl1.Habilita_button9(true);
            if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count > 1 | Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListType == 0 | Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListValue == 1) { Globals.ThisAddIn.iUserControl1.Habilita_button9(false); }
        }

    }
}
