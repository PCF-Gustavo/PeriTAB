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
            //Revisa a habilitação do botao "Reinicia Lista" do TaskPane
            Globals.ThisAddIn.iUserControl1.Habilita_button9(true);
            if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count>1 | Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListType == 0 | Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListValue == 1) { Globals.ThisAddIn.iUserControl1.Habilita_button9(false); }
            
            //Revisa o destaque dos botoes do TaskPane
            Globals.ThisAddIn.iUserControl1.Remove_Destaque_Botoes();
            foreach (Microsoft.Office.Interop.Word.Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs)
            {
                try { Microsoft.Office.Interop.Word.Style s = p.Range.get_Style();     
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
                } catch { }
            }

        }        

    }
}
