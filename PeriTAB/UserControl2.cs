using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PeriTAB
{
    public partial class UserControl1 : UserControl
    {

        public UserControl1()
        {
            InitializeComponent();            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "01 - Sem Formatação (PeriTAB)", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.Run("estilo1_Sem_Formatacao_PeriTAB");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "02 - Corpo do Texto (PeriTAB)", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.Run("estilo2_Corpo_do_Texto_PeriTAB");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "03 - Citações (PeriTAB)", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.Run("estilo3_Citacoes_PeriTAB");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "04 - Seções (PeriTAB)", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.Run("estilo4_Secoes_1_PeriTAB");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "04 - Seções (PeriTAB)", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.Run("estilo4_Secoes_2_PeriTAB");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "04 - Seções (PeriTAB)", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.Run("estilo4_Secoes_3_PeriTAB");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "04 - Seções (PeriTAB)", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.Run("estilo4_Secoes_4_PeriTAB");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "05 - Enumerações (PeriTAB)", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.Run("estilo5_Enumeracoes_PeriTAB");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.Run("reinicia_lista");
            if (Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListValue == 1) { Habilita_button9(false); }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "06 - Figuras (PeriTAB)", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.Run("estilo6_Figuras_PeriTAB");
        }

        private void button11_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "07 - Legendas de Figuras (PeriTAB)", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.Run("estilo7_Legend_Figuras_PeriTAB");            
        }

        private void button12_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "08 - Legendas de Tabelas (PeriTAB)", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.Run("estilo8_Legend_Tabelas_PeriTAB");            
        }

        private void button13_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "09 - Quesitos (PeriTAB)", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.Run("estilo9_Quesitos_PeriTAB");            
        }

        private void button_DockRight_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.TaskPane1.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
            Globals.ThisAddIn.TaskPane1.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            Globals.ThisAddIn.TaskPane1.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
            this.Size = new System.Drawing.Size(100, 900);
            //this.Width = 120;
            this.button_DockRight.Visible = false;
            this.button1.Location = new System.Drawing.Point(5, 5);
            this.button2.Location = new System.Drawing.Point(5, 50);
            this.button3.Location = new System.Drawing.Point(5, 95);
            this.button4.Location = new System.Drawing.Point(5, 140);
            this.button5.Location = new System.Drawing.Point(5, 185);
            this.button6.Location = new System.Drawing.Point(5, 230);
            this.button7.Location = new System.Drawing.Point(5, 275);
            this.button8.Location = new System.Drawing.Point(5, 320);
            this.button9.Location = new System.Drawing.Point(5, 365);
            this.button10.Location = new System.Drawing.Point(5, 410);
            this.button11.Location = new System.Drawing.Point(5, 455);
            this.button12.Location = new System.Drawing.Point(5, 500);
            this.button13.Location = new System.Drawing.Point(5, 545);
            this.button_DockBottom.Location = new System.Drawing.Point(55, 700);
            this.button_DockBottom.Visible = true;
            Globals.ThisAddIn.TaskPane1.Width = 120;
            this.Size = new System.Drawing.Size(100, 900);
        }

        private void button_DockBottom_Click(object sender, EventArgs e)
        {

        }

        public void Habilita_Destaca_button1(bool habilita, bool destaca = false)
        {
            button1.Enabled = habilita;
            if (destaca){button1.BackColor = SystemColors.Highlight;button1.ForeColor = SystemColors.HighlightText;}
        }
        public void Habilita_Destaca_button2(bool habilita, bool destaca = false)
        {
            button2.Enabled = habilita;
            if (destaca) { button2.BackColor = SystemColors.Highlight; button2.ForeColor = SystemColors.HighlightText; }
        }
        public void Habilita_Destaca_button3(bool habilita, bool destaca = false)
        {
            button3.Enabled = habilita;
            if (destaca) { button3.BackColor = SystemColors.Highlight; button3.ForeColor = SystemColors.HighlightText; }
        }
        public void Habilita_Destaca_button4(bool habilita, bool destaca = false)
        {
            button4.Enabled = habilita;
            if (destaca) { button4.BackColor = SystemColors.Highlight; button4.ForeColor = SystemColors.HighlightText; }
        }
        public void Habilita_Destaca_button5(bool habilita, bool destaca = false)
        {
            button5.Enabled = habilita;
            if (destaca) { button5.BackColor = SystemColors.Highlight; button5.ForeColor = SystemColors.HighlightText; }
        }
        public void Habilita_Destaca_button6(bool habilita, bool destaca = false)
        {
            button6.Enabled = habilita;
            if (destaca) { button6.BackColor = SystemColors.Highlight; button6.ForeColor = SystemColors.HighlightText; }
        }
        public void Habilita_Destaca_button7(bool habilita, bool destaca = false)
        {
            button7.Enabled = habilita;
            if (destaca) { button7.BackColor = SystemColors.Highlight; button7.ForeColor = SystemColors.HighlightText; }
        }
        public void Habilita_Destaca_button8(bool habilita, bool destaca = false)
        {
            button8.Enabled = habilita;
            if (destaca) { button8.BackColor = SystemColors.Highlight; button8.ForeColor = SystemColors.HighlightText; }
        }
        public void Habilita_button9(bool habilita)
        {
            button9.Enabled = habilita;            
        }
        public void Habilita_Destaca_button10(bool habilita, bool destaca = false)
        {
            button10.Enabled = habilita;
            if (destaca) { button10.BackColor = SystemColors.Highlight; button10.ForeColor = SystemColors.HighlightText; }
        }
        public void Habilita_Destaca_button11(bool habilita, bool destaca = false)
        {
            button11.Enabled = habilita;
            if (destaca) { button11.BackColor = SystemColors.Highlight; button11.ForeColor = SystemColors.HighlightText; }
        }
        public void Habilita_Destaca_button12(bool habilita, bool destaca = false)
        {
            button12.Enabled = habilita;
            if (destaca) { button12.BackColor = SystemColors.Highlight; button12.ForeColor = SystemColors.HighlightText; }
        }
        public void Habilita_Destaca_button13(bool habilita, bool destaca = false)
        {
            button13.Enabled = habilita;
            if (destaca) { button13.BackColor = SystemColors.Highlight; button13.ForeColor = SystemColors.HighlightText; }
        }

        internal void Remove_Destaque_Botoes()
        {
            foreach (Button b in Globals.ThisAddIn.iUserControl1.Controls)
            {
                if ((b.GetType()).Name == "Button")
                {
                    b.BackColor = SystemColors.Control;
                    b.ForeColor = SystemColors.ControlText;
                }

            }
        }


    }
}
