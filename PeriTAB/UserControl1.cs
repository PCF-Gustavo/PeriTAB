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
using static System.Net.Mime.MediaTypeNames;

namespace PeriTAB
{
    public partial class UserControl1 : UserControl
    {

        public UserControl1()
        {
            InitializeComponent();            
        }
        private void button_sem_formatacao_Click(object sender, EventArgs e)
        {
            string estilo_nome = "01 - Sem Formatação (PeriTAB)";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs){ p.Range.set_Style((object)estilo_nome); }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void button_corpo_do_texto_Click(object sender, EventArgs e)
        {
            string estilo_nome = "02 - Corpo do Texto (PeriTAB)";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        private void button_citacoes_Click(object sender, EventArgs e)
        {
            string estilo_nome = "03 - Citações (PeriTAB)";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void button_secao_1_Click(object sender, EventArgs e)
        {
            string estilo_nome = "04 - Seções (PeriTAB)";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); p.Range.SetListLevel(1); }
            Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "04 - Seções (PeriTAB)", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.Run("estilo4_Secoes_2_PeriTAB");

            Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "04 - Seções (PeriTAB)", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.Run("estilo4_Secoes_3_PeriTAB");

            Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "04 - Seções (PeriTAB)", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.Run("estilo4_Secoes_4_PeriTAB");

            Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco
        }

        private void button_enumeracao_Click(object sender, EventArgs e)
        {
            string estilo_nome = "estilo5_Enumeracoes_PeriTAB";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void button_reinicia_lista_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.Run("reinicia_lista");
            if (Globals.ThisAddIn.Application.Selection.Range.ListFormat.ListValue == 1) { Habilita_button9(false); }
        }

        private void button_figuras_Click(object sender, EventArgs e)
        {
            string estilo_nome = "06 - Figuras (PeriTAB)";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); }
            Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        private void button_legendas_de_figuras_Click(object sender, EventArgs e)
        {
            string estilo_nome = "07 - Legendas de Figuras (PeriTAB)";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        private void button_legendas_de_tabelas_Click(object sender, EventArgs e)
        {
            string estilo_nome = "08 - Legendas de Tabelas (PeriTAB)";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); }
            Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }        

        private void button_quesitos_Click(object sender, EventArgs e)
        {
            string estilo_nome = "09 - Quesitos (PeriTAB)";
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome, WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (Paragraph p in Globals.ThisAddIn.Application.Selection.Paragraphs) { p.Range.set_Style((object)estilo_nome); }
            Range r = Globals.ThisAddIn.Application.Selection.Previous(WdUnits.wdParagraph, 1); if (r != null) { if (r.Text == ((char)13).ToString()) { r.Delete(); } } //Deleta parágrafo anterior em branco
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }        
        
        private void button_DockRight_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.TaskPane1.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
            Globals.ThisAddIn.TaskPane1.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            Globals.ThisAddIn.TaskPane1.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
            //Globals.ThisAddIn.TaskPane1
            this.Size = new System.Drawing.Size(100, 900);
            //this.Width = 120;
            this.button_DockRight.Visible = false;
            this.button_sem_formatacao.Location = new System.Drawing.Point(5, 5);
            this.button_corpo_do_texto.Location = new System.Drawing.Point(5, 50);
            this.button_citacoes.Location = new System.Drawing.Point(5, 95);
            this.button_secao_1.Location = new System.Drawing.Point(5, 140);
            this.button5.Location = new System.Drawing.Point(5, 185);
            this.button6.Location = new System.Drawing.Point(5, 230);
            this.button7.Location = new System.Drawing.Point(5, 275);
            this.button_enumeracao.Location = new System.Drawing.Point(5, 320);
            this.button_reinicia_lista.Size = new System.Drawing.Size(90, 24);
            this.button_reinicia_lista.Location = new System.Drawing.Point(5, 359);
            this.button_figuras.Location = new System.Drawing.Point(5, 388);
            this.button_legendas_de_figuras.Location = new System.Drawing.Point(5, 433);
            this.button_legendas_de_tabelas.Location = new System.Drawing.Point(5, 478);
            this.button_quesitos.Location = new System.Drawing.Point(5, 523);
            this.button_DockBottom.Location = new System.Drawing.Point(30, 568);
            this.button_DockBottom.Visible = true;
            Globals.ThisAddIn.TaskPane1.Width = 120;
            this.Size = new System.Drawing.Size(100, 900);
        }

        private void button_DockBottom_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.TaskPane1.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
            Globals.ThisAddIn.TaskPane1.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
            Globals.ThisAddIn.TaskPane1.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
            //Globals.ThisAddIn.TaskPane1
            this.Size = new System.Drawing.Size(1400, 100);
            //this.Width = 120;
            this.button_DockBottom.Visible = false;
            this.button_DockRight.Visible = true;
            this.button_DockRight.Location = new System.Drawing.Point(1204, 5);
            this.button_sem_formatacao.Location = new System.Drawing.Point(5, 5);
            this.button_corpo_do_texto.Location = new System.Drawing.Point(100, 5);
            this.button_citacoes.Location = new System.Drawing.Point(195, 5);
            this.button_secao_1.Location = new System.Drawing.Point(290, 5);
            this.button5.Location = new System.Drawing.Point(385, 5);
            this.button6.Location = new System.Drawing.Point(480, 5);
            this.button7.Location = new System.Drawing.Point(575, 5);
            this.button_enumeracao.Location = new System.Drawing.Point(670, 5);
            this.button_reinicia_lista.Size = new System.Drawing.Size(60, 40);
            this.button_reinicia_lista.Location = new System.Drawing.Point(759, 5);
            this.button_figuras.Location = new System.Drawing.Point(824, 5);
            this.button_legendas_de_figuras.Location = new System.Drawing.Point(919, 5);
            this.button_legendas_de_tabelas.Location = new System.Drawing.Point(1014, 5);
            this.button_quesitos.Location = new System.Drawing.Point(1109, 5);                  
            this.Size = new System.Drawing.Size(1400, 100);
            Globals.ThisAddIn.TaskPane1.Height = 90;
        }

        public void Habilita_Destaca_button1(bool habilita, bool destaca = false)
        {
            button_sem_formatacao.Enabled = habilita;
            if (destaca){button_sem_formatacao.BackColor = SystemColors.Highlight;button_sem_formatacao.ForeColor = SystemColors.HighlightText;}
        }
        public void Habilita_Destaca_button2(bool habilita, bool destaca = false)
        {
            button_corpo_do_texto.Enabled = habilita;
            if (destaca) { button_corpo_do_texto.BackColor = SystemColors.Highlight; button_corpo_do_texto.ForeColor = SystemColors.HighlightText; }
        }
        public void Habilita_Destaca_button3(bool habilita, bool destaca = false)
        {
            button_citacoes.Enabled = habilita;
            if (destaca) { button_citacoes.BackColor = SystemColors.Highlight; button_citacoes.ForeColor = SystemColors.HighlightText; }
        }
        public void Habilita_Destaca_button4(bool habilita, bool destaca = false)
        {
            button_secao_1.Enabled = habilita;
            if (destaca) { button_secao_1.BackColor = SystemColors.Highlight; button_secao_1.ForeColor = SystemColors.HighlightText; }
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
            button_enumeracao.Enabled = habilita;
            if (destaca) { button_enumeracao.BackColor = SystemColors.Highlight; button_enumeracao.ForeColor = SystemColors.HighlightText; }
        }
        public void Habilita_button9(bool habilita)
        {
            button_reinicia_lista.Enabled = habilita;            
        }
        public void Habilita_Destaca_button10(bool habilita, bool destaca = false)
        {
            button_figuras.Enabled = habilita;
            if (destaca) { button_figuras.BackColor = SystemColors.Highlight; button_figuras.ForeColor = SystemColors.HighlightText; }
        }
        public void Habilita_Destaca_button11(bool habilita, bool destaca = false)
        {
            button_legendas_de_figuras.Enabled = habilita;
            if (destaca) { button_legendas_de_figuras.BackColor = SystemColors.Highlight; button_legendas_de_figuras.ForeColor = SystemColors.HighlightText; }
        }
        public void Habilita_Destaca_button12(bool habilita, bool destaca = false)
        {
            button_legendas_de_tabelas.Enabled = habilita;
            if (destaca) { button_legendas_de_tabelas.BackColor = SystemColors.Highlight; button_legendas_de_tabelas.ForeColor = SystemColors.HighlightText; }
        }
        public void Habilita_Destaca_button13(bool habilita, bool destaca = false)
        {
            button_quesitos.Enabled = habilita;
            if (destaca) { button_quesitos.BackColor = SystemColors.Highlight; button_quesitos.ForeColor = SystemColors.HighlightText; }
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
