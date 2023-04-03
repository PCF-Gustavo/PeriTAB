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
            Globals.ThisAddIn.Application.Run("estilo1_corpo_PeriTAB");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.Run("estilo2_secoes_1_PeriTAB");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.Run("estilo2_secoes_2_PeriTAB");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.Run("estilo2_secoes_3_PeriTAB");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.Run("estilo2_secoes_4_PeriTAB");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.Run("estilo3_citacoes_PeriTAB");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.Run("estilo4_enumeracoes_CONT_PeriTAB");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.Run("estilo4_enumeracoes_NOVO_PeriTAB");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.Run("estilo5_figuras_PeriTAB");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.Run("estilo6_legenda-fig_PeriTAB");
        }

        private void button11_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.Run("estilo7_legenda-tab_PeriTAB");
        }

        private void button12_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.Run("estilo8_quesitos_PeriTAB");
        }

        private void button13_Click(object sender, EventArgs e)
        {

        }

        private void button14_Click(object sender, EventArgs e)
        {

        }

        public void Metodo_button1(bool b)
        {
            button1.Enabled = b;
        }


    }
}
