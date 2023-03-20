using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PeriTAB
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("confere_numeracao_legendas");
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("alinha_legenda");
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("renomeia_documento");
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("alterna_visualizacao_campos");
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("alterna_destaque_campos");
        }

        private void button6_Click_1(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("atualizar_todos_campos");
        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("moeda_por_extenso");
        }

        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("inteiro_por_extenso");
        }

        private void button9_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("DelUnusedStyles");
        }

        private void button10_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("DelUnusedStyles");
        }
    }
}
