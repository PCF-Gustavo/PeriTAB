using Microsoft.Office.Tools.Ribbon;

namespace PeriTAB
{
    public partial class Ribbon
    {
        private void button_teste_Click(object sender, RibbonControlEventArgs e)
        {
        }

        private void button_moeda_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("moeda_por_extenso");
        }

        private void button_inteiro_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("inteiro_por_extenso");
        }
    }
}