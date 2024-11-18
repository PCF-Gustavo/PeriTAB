using Microsoft.Office.Tools.Ribbon;


namespace PeriTAB
{
    public partial class Ribbon
    {
        private void button_limpa_estilos_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("limpa_estilos");
        }
    }
}