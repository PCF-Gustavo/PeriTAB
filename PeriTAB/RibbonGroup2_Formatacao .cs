using Microsoft.Office.Tools.Ribbon;
using System.Diagnostics;
using Tarefa = System.Threading.Tasks.Task;


namespace PeriTAB
{
    public partial class Ribbon
    {

        private async void button_alinha_legenda_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;

            // Executa as tarefas em segundo plano
            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.Run("alinha_legenda");
            });

            // Após a execução das tarefas, atualiza a UI na Thread principal
            RibbonButton.Image = Properties.Resources.lupa;
            RibbonButton.Enabled = true;
        }

        private void toggleButton_painel_de_estilos_Click(object sender, RibbonControlEventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (PeriTAB.Ribbon.Variables.debugging) { stopwatch.Start(); }
            string msg_StatusBar = "";
            var botao_toggle = (Microsoft.Office.Tools.Ribbon.RibbonToggleButton)sender;
            if (botao_toggle.Checked == true)
            {
                iClass_CustomTaskPanes.Visible(true);
                if (PeriTAB.Ribbon.Variables.debugging) msg_StatusBar = "Painel de Estilos: Aberto";
            }
            if (botao_toggle.Checked == false)
            {
                iClass_CustomTaskPanes.Visible(false);
                if (PeriTAB.Ribbon.Variables.debugging) msg_StatusBar = "Painel de Estilos: Fechado";
            }

            if (PeriTAB.Ribbon.Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
            {
                stopwatch.Stop();
                msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
            }
        }
    }
}