using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Tarefa = System.Threading.Tasks.Task;

namespace PeriTAB
{
    public partial class Ribbon
    {
        private void button_teste_Click(object sender, RibbonControlEventArgs e)
        {
        }

        private async void button_moeda_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;

            // Executa as tarefas em segundo plano
            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                Globals.ThisAddIn.Application.Run("moeda_por_extenso");
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            // Após a execução das tarefas, atualiza a UI na Thread principal
            RibbonButton.Image = Properties.Resources.dinheiro;
            RibbonButton.Enabled = true;
        }

        private async void button_inteiro_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;

            // Executa as tarefas em segundo plano
            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                Globals.ThisAddIn.Application.Run("inteiro_por_extenso");
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            // Após a execução das tarefas, atualiza a UI na Thread principal
            RibbonButton.Image = Properties.Resources.numero;
            RibbonButton.Enabled = true;
        }

        // Função para obter a próxima seção
        static Section GetNextSection(Section section)
        {
            // Percorre as seções do documento
            for (int i = 1; i < Globals.ThisAddIn.Application.ActiveDocument.Sections.Count; i++)
            {
                // Se encontramos a seção corrente, verificamos se existe uma próxima seção
                if (Globals.ThisAddIn.Application.ActiveDocument.Sections[i].Range.Start == section.Range.Start)
                {
                    // Verifica se não é a última seção
                    if (i + 1 <= Globals.ThisAddIn.Application.ActiveDocument.Sections.Count)
                    {
                        return Globals.ThisAddIn.Application.ActiveDocument.Sections[i + 1]; // Retorna a próxima seção
                    }
                    break; // Se for a última seção, sai do loop
                }
            }

            // Se não encontrar uma próxima seção, retorna null
            return null;
        }
    }
}