using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using Tarefa = System.Threading.Tasks.Task;


namespace PeriTAB
{
    public partial class Ribbon
    {
        private async void button_inserir_sumario_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            menu_inserir_campos.Image = Properties.Resources.load_icon_png_7969;
            menu_inserir_campos.Enabled = false;

            bool success = true;
            string msg_StatusBar = ((RibbonButton)sender).Label + ": ";

            // Executa as tarefas em segundo plano
            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                Globals.ThisAddIn.Application.OrganizerCopy(Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "Sumário 1", WdOrganizerObject.wdOrganizerObjectStyles);
                Globals.ThisAddIn.Application.OrganizerCopy(Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "Sumário 2", WdOrganizerObject.wdOrganizerObjectStyles);
                Globals.ThisAddIn.Application.OrganizerCopy(Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "Sumário 3", WdOrganizerObject.wdOrganizerObjectStyles);
                Globals.ThisAddIn.Application.OrganizerCopy(Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "Sumário 4", WdOrganizerObject.wdOrganizerObjectStyles);
                Globals.ThisAddIn.Application.OrganizerCopy(Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "Sumário 5", WdOrganizerObject.wdOrganizerObjectStyles);
                Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldTOC, slash + "h " + slash + "z " + slash + "t " + quote + "05 - Seção_1 (PeriTAB);1;06 - Seção_2 (PeriTAB);2;07 - Seção_3 (PeriTAB);3;08 - Seção_4 (PeriTAB);4;09 - Seção_5 (PeriTAB);5" + quote, false);
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;

            // Após a execução das tarefas, atualiza a UI na Thread principal
            menu_inserir_campos.Image = null;
            menu_inserir_campos.Enabled = true;
        }

        private async void button_inserir_pagina_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            menu_inserir_campos.Image = Properties.Resources.load_icon_png_7969;
            menu_inserir_campos.Enabled = false;

            bool success = true;
            string msg_StatusBar = ((RibbonButton)sender).Label + ": ";

            // Executa as tarefas em segundo plano
            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldEmpty, "PAGE", false);
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;

            // Após a execução das tarefas, atualiza a UI na Thread principal
            menu_inserir_campos.Image = null;
            menu_inserir_campos.Enabled = true;
        }

        private async void button_inserir_pagina_extenso_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            menu_inserir_campos.Image = Properties.Resources.load_icon_png_7969;
            menu_inserir_campos.Enabled = false;

            bool success = true;
            string msg_StatusBar = ((RibbonButton)sender).Label + ": ";

            // Executa as tarefas em segundo plano
            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                inserir_autotexto(Globals.ThisAddIn.Application.Selection.Range, "pagina_atual_por_extenso_PeriTAB");
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;

            // Após a execução das tarefas, atualiza a UI na Thread principal
            menu_inserir_campos.Image = null;
            menu_inserir_campos.Enabled = true;
        }

        private async void button_inserir_paginas_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            menu_inserir_campos.Image = Properties.Resources.load_icon_png_7969;
            menu_inserir_campos.Enabled = false;

            bool success = true;
            string msg_StatusBar = ((RibbonButton)sender).Label + ": ";

            // Executa as tarefas em segundo plano
            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldEmpty, "NUMPAGES", false);
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;

            // Após a execução das tarefas, atualiza a UI na Thread principal
            menu_inserir_campos.Image = null;
            menu_inserir_campos.Enabled = true;
        }

        private async void button_inserir_paginas_extenso_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            menu_inserir_campos.Image = Properties.Resources.load_icon_png_7969;
            menu_inserir_campos.Enabled = false;

            bool success = true;
            string msg_StatusBar = ((RibbonButton)sender).Label + ": ";

            // Executa as tarefas em segundo plano
            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                inserir_autotexto(Globals.ThisAddIn.Application.Selection.Range, "paginas_por_extenso_PeriTAB");
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;

            // Após a execução das tarefas, atualiza a UI na Thread principal
            menu_inserir_campos.Image = null;
            menu_inserir_campos.Enabled = true;
        }

        private async void button_inserir_ano_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            menu_inserir_campos.Image = Properties.Resources.load_icon_png_7969;
            menu_inserir_campos.Enabled = false;

            bool success = true;
            string msg_StatusBar = ((RibbonButton)sender).Label + ": ";

            // Executa as tarefas em segundo plano
            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldEmpty, "DATE " + slash + "@ " + quote + "yyyy" + quote, false);
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            if (success) { msg_StatusBar += "Sucesso"; } else { msg_StatusBar += "Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;

            // Após a execução das tarefas, atualiza a UI na Thread principal
            menu_inserir_campos.Image = null;
            menu_inserir_campos.Enabled = true;
        }

        //private static Dictionary<string, string> dict_Secao_de_conclusao_e_Fim_do_preambulo = new Dictionary<string, string>()
        //{
        //     { "RESPOSTA AOS QUESITOS", "respondendo aos quesitos formulados, abaixo transcritos" }
        //    ,{ "CONCLUSÃO" , "atendendo ao abaixo transcrito" }
        //};

        private async void button_atualiza_campos_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;
            bool success = true;
            string msg_StatusBar = RibbonButton.Label + ": ";

            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                List<Field> campos = new List<Field>();

                // Coleta os campos em todas as StoryRanges
                foreach (Range storyRange in Globals.ThisAddIn.Application.ActiveDocument.StoryRanges)
                {
                    foreach (Field field in storyRange.Fields)
                    {
                        campos.Add(field);
                    }


                    // Coleta os campos dentro das caixas de texto
                    foreach (Microsoft.Office.Interop.Word.Shape shape in Globals.ThisAddIn.Application.ActiveDocument.Shapes)
                    {
                        if (shape.Type == MsoShapeType.msoTextBox)
                        {
                            foreach (Field field in shape.TextFrame.TextRange.Fields)
                            {
                                campos.Add(field);
                            }
                        }
                    }
                }

                // Atualiza os campos com barra de progresso
                for (int i = 0; i < campos.Count; i++)
                {
                    campos[i].Update();
                    Globals.ThisAddIn.Application.StatusBar = msg_StatusBar + " " + barra_de_progresso(((i + 1) * 10) / campos.Count);
                }

                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            if (success) { msg_StatusBar += barra_de_progresso(10) + " Sucesso"; } else { msg_StatusBar += " Falha"; }
            Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;

            // Após a execução das tarefas, atualiza a UI na Thread principal
            RibbonButton.Image = Properties.Resources.atualizar;
            RibbonButton.Enabled = true;
        }

        private void checkBox_destaca_campos_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonCheckBox RibbonCheckBox = (RibbonCheckBox)sender;
            if (RibbonCheckBox.Checked == true) Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading = (WdFieldShading)1;
            if (RibbonCheckBox.Checked == false) Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading = (WdFieldShading)2;
        }

        private void checkBox_mostra_indicadores_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonCheckBox RibbonCheckBox = (RibbonCheckBox)sender;
            if (RibbonCheckBox.Checked == true) Globals.ThisAddIn.Application.ActiveWindow.View.ShowBookmarks = true;
            if (RibbonCheckBox.Checked == false) Globals.ThisAddIn.Application.ActiveWindow.View.ShowBookmarks = false;
        }

        private void checkBox_vercodigo_campos_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonCheckBox RibbonCheckBox = (RibbonCheckBox)sender;
            if (RibbonCheckBox.Checked == true) Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes = true;
            if (RibbonCheckBox.Checked == false) Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes = false;
        }

        private void checkBox_atualizar_antes_de_imprimir_campos_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonCheckBox RibbonCheckBox = (RibbonCheckBox)sender;
            if (RibbonCheckBox.Checked == true) Globals.ThisAddIn.Application.Options.UpdateFieldsAtPrint = true;
            if (RibbonCheckBox.Checked == false) Globals.ThisAddIn.Application.Options.UpdateFieldsAtPrint = false;
        }

    }
}