using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using Task = System.Threading.Tasks.Task;


namespace PeriTAB
{
    public partial class Ribbon
    {
        private async void Button_inserir_sumario_Click(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon(sender, e, progress =>
            {
                Application Application = Globals.ThisAddIn.Application;
                string FullName = Globals.ThisAddIn.Application.ActiveDocument.FullName;
                Application.OrganizerCopy(Variables.Caminho_template, FullName, "Sumário 1", WdOrganizerObject.wdOrganizerObjectStyles);
                Application.OrganizerCopy(Variables.Caminho_template, FullName, "Sumário 2", WdOrganizerObject.wdOrganizerObjectStyles);
                Application.OrganizerCopy(Variables.Caminho_template, FullName, "Sumário 3", WdOrganizerObject.wdOrganizerObjectStyles);
                Application.OrganizerCopy(Variables.Caminho_template, FullName, "Sumário 4", WdOrganizerObject.wdOrganizerObjectStyles);
                Application.OrganizerCopy(Variables.Caminho_template, FullName, "Sumário 5", WdOrganizerObject.wdOrganizerObjectStyles);
                Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldTOC, slash + "h " + slash + "z " + slash + "t " + quote + "05 - Seção_1 (PeriTAB);1;06 - Seção_2 (PeriTAB);2;07 - Seção_3 (PeriTAB);3;08 - Seção_4 (PeriTAB);4;09 - Seção_5 (PeriTAB);5" + quote, false);
                return Task.CompletedTask;
            });
        }

        private async void Button_inserir_pagina_Click(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon(sender, e, progress =>
            {
                Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldEmpty, "PAGE", false);
                return Task.CompletedTask;
            });
        }

        private async void Button_inserir_pagina_extenso_Click(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon(sender, e, progress =>
            {
                Inserir_autotexto(Globals.ThisAddIn.Application.Selection.Range, "pagina_atual_por_extenso_PeriTAB");
                return Task.CompletedTask;
            });
        }

        private async void Button_inserir_paginas_Click(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon(sender, e, progress =>
            {
                Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldEmpty, "NUMPAGES", false);
                return Task.CompletedTask;
            });
        }

        private async void Button_inserir_paginas_extenso_Click(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon(sender, e, progress =>
            {
                Inserir_autotexto(Globals.ThisAddIn.Application.Selection.Range, "paginas_por_extenso_PeriTAB");
                return Task.CompletedTask;
            });
        }

        private async void Button_inserir_ano_Click(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon(sender, e, progress =>
            {
                Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldEmpty, "DATE " + slash + "@ " + quote + "yyyy" + quote, false);
                return Task.CompletedTask;
            });
        }

        private async void Button_atualiza_campos_Click(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon(sender, e, progress =>
            {
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
                    progress?.Report((int)((i * 10.0) / campos.Count));
                }
                return Task.CompletedTask;
            }, barra_de_progresso: true, desabilitar_ScreenUpdating: true, desabilitar_TrackRevisions: true);
        }

        private void CheckBox_destaca_campos_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonCheckBox RibbonCheckBox = (RibbonCheckBox)sender;
            if (RibbonCheckBox.Checked == true) Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading = (WdFieldShading)1;
            if (RibbonCheckBox.Checked == false) Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading = (WdFieldShading)2;
        }

        private void CheckBox_mostra_indicadores_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonCheckBox RibbonCheckBox = (RibbonCheckBox)sender;
            if (RibbonCheckBox.Checked == true) Globals.ThisAddIn.Application.ActiveWindow.View.ShowBookmarks = true;
            if (RibbonCheckBox.Checked == false) Globals.ThisAddIn.Application.ActiveWindow.View.ShowBookmarks = false;
        }

    }
}