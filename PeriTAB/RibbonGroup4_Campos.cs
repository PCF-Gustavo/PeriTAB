using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using Tarefa = System.Threading.Tasks.Task;


namespace PeriTAB
{
    public partial class Ribbon
    {

        private void button_inserir_sumario_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.OrganizerCopy(PeriTAB.Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "Sumário 1", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(PeriTAB.Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "Sumário 2", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(PeriTAB.Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "Sumário 3", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(PeriTAB.Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "Sumário 4", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(PeriTAB.Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "Sumário 5", WdOrganizerObject.wdOrganizerObjectStyles);
            //Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldTOC, slash + "h " + slash + "z " + slash + "t " + quote + "04 - Seções (PeriTAB);1" + quote, false);
            Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldTOC, slash + "h " + slash + "z " + slash + "t " + quote + "05 - Seção_1 (PeriTAB);1;06 - Seção_2 (PeriTAB);2;07 - Seção_3 (PeriTAB);3;08 - Seção_4 (PeriTAB);4;09 - Seção_5 (PeriTAB);5" + quote, false);
            //Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldTOC, slash + "h " + slash + "z " + slash + "t " + quote + "04A - SEÇÃO_1 (PERITAB);1;04B - SEÇÃO_2 (PERITAB);2;04C - SEÇÃO_3 (PERITAB);3;04D - SEÇÃO_4 (PERITAB);4" + quote + " " + slash + "c " + quote + "Figura" + quote + " " + slash + "c " + quote + "Tabela" + quote, false);

        }

        private void button_inserir_pagina_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldEmpty, "PAGE", false);
        }

        private void button_inserir_pagina_extenso_Click(object sender, RibbonControlEventArgs e)
        {
            //Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldPage, slash + "* Cardtext " + slash + "* Lower", false);

            //// Procura pelo template_PeriTAB  --- dá pra melhorar chamando o template de forma global
            //Template template_PeriTAB = null;
            //foreach (Microsoft.Office.Interop.Word.Template template in Globals.ThisAddIn.Application.Templates) 
            //{
            //    if (template.Name == "PeriTAB_Template_tmp.dotm")
            //    {
            //        template_PeriTAB = template;
            //        break;
            //    }
            //}

            inserir_autotexto(Globals.ThisAddIn.Application.Selection.Range, "Numero_de_paginas_por_extenso");



            
            //string autotextName = "Numero_de_paginas_por_extenso";
            //BuildingBlockEntries buildingBlockEntries = Variables.Template_PeriTAB.BuildingBlockEntries;
            //for (int i = 1; i <= buildingBlockEntries.Count; i++)
            //{
            //    BuildingBlock bb = buildingBlockEntries.Item(i);
            //    if (bb.Name == autotextName)
            //    {
            //        bb.Insert(Globals.ThisAddIn.Application.Selection.Range);
            //        Globals.ThisAddIn.Application.Selection.Range.Previous().Words[1].Fields.Update();
            //        break;
            //    }
            //}
        }



        private void button_inserir_paginas_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldEmpty, "NUMPAGES", false);
        }
        private void button_inserir_paginas_extenso_Click(object sender, RibbonControlEventArgs e)
        {
            //Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldNumPages, slash + "* Cardtext " + slash + "* Lower", false);

            //// Procura pelo template_PeriTAB  --- dá pra melhorar chamando o template de forma global
            //Template template_PeriTAB = null;
            //foreach (Microsoft.Office.Interop.Word.Template template in Globals.ThisAddIn.Application.Templates)
            //{
            //    if (template.Name == "PeriTAB_Template_tmp.dotm")
            //    {
            //        template_PeriTAB = template;
            //        break;
            //    }
            //}

            //// Procurar pelo autotexto Pagina_atual_por_extenso no template_PeriTAB
            //string autotextName = "Pagina_atual_por_extenso";
            //BuildingBlockEntries buildingBlockEntries = Variables.Template_PeriTAB.BuildingBlockEntries;
            //for (int i = 1; i <= buildingBlockEntries.Count; i++)
            //{
            //    BuildingBlock bb = buildingBlockEntries.Item(i);
            //    if (bb.Name == autotextName)
            //    {
            //        bb.Insert(Globals.ThisAddIn.Application.Selection.Range);
            //        Globals.ThisAddIn.Application.Selection.Range.Previous().Words[1].Fields.Update();
            //        break;
            //    }
            //}
            inserir_autotexto(Globals.ThisAddIn.Application.Selection.Range, "Pagina_atual_por_extenso");

        }

        private void button_inserir_ano_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldEmpty, "DATE " + slash + "@ " + quote + "yyyy" + quote , false);
        }

        private static Dictionary<string, string> dict_Secao_de_conclusao_e_Fim_do_preambulo = new Dictionary<string, string>()
        {
             { "RESPOSTA AOS QUESITOS", "respondendo aos quesitos formulados, abaixo transcritos" }
            ,{ "CONCLUSÃO" , "atendendo ao abaixo transcrito" }
        };

        private async void button_atualiza_campos_Click(object sender, RibbonControlEventArgs e)
        {
            //// Atualiza a UI na Thread principal
            //RibbonButton RibbonButton = (RibbonButton)sender;
            //RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            //RibbonButton.Enabled = false;

            //await Tarefa.Run(() =>
            //{
            //    Globals.ThisAddIn.Application.Run("atualiza_todos_campos");
            //    Globals.ThisAddIn.Application.DisplayStatusBar = true; Globals.ThisAddIn.Application.StatusBar = "Campos atualizados com sucesso.";
            //});

            //// Após a execução das tarefas, atualiza a UI na Thread principal
            //RibbonButton.Image = Properties.Resources.atualizar;
            //RibbonButton.Enabled = true;
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;

            await Tarefa.Run(() =>
            {
                Globals.Ribbons.Ribbon.atualiza_todos_campos(Globals.ThisAddIn.Application.ActiveDocument);
            });

            // Após a execução das tarefas, atualiza a UI na Thread principal
            Globals.ThisAddIn.Application.DisplayStatusBar = true; Globals.ThisAddIn.Application.StatusBar = "Campos atualizados com sucesso.";
            RibbonButton.Image = Properties.Resources.atualizar;
            RibbonButton.Enabled = true;
        }

        private async void button_minuscula_campos_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            menu_formatacao_campos.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            await Tarefa.Run(() =>
            {


                //if (Globals.ThisAddIn.Application.Selection.Paragraphs.Count == 1) 
                //{
                //    Globals.ThisAddIn.Application.Selection.Paragraphs[1].Range.Select();
                //}

                foreach (Field f in Globals.ThisAddIn.Application.Selection.Fields)
                {
                    //MessageBox.Show(f.Code.Text);
                    string texto_campo = f.Code.Text;

                    if (texto_campo.IndexOf(slash + "* Upper ") != -1)
                    {
                        //MessageBox.Show("1");
                        f.Code.Text = texto_campo.Replace(slash + "* Upper ", slash + "* Lower ");
                        f.Update();
                        continue;
                    }
                    if (texto_campo.IndexOf(slash + "* FirstCap ") != -1)
                    {
                        //MessageBox.Show("2");
                        f.Code.Text = texto_campo.Replace(slash + "* FirstCap ", slash + "* Lower ");
                        f.Update();
                        continue;
                    }
                    if (texto_campo.IndexOf(slash + "* Caps ") != -1)
                    {
                        //MessageBox.Show("3");
                        f.Code.Text = texto_campo.Replace(slash + "* Caps ", slash + "* Lower ");
                        f.Update();
                        continue;
                    }

                    if (texto_campo.Replace(" ", "").IndexOf(slash + "*Lower") == -1)
                    {
                        //MessageBox.Show("4");
                        f.Code.Text = texto_campo + " " + slash + "* Lower ";
                        f.Update();
                    }

                }

                /*}).Start();*/
            });

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            RibbonButton.Image = Properties.Resources.formatacao2;
            menu_formatacao_campos.Enabled = true;
        }

        private void checkBox_destaca_campos_Click(object sender, RibbonControlEventArgs e)
        {
            var Botao_checkBox = (Microsoft.Office.Tools.Ribbon.RibbonCheckBox)sender;
            if (Botao_checkBox.Checked == true) Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading = (WdFieldShading)1;
            if (Botao_checkBox.Checked == false) Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading = (WdFieldShading)2;
        }
        private void checkBox_mostra_indicadores_Click(object sender, RibbonControlEventArgs e)
        {
            var Botao_checkBox = (Microsoft.Office.Tools.Ribbon.RibbonCheckBox)sender;
            if (Botao_checkBox.Checked == true) Globals.ThisAddIn.Application.ActiveWindow.View.ShowBookmarks = true;
            if (Botao_checkBox.Checked == false) Globals.ThisAddIn.Application.ActiveWindow.View.ShowBookmarks = false;
        }

        private void checkBox_vercodigo_campos_Click(object sender, RibbonControlEventArgs e)
        {
            var Botao_checkBox = (Microsoft.Office.Tools.Ribbon.RibbonCheckBox)sender;
            if (Botao_checkBox.Checked == true) Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes = true;
            if (Botao_checkBox.Checked == false) Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes = false;
        }
        private void checkBox_atualizar_antes_de_imprimir_campos_Click(object sender, RibbonControlEventArgs e)
        {
            var Botao_checkBox = (Microsoft.Office.Tools.Ribbon.RibbonCheckBox)sender;
            if (Botao_checkBox.Checked == true) Globals.ThisAddIn.Application.Options.UpdateFieldsAtPrint = true;
            if (Botao_checkBox.Checked == false) Globals.ThisAddIn.Application.Options.UpdateFieldsAtPrint = false;
        }

    }
}