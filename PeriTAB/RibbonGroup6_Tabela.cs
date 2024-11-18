using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using Tarefa = System.Threading.Tasks.Task;


namespace PeriTAB
{
    public partial class Ribbon
    {
        private async void button_legenda_tabela_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            menu_inserir_tabela.Image = Properties.Resources.load_icon_png_7969;
            menu_inserir_tabela.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Range r = Globals.ThisAddIn.Application.Selection.Range;

            await Tarefa.Run(() =>
            {


                string estilo_nome_baseado = "Legenda";
                Globals.ThisAddIn.Application.OrganizerCopy(PeriTAB.Ribbon.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);

                List<Table> list_Table = new List<Table>();
                foreach (Table itable in Globals.ThisAddIn.Application.Selection.Tables)
                {
                    list_Table.Add(itable);
                }
                foreach (Table itable in list_Table)
                {
                    itable.Select();
                    //MessageBox.Show(itable.Range.Text);
                    //MessageBox.Show(Globals.ThisAddIn.Application.Selection.Paragraphs[1].Previous().Range.Text);
                    if (Globals.ThisAddIn.Application.Selection.Paragraphs[1].Previous() != null)
                    {
                        if (Globals.ThisAddIn.Application.Selection.Paragraphs[1].Previous().Range.Characters.Count >= 7)
                        {
                            if (Globals.ThisAddIn.Application.Selection.Paragraphs[1].Previous().Range.Text.Substring(0, 7) == "Tabela ")
                            {
                                //r.Select();
                                //Globals.ThisAddIn.Application.ScreenUpdating = true;
                                //return;
                                continue;
                            }
                        }
                    }

                    Globals.ThisAddIn.Application.Selection.InsertCaption(Label: "Tabela", Title: " " + ((char)8211).ToString(), TitleAutoText: "", Position: WdCaptionPosition.wdCaptionPositionAbove, ExcludeLabel: 0);
                    Globals.ThisAddIn.Application.Selection.set_Style((object)"08 - Legendas de Tabelas (PeriTAB)");
                    Globals.ThisAddIn.Application.Selection.InsertAfter(" ");
                    Globals.ThisAddIn.Application.Run("alinha_legenda");
                }

                /*}).Start();*/
            });

            r.Select();
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            menu_inserir_tabela.Image = Properties.Resources._;
            menu_inserir_tabela.Enabled = true;
        }

        private async void button_centralizar_tabela_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            menu_formatacao_tabela.Image = Properties.Resources.load_icon_png_7969;
            menu_formatacao_tabela.Enabled = false;
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            await Tarefa.Run(() =>
            {


                foreach (Table itable in Globals.ThisAddIn.Application.Selection.Tables)
                {
                    itable.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    itable.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
                    foreach (Paragraph iParagraph in itable.Range.Paragraphs)
                    {
                        iParagraph.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                }
                //foreach (Cell icell in Globals.ThisAddIn.Application.Selection.Cells)
                //{
                //    icell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                //}
                //foreach (Paragraph iParagraph in Globals.ThisAddIn.Application.Selection.Paragraphs)
                //{
                //    if (iParagraph.Range.Information[WdInformation.wdWithInTable]) 
                //    {
                //        iParagraph.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                //    }
                //}


                /*}).Start();*/
            });

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            menu_formatacao_tabela.Image = Properties.Resources.formatacao2;
            menu_formatacao_tabela.Enabled = true;
        }
    }
}