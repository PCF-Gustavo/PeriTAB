using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;


namespace PeriTAB
{    
    public partial class Ribbon1
    {
        //private bool first_time = true;
        String caminho_template = Path.GetTempPath() + "PeriTAB_Template_tmp.dotm";

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //Escreve o Template na pasta tmp e adiciona ela como suplemento.
            File.WriteAllBytes(caminho_template, Properties.Resources.Normal);
            Globals.ThisAddIn.Application.AddIns.Add(caminho_template);

            // Escreve o número da versão
            System.Version publish_version = Assembly.GetExecutingAssembly().GetName().Version;
            Globals.Ribbons.Ribbon1.label1.Label = "PeriTAB " + publish_version.Major + "." + publish_version.Minor + "." + publish_version.Build;

            Class_Buttons iClass_Buttons = new Class_Buttons(); iClass_Buttons.Default();
            Class_DocChange_Event iClass_DocChange_Event = new Class_DocChange_Event(); iClass_DocChange_Event.Evento_DocChange();
            Class_DocSave_Event iClass_DocSave_Event = new Class_DocSave_Event(); iClass_DocSave_Event.Evento_DocSave();
            Class_New_or_Open_Event iClass_New_or_Open_Event = new Class_New_or_Open_Event(); iClass_New_or_Open_Event.Evento_New_or_Open();
            Class_SelectionChange_Event iClass_SelectionChange_Event = new Class_SelectionChange_Event(); iClass_SelectionChange_Event.Evento_SelectionChange();
            Class_WindowActivate_Event iClass_WindowActivate_Event = new Class_WindowActivate_Event(); iClass_WindowActivate_Event.Evento_WindowActivate();
        }

        private void Anybutton_Click(object sender, RibbonControlEventArgs e)
        {
            
            var botao = (Microsoft.Office.Tools.Ribbon.RibbonButton)sender;
            switch (botao.Name)
            {
                case "button1":
                    Globals.ThisAddIn.Application.Run("confere_numeracao_legendas");
                    break;
                case "button2":
                    Globals.ThisAddIn.Application.Run("alinha_legenda");
                    break;
                case "button3":
                    Globals.ThisAddIn.Application.Run("renomeia_documento");
                    break;
                case "button4":
                    Globals.ThisAddIn.Application.Run("alterna_visualizacao_campos");
                    break;
                case "button5":
                    Globals.ThisAddIn.Application.Run("alterna_destaque_campos");
                    break;
                case "button6":
                    Globals.ThisAddIn.Application.Run("atualiza_todos_campos");
                    break;
                case "button7":
                    Globals.ThisAddIn.Application.Run("moeda_por_extenso");
                    break;
                case "button8":
                    Globals.ThisAddIn.Application.Run("inteiro_por_extenso");
                    break;
                case "button9":
                    string[] aStyles = { "01 - Sem Formatação (PeriTAB)", "02 - Corpo do Texto (PeriTAB)", "03 - Citações (PeriTAB)", "04 - Seções (PeriTAB)", "05 - Enumerações (PeriTAB)", "06 - Figuras (PeriTAB)", "07 - Legendas de Figuras (PeriTAB)", "08 - Legendas de Tabelas (PeriTAB)", "09 - Quesitos (PeriTAB)", "Normal", "Texto de nota de rodapé", "Legenda" };
                    for (int i = 0; i <= aStyles.Length - 1; i++)
                    {
                        Globals.ThisAddIn.Application.OrganizerCopy(caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, aStyles[i], WdOrganizerObject.wdOrganizerObjectStyles);
                    }
                    break;
                case "button10":
                    Globals.ThisAddIn.Application.Run("limpa_estilos");
                    break;
                default:
                    break;
            }
        }

        private void toggleButton_Click(object sender, RibbonControlEventArgs e)
        {
            var botao_toggle = (Microsoft.Office.Tools.Ribbon.RibbonToggleButton)sender;
            if (botao_toggle.Checked == true) Globals.ThisAddIn.TaskPane1.Visible = true;
            if (botao_toggle.Checked == false) Globals.ThisAddIn.TaskPane1.Visible = false;
        }

    }
}
