using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using PeriTAB.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Deployment;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

    namespace PeriTAB
{
    
    public partial class Ribbon1
    {
        private bool first_time = true;
        String caminho_template = Path.GetTempPath() + "PeriTAB_Template_tmp.dotm";

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void Anybutton_Click(object sender, RibbonControlEventArgs e)
        {
            if (first_time) //Escreve o Template na pasta tmp e adiciona ela como suplemento.
            {
                File.WriteAllBytes(caminho_template, Properties.Resources.Normal_copy);
                Globals.ThisAddIn.Application.AddIns.Add(caminho_template);
                first_time = false;
            }
            
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
                    Globals.ThisAddIn.Application.Run("atualizar_todos_campos");
                    break;
                case "button7":
                    Globals.ThisAddIn.Application.Run("moeda_por_extenso");
                    break;
                case "button8":
                    Globals.ThisAddIn.Application.Run("inteiro_por_extenso");
                    break;
                case "button9":
                    string[] aStyles = { "01 - corpo de texto", "02 - seções e subseções", "03 - citações", "04 - enumerações", "05 - figuras", "06 - legendas de figuras", "07 - notas de rodapé", "08 - legendas de tabelas", "09 - quesitos", "10 - anexo" };
                    for (int i = 0; i <= aStyles.Length - 1; i++)
                    {
                        Globals.ThisAddIn.Application.OrganizerCopy(caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, aStyles[i], WdOrganizerObject.wdOrganizerObjectStyles);
                    }
                    break;
                case "button10":
                    Globals.ThisAddIn.Application.Run("DelUnusedStyles");
                    break;
                default:
                    break;
            }
        }
    }
}
