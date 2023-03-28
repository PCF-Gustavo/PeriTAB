using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using PeriTAB.Properties;
using System;
using System.Windows;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Deployment;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using System.Reflection.Emit;
using System.Reflection;
using System.Configuration;
using System.ComponentModel.Design.Serialization;
using System.ComponentModel.Design;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Drawing;

namespace PeriTAB
{    
    public partial class Ribbon1
    {
        private bool first_time = true;
        String caminho_template = Path.GetTempPath() + "PeriTAB_Template_tmp.dotm";

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {            
            this.Evento_abre_cria_doc();
            //System.Version publish_version = new System.Version("9.9.9");
            //if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
            //{
            //    publish_version = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion;
            //}
            //Globals.Ribbons.Ribbon1.label1.Label = "PeriTAB " + publish_version;
            System.Version publish_version = Assembly.GetExecutingAssembly().GetName().Version;
            Globals.Ribbons.Ribbon1.label1.Label = "PeriTAB " + publish_version.Major + "." + publish_version.Minor + "." + publish_version.Build;
        }

        public void abre_cria_doc(Microsoft.Office.Interop.Word.Document Doc)
        {
            MessageBox.Show("abre_cria_doc");
            if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes == true) checkBox2.Checked = true;
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

        public void Evento_abre_cria_doc()
        {
            Globals.ThisAddIn.Application.DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(abre_cria_doc);
            ((Microsoft.Office.Interop.Word.ApplicationEvents4_Event)Globals.ThisAddIn.Application).NewDocument += new ApplicationEvents4_NewDocumentEventHandler(abre_cria_doc);
        }
    }
}
