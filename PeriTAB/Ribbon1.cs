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
        private bool first_time = true;
        String caminho_template = Path.GetTempPath() + "PeriTAB_Template_tmp.dotm";

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            
            //this.Evento_abre_cria_doc();
            //System.Version publish_version = new System.Version("9.9.9");
            //if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
            //{
            //    publish_version = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion;
            //}
            //Globals.Ribbons.Ribbon1.label1.Label = "PeriTAB " + publish_version;
            System.Version publish_version = Assembly.GetExecutingAssembly().GetName().Version;
            Globals.Ribbons.Ribbon1.label1.Label = "PeriTAB " + publish_version.Major + "." + publish_version.Minor + "." + publish_version.Build;

            //Class1 var_Class1 = new Class1();
            //var_Class1.Evento_DocumentOpen();
            //var_Class1.Evento_NewDocument();
            //var_Class1.Evento_DocumentBeforeClose();
            ////var_Class1.Evento_DocumentChange();
            //var_Class1.Evento_DocumentSync();
            //var_Class1.Evento_ProtectedViewWindowActivate();
            //var_Class1.Evento_WindowActivate();
            //var_Class1.Evento_WindowBeforeDoubleClick();
            //var_Class1.Evento_WindowBeforeRightClick();
            //var_Class1.Evento_WindowDeactivate();
            //var_Class1.Evento_WindowSelectionChange();
            //var_Class1.Evento_WindowSize();
            //var_Class1.Evento_SettingChanging();
            //var_Class1.Evento_TabDisposed();
            //var_Class1.Evento_Ribbon1Close();
            //var_Class1.Evento_menu1ItemsLoading();
            //Globals.ThisAddIn.Application.WindowSelectionChange += Application_WindowSelectionChange;
            Class_SelectionChange instace1 = new Class_SelectionChange(); instace1.Evento_WindowSelectionChange();
        }

        //public void abre_cria_doc(Microsoft.Office.Interop.Word.Document Doc)
        //{
        //    Class1 var_Class1 = new Class1();
        //    var_Class1.Evento_ContentControlOnEnter();
        //    MessageBox.Show("abre_cria_doc");
        //    if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes == true) checkBox2.Checked = true;
        //}


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

        private void toggleButton_Click(object sender, RibbonControlEventArgs e)
        {
            var botao_toggle = (Microsoft.Office.Tools.Ribbon.RibbonToggleButton)sender;
            if (botao_toggle.Checked == true) Globals.ThisAddIn.myCustomTaskPane.Visible = true;
            if (botao_toggle.Checked == false) Globals.ThisAddIn.myCustomTaskPane.Visible = false;
        }

        //public void Evento_abre_cria_doc()
        //{
        //    Globals.ThisAddIn.Application.DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(abre_cria_doc);
        //    ((Microsoft.Office.Interop.Word.ApplicationEvents4_Event)Globals.ThisAddIn.Application).NewDocument += new ApplicationEvents4_NewDocumentEventHandler(abre_cria_doc);
        //}
    }
}
