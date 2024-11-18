using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.IO;
using System.Windows.Forms;

namespace PeriTAB
{
    public partial class Ribbon
    {
        // Cria instância das classes
        Class_CustomTaskPanes iClass_CustomTaskPanes = new Class_CustomTaskPanes();

        // Gerencia variáveis "globais"
        public class Variables
        {
            // Declara variáveis privadas

            //private static string var1 = Path.GetTempPath() + "PeriTAB_Template_tmp.dotm";
            //private static string var2 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PeriTAB");
            private static readonly string private_caminho_template, private_caminho_AppData_Roaming_PeriTAB, private_caminho_preferences;
            //private static string private_editBox_largura_Text, private_editBox_altura_Text;
            private static readonly bool private_debugging;
            private static AddIn private_AddIn_PeriTAB;
            private static Template private_Template_PeriTAB;
            static Variables() // Bloco estático para definir o valor inicial das variáveis
            {
                private_caminho_template = Path.GetTempPath() + "PeriTAB_Template_tmp.dotm";
                private_caminho_AppData_Roaming_PeriTAB = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PeriTAB");
                private_caminho_preferences = Path.Combine(private_caminho_AppData_Roaming_PeriTAB, "preferences.xml");
                private_debugging = !System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed;
                //MessageBox.Show("Variables");
            }

            // Declara variáveis públicas
            public static string caminho_template { get { return private_caminho_template; } }
            public static AddIn AddIn_PeriTAB { get { return private_AddIn_PeriTAB; } set { private_AddIn_PeriTAB = value; } }
            public static Template Template_PeriTAB { get { return private_Template_PeriTAB; } set { private_Template_PeriTAB = value; } }
            public static string caminho_AppData_Roaming_PeriTAB { get { return private_caminho_AppData_Roaming_PeriTAB; } }
            public static string caminho_preferences { get { return private_caminho_preferences; } }
            //public static string editBox_largura_Text { get { return private_editBox_largura_Text; } set { private_editBox_largura_Text = value; } }
            //public static string editBox_altura_Text { get { return private_editBox_altura_Text; } set { private_editBox_altura_Text = value; } }
            //public static X509Certificate2 cert { get { return var_cert; } set { var_cert = value; } }
            //public static IExternalSignature sig { get { return var_sig; } set { var_sig = value; } }
            public static bool debugging { get { return private_debugging; } }
        }

        // Define constantes
        const string quote = "\"";
        const string slash = @"\";

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            //iClass_Buttons.le_preferencias(Ribbon.Variables.caminho_preferences);
            //MessageBox.Show("Ribbon_Load");
            //Escreve o Template na pasta tmp e adiciona ela como suplemento.
            //try { File.WriteAllBytes(Variables.caminho_template, Properties.Resources.Normal); } catch (IOException ex) { MessageBox.Show("PeriTAB_Template_tmp.dotm em uso"); Globals.ThisAddIn.Application.Quit(); return; }
            //File.WriteAllBytes(Variables.caminho_template, Properties.Resources.Normal);
            try { File.WriteAllBytes(Variables.caminho_template, Properties.Resources.Normal); }
            catch (IOException)
            {
                if (!File.Exists(Variables.caminho_template))
                {
                    MessageBox.Show("PeriTAB_Template_tmp.dotm não encontrado"); Globals.ThisAddIn.Application.Quit(); return;
                }
            }
            Variables.AddIn_PeriTAB = Globals.ThisAddIn.Application.AddIns.Add(Variables.caminho_template);

            // Retorna o valor de PeriTAB como tipo Template
            foreach (Microsoft.Office.Interop.Word.Template template in Globals.ThisAddIn.Application.Templates)
            {
                if (template.Name == "PeriTAB_Template_tmp.dotm")
                {
                    Variables.Template_PeriTAB = template;
                    break;
                }
            }

            // Escreve o número da versão
            //System.Version publish_version = Assembly.GetExecutingAssembly().GetName().Version;
            //Globals.Ribbons.Ribbon.label_nome.Label = "PeriTAB " + publish_version.Major + "." + publish_version.Minor + "." + publish_version.Build;
            //Globals.Ribbons.Ribbon.label_nome.Label = "PeriTAB " + publish_version.Major + "." + publish_version.Minor + "." + publish_version.Build;

            if (Variables.debugging)
            {
                button_teste.Visible = true;
                Globals.Ribbons.Ribbon.label_nome.Label = "PeriTAB Debugging";
            }
            else
            {
                System.Version publish_version = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion;
                Globals.Ribbons.Ribbon.label_nome.Label = "PeriTAB " + publish_version.Major + "." + publish_version.Minor + "." + publish_version.Build;
            }


            //if (versao() != null)
            //{
            //    Globals.Ribbons.Ribbon.label_nome.Label = "PeriTAB " + versao().Major + "." + versao().Minor + "." + versao().Build;
            //    //Variables.debugging = false;
            //}
            //else
            //{
            //    //Variables.debugging = true;
            //    Globals.Ribbons.Ribbon.label_nome.Label = "PeriTAB Debugging";
            //}

            //if (Variables.debugging) { button_teste.Visible = true; }
        }
    }
}