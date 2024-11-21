using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.IO;
using System.Windows.Forms;

namespace PeriTAB
{
    public partial class Ribbon
    {
        // Define constantes
        const string quote = "\"";
        const string slash = @"\";

        // Cria instância das classes
        Class_CustomTaskPanes iClass_CustomTaskPanes = new Class_CustomTaskPanes();

        // Gerencia variáveis "globais"
        public class Variables
        {
            // Declara variáveis privadas
            private static readonly string private_caminho_template, private_caminho_AppData_Roaming_PeriTAB, private_caminho_preferences;
            private static readonly bool private_debugging;
            private static AddIn private_AddIn_PeriTAB;
            private static Template private_Template_PeriTAB;
            static Variables() // Bloco estático para definir o valor inicial das variáveis
            {
                private_caminho_template = Path.GetTempPath() + "PeriTAB_Template_tmp.dotm";
                private_caminho_AppData_Roaming_PeriTAB = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PeriTAB");
                private_caminho_preferences = Path.Combine(private_caminho_AppData_Roaming_PeriTAB, "preferences.xml");
                private_debugging = !System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed;
            }

            // Declara variáveis públicas
            public static string caminho_template { get { return private_caminho_template; } }
            public static AddIn AddIn_PeriTAB { get { return private_AddIn_PeriTAB; } set { private_AddIn_PeriTAB = value; } }
            public static Template Template_PeriTAB { get { return private_Template_PeriTAB; } set { private_Template_PeriTAB = value; } }
            public static string caminho_AppData_Roaming_PeriTAB { get { return private_caminho_AppData_Roaming_PeriTAB; } }
            public static string caminho_preferences { get { return private_caminho_preferences; } }
            public static bool debugging { get { return private_debugging; } }
        }



        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            //MessageBox.Show("Ribbon_Load");

            //Escreve o Template na pasta tmp e adiciona ela como suplemento.
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

            if (Variables.debugging)
            {
                button_teste.Visible = true;
                Globals.Ribbons.Ribbon.label_nome.Label = "PeriTAB Debugging";
            }
            else
            {
                // Escreve o número da versão
                System.Version publish_version = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion;
                Globals.Ribbons.Ribbon.label_nome.Label = "PeriTAB " + publish_version.Major + "." + publish_version.Minor + "." + publish_version.Build;
            }
        }

        public void inserir_autotexto(Range range, string autotextName)
        {
            // Procura pelo autotexto Numero_de_paginas_por_extenso no template_PeriTAB
            for (int i = 1; i <= Variables.Template_PeriTAB.BuildingBlockEntries.Count; i++)
            {
                BuildingBlock bb = Variables.Template_PeriTAB.BuildingBlockEntries.Item(i);
                if (bb.Name == autotextName)
                {
                    bb.Insert(range);
                    Range Previous = Globals.ThisAddIn.Application.Selection.Range.Previous();
                    if (Previous != null) if (Previous.Fields.Count > 0) Previous.Words[1].Fields.Update();
                    break;
                }
            }
        }

        public void atualiza_todos_campos(Document document)
        {
            // Percorre todas as StoryRanges no documento
            foreach (Range storyRange in document.StoryRanges)
            {
                // Atualiza os campos em cada StoryRange
                storyRange.Fields.Update();

                // Percorre os shapes (caixas de texto) em cada StoryRange
                foreach (Microsoft.Office.Interop.Word.Shape shape in document.Shapes)
                {
                    if (shape.Type == MsoShapeType.msoTextBox) // Verifica se é uma caixa de texto
                    {
                        // Atualiza os campos dentro da caixa de texto
                        shape.TextFrame.TextRange.Fields.Update();
                    }
                }
            }
        }

    }
}