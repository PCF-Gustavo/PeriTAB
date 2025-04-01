using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.IO;
using System.Windows.Forms;
using System.Linq;
using System.Collections.Generic;

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
            private static readonly string private_caminho_template, private_caminho_AppData_Roaming_PeriTAB, private_caminho_preferences, private_caminho_arquivos_para_excluir;
            private static AddIn private_AddIn_PeriTAB;
            private static Template private_Template_PeriTAB;
            private static List<string> private_lista_arquivos_para_excluir;
            static Variables() // Bloco estático para definir o valor inicial das variáveis
            {
                private_caminho_template = Path.GetTempPath() + "PeriTAB_Template_tmp.dotm";
                private_caminho_AppData_Roaming_PeriTAB = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PeriTAB");
                private_caminho_preferences = Path.Combine(private_caminho_AppData_Roaming_PeriTAB, "preferences.xml");
                private_caminho_arquivos_para_excluir = Path.Combine(private_caminho_AppData_Roaming_PeriTAB, "arquivos_para_excluir.txt");
                if (File.Exists(private_caminho_arquivos_para_excluir))
                {
                    private_lista_arquivos_para_excluir = File.ReadAllLines(private_caminho_arquivos_para_excluir).ToList();
                }
                else
                {
                    private_lista_arquivos_para_excluir = new List<string>();
                }
            }

            // Declara variáveis públicas
            public static string caminho_template { get { return private_caminho_template; } }
            public static AddIn AddIn_PeriTAB { get { return private_AddIn_PeriTAB; } set { private_AddIn_PeriTAB = value; } }
            public static Template Template_PeriTAB { get { return private_Template_PeriTAB; } set { private_Template_PeriTAB = value; } }
            public static string caminho_AppData_Roaming_PeriTAB { get { return private_caminho_AppData_Roaming_PeriTAB; } }
            public static string caminho_preferences { get { return private_caminho_preferences; } }
            public static List<string> lista_arquivos_para_excluir { get { return private_lista_arquivos_para_excluir; } }
            public static string caminho_arquivos_para_excluir { get { return private_caminho_arquivos_para_excluir; } }
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

            Globals.Ribbons.Ribbon.label_nome.Label = "PeriTAB " + "1.2.1";

            ThisAddIn.Excluir_arquivos_da_lista(Variables.lista_arquivos_para_excluir);
        }

        public BuildingBlock inserir_autotexto(Range range, string autotextName)
        {
            if (range != null)
            {
                for (int i = 1; i <= Variables.Template_PeriTAB.BuildingBlockEntries.Count; i++)
                {
                    BuildingBlock bb = Variables.Template_PeriTAB.BuildingBlockEntries.Item(i);
                    if (bb.Name == autotextName)
                    {
                        bb.Insert(range);
                        Range Previous = Globals.ThisAddIn.Application.Selection.Range.Previous();
                        if (Previous != null) if (Previous.Fields.Count > 0) Previous.Words[1].Fields.Update();
                        return bb;
                    }
                }
            }
            return null;
        }
    }
}