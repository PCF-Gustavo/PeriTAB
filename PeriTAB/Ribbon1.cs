using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Forms;
using System.Windows.Markup;
using System.Xml.Linq;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Windows.Media.TextFormatting;
using System.Configuration;
using System.Collections.Specialized;

namespace PeriTAB
{    
    public partial class Ribbon1
    {
        public class Variables
        {                     
            private static string var1 = Path.GetTempPath() + "PeriTAB_Template_tmp.dotm";
            private static string var2 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PeriTAB");
            private static string var3 = "";
            private static string var4 = "";
            public static string caminho_template { get { return var1; } set { } }
            public static string caminho_AppData_Roaming_PeriTAB { get { return var2; } set { } }
            public static string editBox_largura_Text { get { return var3; } set { var3 = value; } }
            public static string editBox_altura_Text { get { return var4; } set { var4 = value; } }
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //Escreve o Template na pasta tmp e adiciona ela como suplemento.
            File.WriteAllBytes(Variables.caminho_template, Properties.Resources.Normal);
            Globals.ThisAddIn.Application.AddIns.Add(Variables.caminho_template);

            // Escreve o número da versão
            System.Version publish_version = Assembly.GetExecutingAssembly().GetName().Version;
            Globals.Ribbons.Ribbon1.label_nome.Label = "PeriTAB " + publish_version.Major + "." + publish_version.Minor + "." + publish_version.Build;


        }

        private void button_confere_num_legenda_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("confere_numeracao_legendas");
        }

        private void button_alinha_legenda_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("alinha_legenda");
        }

        private void button_renomeia_documento_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("renomeia_documento");
        }

        private void button_atualiza_campos_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("atualiza_todos_campos");
        }
        private void checkBox_destaca_campos_Click(object sender, RibbonControlEventArgs e)
        {
            var Botao_checkBox = (Microsoft.Office.Tools.Ribbon.RibbonCheckBox)sender;
            if (Botao_checkBox.Checked == true) Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading = (WdFieldShading)1;
            if (Botao_checkBox.Checked == false) Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading = (WdFieldShading)2;
        }

        private void checkBox_vercodigo_campos_Click(object sender, RibbonControlEventArgs e)
        {
            var Botao_checkBox = (Microsoft.Office.Tools.Ribbon.RibbonCheckBox)sender;
            if (Botao_checkBox.Checked == true) Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes = true;
            if (Botao_checkBox.Checked == false) Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes = false;
        }

        private void button_moeda_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("moeda_por_extenso");
        }

        private void button_inteiro_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("inteiro_por_extenso");
        }

        private void button_importa_estilos_Click(object sender, RibbonControlEventArgs e)
        {
            string[] aStyles = { "01 - Sem Formatação (PeriTAB)", "02 - Corpo do Texto (PeriTAB)", "03 - Citações (PeriTAB)", "04 - Seções (PeriTAB)", "05 - Enumerações (PeriTAB)", "06 - Figuras (PeriTAB)", "07 - Legendas de Figuras (PeriTAB)", "08 - Legendas de Tabelas (PeriTAB)", "09 - Quesitos (PeriTAB)", "Normal", "Texto de nota de rodapé", "Legenda" };
            for (int i = 0; i <= aStyles.Length - 1; i++)
            {
                Globals.ThisAddIn.Application.OrganizerCopy(Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, aStyles[i], WdOrganizerObject.wdOrganizerObjectStyles);
            }
        }

        private void button_limpa_estilos_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("limpa_estilos");
        }

        private void toggleButton_estilos_Click(object sender, RibbonControlEventArgs e)
        {
            var botao_toggle = (Microsoft.Office.Tools.Ribbon.RibbonToggleButton)sender;
            if (botao_toggle.Checked == true) Globals.ThisAddIn.TaskPane1.Visible = true;
            if (botao_toggle.Checked == false) Globals.ThisAddIn.TaskPane1.Visible = false;
        }

        private void button_inserir_sumario_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void dropDown1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void button_cola_imagem_Click(object sender, RibbonControlEventArgs e)
        {

            if (System.Windows.Clipboard.ContainsData("FileDrop"))
            {
                object obj = System.Windows.Clipboard.GetData("FileDrop");
                string[] pathfile = (string[])obj;

                string[] pathfile2 = { "" };
                int n = 0;
                for (int i = 0; i <= pathfile.Length - 1; i++)
                {
                    if (File.Exists(pathfile[i]))
                    {
                        string extensao = pathfile[i].Substring(pathfile[i].Length - 4);
                        if (extensao == ".jpg" | extensao == "jpeg" | extensao == ".png" | extensao == ".bmp" | extensao == ".gif" | extensao == "tiff") //Se tem extensao de imagem
                        {
                            Array.Resize(ref pathfile2, n+1);
                            pathfile2[n] = pathfile[i];
                            n++;
                        }
                    }
                }

                if (pathfile2[0] != "")                 
                {
                    if (dropDown_ordem.SelectedItem.Label == "Alfabética") { Array.Sort(pathfile2); } //Ordem alfabética               

                    for (int i = 0; i <= pathfile2.Length - 1; i++)
                    {
                        Globals.ThisAddIn.Application.ScreenUpdating = false;

                        bool link = false; bool save = true;
                        if (Globals.Ribbons.Ribbon1.checkBox_referencia.Checked == true) { link = true; save = false; }

                        if (checkBox_largura.Checked)
                        {
                            InlineShape imagem = Globals.ThisAddIn.Application.Selection.InlineShapes.AddPicture(pathfile2[i], link, save);

                            MsoTriState LockAspectRatio_i = imagem.LockAspectRatio;
                            imagem.LockAspectRatio = (MsoTriState)1;
                            string larg_string = Globals.Ribbons.Ribbon1.editBox_largura.Text;
                            float.TryParse(larg_string, out float larg);
                            imagem.Width = Globals.ThisAddIn.Application.CentimetersToPoints(larg);

                            imagem.LockAspectRatio = LockAspectRatio_i;
                        }

                        if (checkBox_altura.Checked)
                        {
                            InlineShape imagem = Globals.ThisAddIn.Application.Selection.InlineShapes.AddPicture(pathfile2[i], link, save);

                            MsoTriState LockAspectRatio_i = imagem.LockAspectRatio;
                            imagem.LockAspectRatio = (MsoTriState)1;
                            string alt_string = Globals.Ribbons.Ribbon1.editBox_altura.Text;
                            float.TryParse(alt_string, out float alt);
                            imagem.Height = Globals.ThisAddIn.Application.CentimetersToPoints(alt);

                            imagem.LockAspectRatio = LockAspectRatio_i;
                        }

                        if (i != pathfile2.Length -1) //Exceto última imagem
                        {

                            switch (dropDown_separador.SelectedItem.Label) //Insere separador
                            {
                                case "Espaço":
                                    Globals.ThisAddIn.Application.Selection.InsertAfter(" ");
                                    Globals.ThisAddIn.Application.Selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                                    break;
                                case "Parágrafo":
                                    Globals.ThisAddIn.Application.Selection.InsertAfter(System.Environment.NewLine);
                                    Globals.ThisAddIn.Application.Selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                                    break;
                            }
                        }
                        Globals.ThisAddIn.Application.ScreenUpdating = true;
                    }
                }
                else MessageBox.Show("Imagem não encontrada.");
            }
            else MessageBox.Show("Imagem não encontrada.");
        }      

        private void checkBox_largura_Click(object sender, RibbonControlEventArgs e)
        {
            if (Variables.editBox_largura_Text == "") { Variables.editBox_largura_Text = Class_Buttons.preferences.largura; }

            if (checkBox_largura.Checked)
            {
                checkBox_altura.Checked = false;
                editBox_altura.Enabled = false;
                editBox_altura.Text = "";
                editBox_largura.Enabled = true;
                editBox_largura.Text = Variables.editBox_largura_Text;
            }
            else
            {
                checkBox_largura.Checked = true;
            }
        }

        private void checkBox_altura_Click(object sender, RibbonControlEventArgs e)
        {
            if (Variables.editBox_altura_Text == "") { Variables.editBox_altura_Text = Class_Buttons.preferences.altura; }

            if (checkBox_altura.Checked)
            {
                checkBox_largura.Checked = false;
                editBox_largura.Enabled= false;
                editBox_largura.Text = "";
                editBox_altura.Enabled = true;
                editBox_altura.Text = Variables.editBox_altura_Text;
            }
        else
        {
                checkBox_altura.Checked = true;
            }
}

        private void checkBox_referencia_Click(object sender, RibbonControlEventArgs e)
        {
            if (checkBox_referencia.Checked) {
                System.Windows.Forms.MessageBox.Show("Cuidado! Excluir/mover/renomear o arquivo da imagem causará perda de referência.");
            }
        }
       

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            foreach (InlineShape s in Globals.ThisAddIn.Application.ActiveDocument.InlineShapes)
            {
                if (s.Type == WdInlineShapeType.wdInlineShapePicture | s.Type == WdInlineShapeType.wdInlineShapeLinkedPicture)
                {
                    string str1 = "pkg:name=" + "\"" + "/word/media/";
                    string str2 = "\"" + " pkg:contentType=";

                    string xml1 = s.Range.WordOpenXML;

                    int index1 = xml1.IndexOf(str1);
                    int index2 = xml1.Substring(index1, 100).IndexOf(str2);

                    string pathimage = xml1.Substring(index1 + 10, index2 - 10);


                    string fullname = Globals.ThisAddIn.Application.ActiveDocument.FullName;

                    MessageBox.Show(fullname + pathimage.Replace("/", @"\"));
                }

            }
            string xml = Globals.ThisAddIn.Application.ActiveDocument.WordOpenXML;
            long compress = xml.IndexOf("<w:doNotAutoCompressPictures/>");

            if (compress.ToString() == "-1")
            {
                //Comprime
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Spire.Doc.Document document = new Spire.Doc.Document(@"C:\Users\Gustavo\Desktop\FFFF.docx");

            foreach (Spire.Doc.Section section in document.Sections)
            {
                foreach (Spire.Doc.Documents.Paragraph paragraph in section.Paragraphs)
                {
                    foreach (Spire.Doc.DocumentObject docObject in paragraph.ChildObjects)
                    {
                        if (docObject.DocumentObjectType == DocumentObjectType.Picture)
                        {
                            DocPicture pic = docObject as DocPicture;
                            MessageBox.Show(pic.Image.HorizontalResolution.ToString());
                        }
                    }
                }
            }
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Directory.Exists(Ribbon1.Variables.caminho_AppData_Roaming_PeriTAB)) { Directory.CreateDirectory(Ribbon1.Variables.caminho_AppData_Roaming_PeriTAB); } //Cria a pasta AppData/Roaming/PeriTAB caso não exista

            string preferences_path = Path.Combine(Ribbon1.Variables.caminho_AppData_Roaming_PeriTAB, "preferences");

            string preferences = "";

            if (Globals.Ribbons.Ribbon1.editBox_largura.Text != ""){ preferences += "<largura>" + Globals.Ribbons.Ribbon1.editBox_largura.Text + "</largura>" + System.Environment.NewLine; } else if (Variables.editBox_largura_Text != "") { preferences += "<largura>" + Variables.editBox_largura_Text + "</largura>" + System.Environment.NewLine;} else { preferences += "<largura>" + Class_Buttons.preferences.largura + "</largura>" + System.Environment.NewLine; }
            if (Globals.Ribbons.Ribbon1.editBox_altura.Text != "") { preferences += "<altura>" + Globals.Ribbons.Ribbon1.editBox_altura.Text + "</altura>" + System.Environment.NewLine; } else if (Variables.editBox_altura_Text != "") { preferences += "<altura>" + Variables.editBox_altura_Text + "</altura>" + System.Environment.NewLine; } else { preferences += "<altura>" + Class_Buttons.preferences.altura + "</altura>" + System.Environment.NewLine; }
            preferences += "<largura_checked>" + Globals.Ribbons.Ribbon1.checkBox_largura.Checked.ToString() + "</largura_checked>" + System.Environment.NewLine;
            preferences += "<ordem>" + Globals.Ribbons.Ribbon1.dropDown_ordem.SelectedItem.Label + "</ordem>" + System.Environment.NewLine;
            preferences += "<separador>" + Globals.Ribbons.Ribbon1.dropDown_separador.SelectedItem.Label + "</separador>" + System.Environment.NewLine;

            File.WriteAllText(preferences_path, preferences);
        }

        private void editBox_largura_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (Variables.editBox_largura_Text == "") { Variables.editBox_largura_Text = Class_Buttons.preferences.largura; }

            if (float.TryParse(editBox_largura.Text, out float alt) & alt.ToString() == editBox_largura.Text)
            {
                Variables.editBox_largura_Text = editBox_largura.Text;
            }
            else 
            {
                editBox_largura.Text = Variables.editBox_largura_Text;
            }            
        }

        private void editBox_altura_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (Variables.editBox_altura_Text == "") { Variables.editBox_altura_Text = Class_Buttons.preferences.altura; }

            if (float.TryParse(editBox_altura.Text, out float alt) & alt.ToString() == editBox_altura.Text)
            {
                Variables.editBox_altura_Text = editBox_altura.Text;
            }
            else
            {
                editBox_altura.Text = Variables.editBox_altura_Text;
            }
        }
    }
}
