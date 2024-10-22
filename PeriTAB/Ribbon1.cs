using iTextSharp.text.pdf.security;
using iTextSharp.text.pdf;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Org.BouncyCastle.Crypto.Parameters;
using Org.BouncyCastle.Pkcs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Windows.Forms;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.X509;
using System.Security.Cryptography.X509Certificates;
using X509Certificate = Org.BouncyCastle.X509.X509Certificate;
using System.IdentityModel.Tokens;
using System.Runtime.Remoting.Channels;
using System.Diagnostics.Eventing.Reader;
using static System.Windows.Forms.LinkLabel;
using System.Net;
using System.Text;
using System.Net.Http;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using System.Drawing;
using System.Windows.Controls;
using System.Net.Security;
using Org.BouncyCastle.Crypto.Tls;
using System.Security.Authentication;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;
using ShapeRange = Microsoft.Office.Interop.Word.ShapeRange;
//using Spire.Doc.Interface;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Policy;
using Microsoft.VisualBasic.Devices;
using iTextSharp.text.pdf.codec.wmf;
using iTextSharp.xmp.impl.xpath;
using Microsoft.VisualBasic;
//using System.Windows;
//using Microsoft.VisualBasic;
using System.Text.RegularExpressions;


namespace PeriTAB
{
    public partial class Ribbon1
    {
        Class_Buttons iClass_Buttons = new Class_Buttons();

        public class Variables
        {
            private static string var1 = Path.GetTempPath() + "PeriTAB_Template_tmp.dotm";
            private static string var2 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PeriTAB");
            private static string var3, var4;
            //private static X509Certificate2 var_cert = null;
            //private static IExternalSignature var_sig = null;
            public static string caminho_template { get { return var1; } set { } }
            public static string caminho_AppData_Roaming_PeriTAB { get { return var2; } set { } }
            public static string editBox_largura_Text { get { return var3; } set { var3 = value; } }
            public static string editBox_altura_Text { get { return var4; } set { var4 = value; } }
            //public static X509Certificate2 cert { get { return var_cert; } set { var_cert = value; } }
            //public static IExternalSignature sig { get { return var_sig; } set { var_sig = value; } }
        }

        const string quote = "\"";
        const string slash = @"\";

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //MessageBox.Show("load");
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
            Globals.ThisAddIn.Application.AddIns.Add(Variables.caminho_template);

            // Escreve o número da versão
            //System.Version publish_version = Assembly.GetExecutingAssembly().GetName().Version;
            //Globals.Ribbons.Ribbon1.label_nome.Label = "PeriTAB " + publish_version.Major + "." + publish_version.Minor + "." + publish_version.Build;
            //Globals.Ribbons.Ribbon1.label_nome.Label = "PeriTAB " + publish_version.Major + "." + publish_version.Minor + "." + publish_version.Build;

            if (versao() != null)
            {
                Globals.Ribbons.Ribbon1.label_nome.Label = "PeriTAB " + versao().Major + "." + versao().Minor + "." + versao().Build;
            }
            else
            {
                Globals.Ribbons.Ribbon1.label_nome.Label = "PeriTAB Debugging";
            }
        }

        //public void Add_Button(object sender)
        //{
        //    Globals.Ribbons.Ribbon1.group_formatacao.Items.Add((RibbonControl)sender);
        //}

        public System.Version versao()
        {
            System.Version publish_version = null;
            if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
            {
                publish_version = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion;
            }
            //else
            //{
            //    publish_version = Assembly.GetExecutingAssembly().GetName().Version;
            //}
            return publish_version;
        }

        private void button_confere_num_legenda_Click(object sender, RibbonControlEventArgs e)
        {
            new Thread(() =>
            {
                iClass_Buttons.muda_imagem("button_confere_num_legenda", Properties.Resources.load_icon_png_7969);
                button_confere_num_legenda.Enabled = false;
                Globals.ThisAddIn.Application.Run("atualiza_todos_campos"); //****************
                Globals.ThisAddIn.Application.Run("confere_numeracao_legendas");
                iClass_Buttons.muda_imagem("button_confere_num_legenda", Properties.Resources.lupa);
                button_confere_num_legenda.Enabled = true;
            }).Start();
        }

        private void button_alinha_legenda_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("alinha_legenda");
        }

        private void button_atualiza_campos_Click(object sender, RibbonControlEventArgs e)
        {
            new Thread(() =>
            {
                iClass_Buttons.muda_imagem("button_atualiza_campos", Properties.Resources.load_icon_png_7969);
                button_atualiza_campos.Enabled = false;
                Globals.ThisAddIn.Application.Run("atualiza_todos_campos");
                Globals.ThisAddIn.Application.DisplayStatusBar = true; Globals.ThisAddIn.Application.StatusBar = "Campos atualizados com sucesso.";
                iClass_Buttons.muda_imagem("button_atualiza_campos", Properties.Resources.atualizar);
                button_atualiza_campos.Enabled = true;
            }).Start();
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

        private void toggleButton_painel_de_estilos_Click(object sender, RibbonControlEventArgs e)
        {
            //    var botao_toggle = (Microsoft.Office.Tools.Ribbon.RibbonToggleButton)sender;
            //    if (botao_toggle.Checked == true) Globals.ThisAddIn.TaskPane1.Visible = true;
            //    if (botao_toggle.Checked == false) Globals.ThisAddIn.TaskPane1.Visible = false;

            //foreach (MyUserControl UC1 in Class_New_or_Open_Event.list_UserControl)
            //{
            //    var botao_toggler = (Microsoft.Office.Tools.Ribbon.RibbonToggleButton)sender;
            //    if (botao_toggler.Checked == true) Class_New_or_Open_Event.Metodo_TaskPanes_Visible(true);
            //    if (botao_toggler.Checked == false) Class_New_or_Open_Event.Metodo_TaskPanes_Visible(false);
            //}
            //foreach (MyUserControl UC1 in Globals.ThisAddIn.Dicionario_Doc_e_UserControl.Values)
            //{
            var botao_toggle = (Microsoft.Office.Tools.Ribbon.RibbonToggleButton)sender;
            if (botao_toggle.Checked == true) Class_New_or_Open_Event.Metodo_TaskPanes_Visible(true);
            if (botao_toggle.Checked == false) Class_New_or_Open_Event.Metodo_TaskPanes_Visible(false);
            //}

        }

        private void button_inserir_sumario_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "Sumário 1", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "Sumário 2", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "Sumário 3", WdOrganizerObject.wdOrganizerObjectStyles);
            Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, "Sumário 4", WdOrganizerObject.wdOrganizerObjectStyles);
            //Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldTOC, slash + "h " + slash + "z " + slash + "t " + quote + "04 - Seções (PeriTAB);1" + quote, false);
            Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldTOC, slash + "h " + slash + "z " + slash + "t " + quote + "04A - SEÇÃO_1 (PERITAB);1;04B - SEÇÃO_2 (PERITAB);2;04C - SEÇÃO_3 (PERITAB);3;04D - SEÇÃO_4 (PERITAB);4" + quote, false);
            //Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldTOC, slash + "h " + slash + "z " + slash + "t " + quote + "04A - SEÇÃO_1 (PERITAB);1;04B - SEÇÃO_2 (PERITAB);2;04C - SEÇÃO_3 (PERITAB);3;04D - SEÇÃO_4 (PERITAB);4" + quote + " " + slash + "c " + quote + "Figura" + quote + " " + slash + "c " + quote + "Tabela" + quote, false);

        }

        private void button_inserir_pagina_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldEmpty, "PAGE", false);
        }

        private void button_inserir_pagina_extenso_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldPage, slash + "* Cardtext " + slash + "* Lower", false);
        }

        private void button_inserir_paginas_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldEmpty, "NUMPAGES", false);
        }
        private void button_inserir_paginas_extenso_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Selection.Fields.Add(Globals.ThisAddIn.Application.Selection.Range, WdFieldType.wdFieldNumPages, slash + "* Cardtext " + slash + "* Lower", false);
        }


        private void button_cola_imagem_Click(object sender, RibbonControlEventArgs e)
        {
            object obj = System.Windows.Clipboard.GetData("FileDrop");
            new Thread(() =>
            {
                // Configurações iniciais
                Stopwatch stopwatch = new Stopwatch(); if (versao() == null) { stopwatch.Start(); } // Inicia o cronômetro para medir o tempo de execução da Thread
                bool success = true;
                string msg_StatusBar = "";
                string msg_Falha = "";
                iClass_Buttons.muda_imagem("button_cola_imagem", Properties.Resources.load_icon_png_7969);
                button_cola_imagem.Enabled = false;
                Globals.ThisAddIn.Application.ScreenUpdating = false;

                if (System.Windows.Clipboard.ContainsData("FileDrop"))
                {
                    //object obj = System.Windows.Clipboard.GetData("FileDrop");
                    string[] pathfile = (string[])obj;
                    //for (int i = 0; i <= pathfile.Length - 1; i++) MessageBox.Show(pathfile[i]);
                    string[] pathfile2 = { "" };
                    string[] pathfile3 = { "" };
                    int n = 0;
                    for (int i = 0; i <= pathfile.Length - 1; i++)
                    {
                        if (File.Exists(pathfile[i]))
                        {
                            string extensao = (pathfile[i].Substring(pathfile[i].Length - 4)).ToLower();
                            if (extensao == ".jpg" | extensao == "jpeg" | extensao == ".png" | extensao == ".bmp" | extensao == ".gif" | extensao == "tiff") //Se tem extensao de imagem
                            {
                                Array.Resize(ref pathfile2, n + 1);
                                //Array.Resize(ref pathfile3, n + 1);
                                pathfile2[n] = pathfile[i];
                                //pathfile3[n] = pathfile[i];
                                n++;
                            }
                        }
                    }

                    if (pathfile2[0] != "")
                    {
                        //for (int i = 0; i <= pathfile2.Length - 1; i++) MessageBox.Show(pathfile2[i]);
                        //string[] pathfile3 = null;
                        //for (int i = 0; i <= pathfile2.Length - 1; i++) pathfile3[i] = pathfile2[i];
                        //if (testa_igualdade(pathfile2,pathfile3))
                        //{
                        //    MessageBox.Show("igual");
                        //}


                        //Array.Sort(pathfile3);

                        //if (dropDown_ordem.SelectedItem.Label == "Alfabética") {

                        //pathfile2.OrderBy(x => Convert.ToInt16(Path.GetFileNameWithoutExtension(x)));
                        //pathfile2.OrderBy(x => x);

                        Array.Sort(pathfile2, new Comparer_Windows_order());

                        //Array.Sort(pathfile2, StringComparer.Ordinal);
                        //DirectoryInfo[] di = new DirectoryInfo(pathfile2);
                        //FileSystemInfo[] files = di.GetFileSystemInfos();
                        //var orderedFiles = files.OrderBy(f => f.Name);
                        //pathfile2 = orderedFiles

                        //Array.Sort(pathfile2);
                        //pathfile2.OrderBy(System.IO.Path.GetFileNameWithoutExtension);
                        //Array.Sort(pathfile2, (s1, s2) => Path.GetFileName(s1).CompareTo(Path.GetFileName(s2)));
                        //pathfile2.OrderBy(f => f);
                        //pathfile2.OrderBy(System.IO.Path.GetFileName);
                        //pathfile2 = pathfile2.OrderBy(System.IO.Path.GetFileName).ToList();
                        //List<string> pathfile_list = new List<string> { };
                        //pathfile_list = pathfile2.OrderBy(System.IO.Path.GetFileName).ToList();
                        //pathfile2.OrderBy(x => x.Substring(0,x.LastIndexOf(".")));
                        //Array.Sort(pathfile2, (a,b) => ;

                        //string[] pathfile = (string[])obj;
                        //string[] pathfile2 = { "" };
                        //} 
                        //Ordem alfabética        


                        //if (dropDown_ordem.SelectedItem.Label == "Seleção")
                        //{
                        //    if (pathfile2.Length == 2)
                        //    {
                        //        string temp = pathfile2[0];
                        //        pathfile2[0] = pathfile2[1];
                        //        pathfile2[1] = temp;
                        //    }
                        //    else
                        //    {
                        //        MessageBox.Show("A opção ORDEM: SELEÇÃO só funciona para até 2 imagens.");
                        //        Globals.ThisAddIn.Application.ScreenUpdating = true;
                        //        iClass_Buttons.muda_imagem("button_cola_imagem", Properties.Resources.image_icon);
                        //        return;
                        //    }
                        //}
                        //    //Array.Sort(pathfile2);
                        //    for (int i = 0; i <= pathfile2.Length - 1; i++) MessageBox.Show(pathfile2[i]);
                        //    //for (int i = 0; i <= pathfile3.Length - 1; i++) MessageBox.Show(pathfile3[i]);

                        //    //if (!testa_igualdade(pathfile2, pathfile3))
                        //    //{
                        //    //MessageBox.Show("diferente");

                        //    string first = pathfile2[0];
                        //    for (int i = 0; i <= pathfile2.Length - 2; i++)
                        //    {
                        //        //if (i != pathfile2.Length - 1) 
                        //        //{
                        //        pathfile2[i] = pathfile2[i + 1];
                        //        //}
                        //        //pathfile2[pathfile2.Length - 1] = first;
                        //    }
                        //    pathfile2[pathfile2.Length - 1] = first;
                        //    //for (int i = 0; i <= pathfile2.Length - 1; i++) MessageBox.Show(pathfile2[i]);
                        //    //}
                        //    //else
                        //    //{
                        //    //    pathfile2 = pathfile3;
                        //    //}
                        //    //for (int i = 0; i <= pathfile2.Length - 1; i++) MessageBox.Show(pathfile2[i]);
                        //}
                        //for (int i = 0; i <= pathfile2.Length - 1; i++) MessageBox.Show(pathfile2[i]);

                        for (int i = 0; i <= pathfile2.Length - 1; i++)
                        {
                            //Globals.ThisAddIn.Application.ScreenUpdating = false;

                            bool link = false; bool save = true;
                            if (Globals.Ribbons.Ribbon1.checkBox_referencia.Checked == true) { link = true; save = false; }

                            InlineShape imagem = Globals.ThisAddIn.Application.Selection.InlineShapes.AddPicture(pathfile2[i], link, save);
                            imagem.LockAspectRatio = MsoTriState.msoTrue;
                            //MsoTriState LockAspectRatio_i = imagem.LockAspectRatio;
                            //imagem.LockAspectRatio = (MsoTriState)1;
                            if (checkBox_largura.Checked)
                            {
                                string larg_string = Globals.Ribbons.Ribbon1.editBox_largura.Text;
                                float.TryParse(larg_string, out float larg);
                                imagem.Width = Globals.ThisAddIn.Application.CentimetersToPoints(larg);
                            }

                            if (checkBox_altura.Checked)
                            {
                                string alt_string = Globals.Ribbons.Ribbon1.editBox_altura.Text;
                                float.TryParse(alt_string, out float alt);
                                imagem.Height = Globals.ThisAddIn.Application.CentimetersToPoints(alt);
                            }
                            //imagem.LockAspectRatio = LockAspectRatio_i;

                            if (i != pathfile2.Length - 1) //Exceto última imagem
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
                                    case "Parágrafo + 3pt":
                                        Globals.ThisAddIn.Application.Selection.ParagraphFormat.SpaceAfter = 3;
                                        Globals.ThisAddIn.Application.Selection.InsertAfter(System.Environment.NewLine);
                                        Globals.ThisAddIn.Application.Selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                                        break;
                                }
                            }
                        }
                        // Seleção das imagens ao final da colagem
                        if (dropDown_separador.SelectedItem.Label == "Nenhum")
                        {
                            int L = pathfile2.Length;
                            Globals.ThisAddIn.Application.Selection.MoveEnd(WdUnits.wdCharacter, -L);
                            Globals.ThisAddIn.Application.Selection.MoveRight(WdUnits.wdCharacter, L, WdMovementType.wdExtend);
                        }
                        else
                        {
                            int L = pathfile2.Length;
                            Globals.ThisAddIn.Application.Selection.MoveEnd(WdUnits.wdCharacter, -(2 * L - 1));
                            Globals.ThisAddIn.Application.Selection.MoveRight(WdUnits.wdCharacter, 2 * L - 1, WdMovementType.wdExtend);
                        }
                        //Globals.ThisAddIn.Application.ScreenUpdating = true;
                    }
                    else
                    {
                        success = false;
                        msg_Falha = "Não há imagens no Clipboard.";
                    }
                }
                else
                {
                    success = false;
                    msg_Falha = "Não há imagens no Clipboard.";
                }

                // Mensagens da Thread
                if (success) { msg_StatusBar = "Cola imagem: Sucesso"; } else { msg_StatusBar = "Cola imagem: Falha"; }
                if (versao() == null) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
                {
                    stopwatch.Stop();
                    msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                }
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
                if (!success && msg_Falha != "") MessageBox.Show(msg_Falha, "Cola imagem");

                // Configurações finais
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                iClass_Buttons.muda_imagem("button_cola_imagem", Properties.Resources.image_icon);
                button_cola_imagem.Enabled = true;
            }).Start();
        }

        public class Comparer_Windows_order : IComparer<string> /*implement an IComparer to get the same sort behavior as Windows Explorer*/
        {

            [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
            static extern int StrCmpLogicalW(String x, String y);

            public int Compare(string x, string y)
            {
                return StrCmpLogicalW(x, y);
            }
        }

        private void checkBox_largura_Click(object sender, RibbonControlEventArgs e)
        {
            if (Variables.editBox_largura_Text == null)
            {
                if (Class_Buttons.preferences.largura == "") { Class_Buttons.preferences.largura = "10"; }
                Variables.editBox_largura_Text = Class_Buttons.preferences.largura;
            }

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
            if (Variables.editBox_altura_Text == null)
            {
                if (Class_Buttons.preferences.altura == "") { Class_Buttons.preferences.altura = "10"; }
                Variables.editBox_altura_Text = Class_Buttons.preferences.altura;
            }

            if (checkBox_altura.Checked)
            {
                checkBox_largura.Checked = false;
                editBox_largura.Enabled = false;
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
            //if (checkBox_referencia.Checked)
            //{
            //    System.Windows.Forms.MessageBox.Show("Cuidado! Excluir/mover/renomear o arquivo da imagem causará perda de referência.","Referência");
            //}
        }


        private void editBox_largura_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (Variables.editBox_largura_Text == null) { Variables.editBox_largura_Text = Class_Buttons.preferences.largura; }

            if (float.TryParse(editBox_largura.Text, out float larg) & larg.ToString() == editBox_largura.Text & larg >= 0.1 & larg < 100)
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
            if (Variables.editBox_altura_Text == null) { Variables.editBox_altura_Text = Class_Buttons.preferences.altura; }

            if (float.TryParse(editBox_altura.Text, out float alt) & alt.ToString() == editBox_altura.Text & alt >= 0.1 & alt < 100)
            {
                Variables.editBox_altura_Text = editBox_altura.Text;
            }
            else
            {
                editBox_altura.Text = Variables.editBox_altura_Text;
            }
        }



        private void button_renomeia_documento_Click(object sender, RibbonControlEventArgs e)
        {
            new Thread(() =>
            {
                // Configurações iniciais
                Stopwatch stopwatch = new Stopwatch(); if (versao() == null) { stopwatch.Start(); } // Inicia o cronômetro para medir o tempo de execução da Thread
                bool success = true;
                string msg_StatusBar = "";
                string msg_Falha = "";
                iClass_Buttons.muda_imagem("button_renomeia_documento", Properties.Resources.load_icon_png_7969);
                button_cola_imagem.Enabled = false;
                //Globals.ThisAddIn.Application.ScreenUpdating = false;
                //Globals.ThisAddIn.Application.Run("renomeia_documento");
                //Globals.ThisAddIn.Application.DisplayStatusBar = true; Globals.ThisAddIn.Application.StatusBar = "Documento renomeado com sucesso.";

                string nome_doc_completo = Globals.ThisAddIn.Application.ActiveDocument.FullName;
                string caminho_doc = Globals.ThisAddIn.Application.ActiveDocument.Path;
                string nome_doc_antigo = Globals.ThisAddIn.Application.ActiveDocument.Name;
                string nome_doc = null;

                //MessageBox.Show(nome_doc_completo);
                //MessageBox.Show(caminho_doc);
                //MessageBox.Show(nome_doc_antigo);

                nome_doc_completo = GetLocalPath(nome_doc_completo);
                //if (nome_doc_completo.StartsWith("http"))
                //{
                //    MessageBox.Show("Este documento está armazenado na internet, o que impossibilita o uso dessa Macro. Caso esteja usando o Microsoft Onedrive, você pode resolver esse problema desmarcando a opção 'Usar os aplicativos do Office para sincronizar os arquivos do Office que eu abri', localizada na aba 'Office' nas configurações do Microsoft OneDrive.");
                //    return;
                //}

                //if (caminho_doc == "")
                //{
                //    MessageBox.Show("Documentos que ainda não foram salvos não podem ser renomeados.");
                //    return;
                //}
                if (versao() == null) { stopwatch.Stop(); }
                nome_doc = Microsoft.VisualBasic.Interaction.InputBox("Novo nome do documento:", "", nome_doc_antigo.Substring(0, nome_doc_antigo.LastIndexOf(".")));
                if (versao() == null) { stopwatch.Start(); }

                // Expressão regular para validar nome de arquivo no Windows
                string regex_Windows = @"^[^\\\/\:\*\?\""<>\|]+$";

                // Usa Regex.IsMatch para validar o nome do arquivo
                //bool nomeValido = 

                if (/*nome_doc == "" || */!Regex.IsMatch(nome_doc, regex_Windows) || string.IsNullOrWhiteSpace(nome_doc))
                {
                    //MessageBox.Show("ok");
                    //return;
                    success = false;
                    msg_Falha = "Nome inválido.";
                }
                else if (nome_doc == null || nome_doc == nome_doc_antigo.Substring(0, nome_doc_antigo.LastIndexOf("."))) { }
                else
                {
                    Globals.ThisAddIn.Application.ActiveDocument.SaveAs2(FileName: Path.Combine(caminho_doc, nome_doc + ".docx"), FileFormat: WdSaveFormat.wdFormatDocumentDefault);

                    try { File.Delete(nome_doc_completo); }
                    catch
                    {
                        success = false;
                        msg_Falha = "Falha ao deletar o documento antigo.";
                        //MessageBox.Show("Falha ao deletar o documento antigo.");
                        //MessageBox.Show("Falha ao deletar o documento antigo 1.");

                        //GC.Collect();
                        //GC.WaitForPendingFinalizers();

                        //try { File.Delete(nome_doc_completo); } catch { MessageBox.Show("Falha ao deletar o documento antigo 2."); }
                        //GC.Collect();
                        //GC.WaitForPendingFinalizers();
                        //try { File.Delete(nome_doc_completo); } catch { MessageBox.Show("Falha ao deletar o documento antigo 3."); }
                        //try { foreach (var process in Process.GetProcessesByName(nome_doc_completo)) { process.Kill(); }; } catch { MessageBox.Show("Falha ao deletar o documento antigo 4."); }
                        //try {  File.Delete(nome_doc_completo); } catch { MessageBox.Show("Falha ao deletar o documento antigo 5."); }
                        //System.Runtime.InteropServices.Marshal.FinalReleaseComObject((object)nome_doc_completo);
                        //try { File.Delete(nome_doc_completo); } catch { MessageBox.Show("Falha ao deletar o documento antigo 6."); }
                    }
                }

                // Mensagens da Thread
                if (success) { msg_StatusBar = "Renomeia documento: Sucesso"; } else { msg_StatusBar = "Renomeia documento: Falha"; }
                if (versao() == null) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
                {
                    stopwatch.Stop();
                    msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                }
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
                if (!success && msg_Falha != "") MessageBox.Show(msg_Falha, "Renomeia documento");

                // Configurações finais
                //Globals.ThisAddIn.Application.ScreenUpdating = true;
                iClass_Buttons.muda_imagem("button_renomeia_documento", Properties.Resources.abc);
                button_cola_imagem.Enabled = true;
            }).Start();
        }

        private void button_gerar_pdf_Click(object sender, RibbonControlEventArgs e)
        {
            new Thread(() =>
            {
                Globals.ThisAddIn.Application.DisplayStatusBar = false;
                // Configurações iniciais
                Stopwatch stopwatch = new Stopwatch(); if (versao() == null) { stopwatch.Start(); } // Inicia o cronômetro para medir o tempo de execução da Thread
                bool success = true;
                string msg_StatusBar = "";
                string msg_Falha = "";
                iClass_Buttons.muda_imagem("button_gera_pdf", Properties.Resources.load_icon_png_7969);
                button_gera_pdf.Enabled = false;

                //iClass_Buttons.button_gera_pdf_image(load: true);
                PdfReader inputPdf = null;
                bool inputPdf_open = false;
                string path = Globals.ThisAddIn.Application.ActiveDocument.FullName;
                string localpath = GetLocalPath(path);
                if (localpath == null) {
                    success = false;
                    msg_Falha = "Não foi possível gerar o PDF.";
                    goto saida;
                    //iClass_Buttons.muda_imagem("button_gera_pdf", Properties.Resources.icone_pdf2); MessageBox.Show("Não foi possível gerar o PDF."); button_gera_pdf.Enabled = true; return; 
                }
                string path_pdf = localpath.Substring(0, localpath.LastIndexOf(".")) + ".pdf";
                //Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(localpath.Substring(0, localpath.LastIndexOf(".")), WdExportFormat.wdExportFormatPDF, UseISO19005_1: true);

                //try { Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(localpath.Substring(0, localpath.LastIndexOf(".")), WdExportFormat.wdExportFormatPDF, UseISO19005_1: true); } catch (COMException ex) { MessageBox.Show("O PDF está aberto. Feche-o para gerar um novo PDF."); return; }
                Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(Path.Combine(Path.GetTempPath(), "tmp_pdf_PeriTAB"), WdExportFormat.wdExportFormatPDF, UseISO19005_1: true);

                //if (File.Exists(Path.Combine(Path.GetTempPath(), "tmp_pdf_PeriTAB.pdf")))
                //{
                //    File.Move(Path.Combine(Path.GetTempPath(), "tmp_pdf_PeriTAB.pdf"), path_pdf);
                //    return;
                //}
                //else { MessageBox.Show("Não foi possível gerar o PDF."); return; }

                if (Globals.Ribbons.Ribbon1.checkBox_assinar.Checked)
                {
                    string path_pdf_assinado = localpath.Substring(0, localpath.LastIndexOf(".")) + "_assinado.pdf";

                    X509Certificate2 certClient = null;
                    X509Store st = new X509Store(StoreName.My, StoreLocation.CurrentUser);
                    st.Open(OpenFlags.MaxAllowed);
                    IExternalSignature s = null;
                    //MessageBox.Show("1");
                    foreach (X509Certificate2 c in st.Certificates)
                    {
                        if (c.Verify() == false) { st.Remove(c); continue; } //Elimina certificado não validados
                        try { s = new X509Certificate2Signature(c, "SHA-256"); } catch { st.Remove(c); } //Elimina certificado que não se pode pegar a assinatura
                    }
                    //MessageBox.Show("2");
                    switch (st.Certificates.Count)
                    {
                        case 0:
                            //MessageBox.Show("Nenhum certificado válido encontrado.");
                            success = false;
                            msg_Falha = "Nenhum certificado válido encontrado.";
                            goto saida;
                        case 1:
                            certClient = st.Certificates[0];
                            break;
                        default:
                            if (versao() == null) { stopwatch.Stop(); }
                            X509Certificate2Collection collection = X509Certificate2UI.SelectFromCollection(st.Certificates, "Escolha o certificado:", "", X509SelectionFlag.SingleSelection);
                            if (versao() == null) { stopwatch.Start(); }
                            if (collection.Count > 0)
                            {
                                certClient = collection[0];
                            }
                            else
                            {
                                //MessageBox.Show("Nenhum certificado foi selecionado.");
                                success = false;
                                //msg_Falha = "Nenhum certificado foi selecionado.";
                                goto saida;
                            }
                            break;
                    }
                    //Variables.cert = certClient;
                    //st.Dispose();
                    st.Close();

                    //st.Remove(certClient);
                    //Debug.WriteLine("1");
                    //Get Cert Chain
                    IList<X509Certificate> chain = new List<X509Certificate>();

                    X509Chain x509Chain = new X509Chain();
                    //MessageBox.Show("3");
                    x509Chain.Build(certClient);



                    //new Thread(() =>
                    //{
                    //    x509Chain.Build(certClient);
                    //}).Start();



                    //System.Threading.Tasks.Task t = System.Threading.Tasks.Task.Factory.StartNew(() =>
                    //{
                    //    x509Chain.Build(certClient);
                    //});
                    //t.Wait();

                    //bool thread_finnished = false;
                    //new Thread(() =>
                    //{
                    //    x509Chain.Build(certClient);
                    //    thread_finnished = true;
                    //}).Start();

                    //while (true) 
                    //{
                    //    if (thread_finnished) break;
                    //}



                    //MessageBox.Show("4");
                    foreach (X509ChainElement x509ChainElement in x509Chain.ChainElements)
                    {
                        chain.Add(DotNetUtilities.FromX509Certificate(x509ChainElement.Certificate));
                    }

                    //Debug.WriteLine("2");
                    //PdfReader inputPdf = new PdfReader(path_pdf);
                    //PdfReader inputPdf = new PdfReader(Path.Combine(Path.GetTempPath(), "tmp_pdf_PeriTAB.pdf"));
                    inputPdf = new PdfReader(Path.Combine(Path.GetTempPath(), "tmp_pdf_PeriTAB.pdf"));
                    inputPdf_open = true;

                    FileStream signedPdf = null;
                    try { 
                        signedPdf = new FileStream(path_pdf_assinado, FileMode.Create); 
                    } 
                    catch (IOException) 
                    {
                        success = false;
                        msg_Falha = "O PDF está aberto. Feche-o para gerar um novo PDF.";
                        goto saida;
                        //iClass_Buttons.muda_imagem("button_gera_pdf", Properties.Resources.icone_pdf2); 
                        //MessageBox.Show("O PDF está aberto. Feche-o para gerar um novo PDF."); 
                        //button_gera_pdf.Enabled = true;
                    }




                    PdfStamper pdfStamper = PdfStamper.CreateSignature(inputPdf, signedPdf, '\0');

                    // Desativa a persistência da chave no CSP, garantindo que a senha seja solicitada sempre
                    //RSACryptoServiceProvider rsa = (RSACryptoServiceProvider)certClient.PrivateKey;
                    //rsa.PersistKeyInCsp = false; // Força a solicitação da senha

                    RSACryptoServiceProvider rsa2 = new RSACryptoServiceProvider();
                    rsa2.PersistKeyInCsp = true;

                    IExternalSignature externalSignature = new X509Certificate2Signature(certClient, "SHA-256");

                    PdfSignatureAppearance signatureAppearance = pdfStamper.SignatureAppearance;





                    //signatureAppearance.SignatureGraphic = Image.GetInstance(pathToSignatureImage);
                    //signatureAppearance.SetVisibleSignature(new iTextSharp.text.Rectangle(0, 00, 250, 150), inputPdf.NumberOfPages, "Signature");
                    //signatureAppearance.SignatureRenderingMode = PdfSignatureAppearance.RenderingMode.DESCRIPTION;


                    //RSACryptoServiceProvider rsa = new RSACryptoServiceProvider();
                    //CspParameters cspp = new CspParameters();
                    //cspp.KeyContainerName = rsa.CspKeyContainerInfo.KeyContainerName;
                    //cspp.ProviderName = rsa.CspKeyContainerInfo.ProviderName;
                    //cspp.ProviderType = rsa.CspKeyContainerInfo.ProviderType;
                    //cspp.Flags = CspProviderFlags.NoPrompt;
                    //RSACryptoServiceProvider rsa2 = new RSACryptoServiceProvider(cspp);
                    //rsa.PersistKeyInCsp = true;

                    //(new RSACryptoServiceProvider()).PersistKeyInCsp = true; //Define chave persistente. Só pede a senha da primeira vez.

                    //RSACryptoServiceProvider rsa = new RSACryptoServiceProvider();
                    //rsa.PersistKeyInCsp = false;

                    //(new RSACryptoServiceProvider()).PersistKeyInCsp = false;

                    //CspParameters cspp = new CspParameters();
                    //cspp.KeyContainerName = "MyKeyContainer";
                    //RSACryptoServiceProvider rsa = new RSACryptoServiceProvider(cspp);

                    //if (Globals.Ribbons.Ribbon1.checkBox_senha.Checked)
                    //{
                    //    //(new RSACryptoServiceProvider()).PersistKeyInCsp = true; //Define chave persistente. Só pede a senha da primeira vez.
                    //    rsa.PersistKeyInCsp = true;
                    //    //MessageBox.Show("1");
                    //    //rsa.Clear();
                    //}
                    //if (!Globals.Ribbons.Ribbon1.checkBox_senha.Checked)
                    //{
                    //    //(new RSACryptoServiceProvider()).PersistKeyInCsp = false;
                    //    rsa.PersistKeyInCsp = false;
                    //    //MessageBox.Show("2");
                    //    rsa.Clear();
                    //}


                    try 
                    { 
                        MakeSignature.SignDetached(signatureAppearance, externalSignature, chain, null, null, null, 0, CryptoStandard.CMS);
                        // Descarrega a chave da memória após a assinatura
                        //if (certClient != null)
                        //{
                        //    var rsa = certClient.GetRSAPrivateKey() as RSACryptoServiceProvider;
                        //    if (rsa != null)
                        //    {
                        //        rsa.PersistKeyInCsp = false; // Força a não persistência da chave
                        //        rsa.Clear(); // Libera o CSP, garantindo que a senha seja solicitada novamente
                        //    }
                        //}
                    }
                    //try { MakeSignature.SignDetached(pdfStamper.SignatureAppearance, new X509Certificate2Signature(certClient, "SHA-256"), chain, null, null, null, 0, CryptoStandard.CMS); }
                    catch (CryptographicException)
                    {
                        //Cancelamento da senha do token
                        signedPdf.Close();
                        File.Delete(path_pdf_assinado);
                        success = false;
                        goto saida;
                    }
                    //****************************************************
                    //finally
                    //{
                    //    // Aqui liberamos o contexto da chave
                    //    if (certClient != null)
                    //    {
                    //        var rsa1 = certClient.GetRSAPrivateKey() as RSACryptoServiceProvider;
                    //        if (rsa1 != null)
                    //        {
                    //            rsa1.PersistKeyInCsp = false; // Garante que a chave não será persistida
                    //            rsa1.Clear(); // Libera o CSP, garantindo que a senha seja solicitada novamente
                    //        }
                    //    }
                    //}
                    //******************************************************
                    //inputPdf.Close();
                    //chain.Clear();
                    pdfStamper.Close();
                    if (File.Exists(path_pdf_assinado))
                    {
                        //iClass_Buttons.muda_imagem("button_gera_pdf", Properties.Resources.icone_pdf2);
                        //button_gera_pdf.Enabled = true;
                        //Globals.ThisAddIn.Application.DisplayStatusBar = true;
                        //Globals.ThisAddIn.Application.StatusBar = "PDF gerado com sucesso.";
                        //if (File.Exists(path_pdf)) { File.Delete(path_pdf); }
                        if (Globals.Ribbons.Ribbon1.checkBox_abrir.Checked) { System.Diagnostics.Process.Start(path_pdf_assinado); }
                    }
                    else
                    {
                        //iClass_Buttons.muda_imagem("button_gera_pdf", Properties.Resources.icone_pdf2);
                        //button_gera_pdf.Enabled = true;
                        //Globals.ThisAddIn.Application.DisplayStatusBar = true; Globals.ThisAddIn.Application.StatusBar = "A geração do PDF falhou.";
                        success = false;
                    }
                }
                else
                {
                    if (File.Exists(Path.Combine(Path.GetTempPath(), "tmp_pdf_PeriTAB.pdf")))
                    {
                        if (File.Exists(path_pdf)) { try { File.Delete(path_pdf); } catch (IOException) 
                            {
                                success = false;
                                msg_Falha = "O PDF está aberto. Feche-o para gerar um novo PDF.";
                                //MessageBox.Show("O PDF está aberto. Feche-o para gerar um novo PDF."); 
                                goto saida; 
                            } 
                        }
                        File.Move(Path.Combine(Path.GetTempPath(), "tmp_pdf_PeriTAB.pdf"), path_pdf);
                        //iClass_Buttons.muda_imagem("button_gera_pdf", Properties.Resources.icone_pdf2);
                        //button_gera_pdf.Enabled = true;
                        //Globals.ThisAddIn.Application.DisplayStatusBar = true; Globals.ThisAddIn.Application.StatusBar = "PDF gerado com sucesso.";
                        if (Globals.Ribbons.Ribbon1.checkBox_abrir.Checked) { System.Diagnostics.Process.Start(path_pdf); }
                    }
                    else
                    {
                        success = false;
                        msg_Falha = "Não foi possível gerar o PDF.";
                        goto saida;
                        //iClass_Buttons.muda_imagem("button_gera_pdf", Properties.Resources.icone_pdf2);
                        //button_gera_pdf.Enabled = true;
                        //MessageBox.Show("Não foi possível gerar o PDF.");
                    }
                }

            saida:
                if (inputPdf_open) inputPdf.Close();
                if (File.Exists(Path.Combine(Path.GetTempPath(), "tmp_pdf_PeriTAB.pdf"))) { File.Delete(Path.Combine(Path.GetTempPath(), "tmp_pdf_PeriTAB.pdf")); } //Deleta tmp.pdf

                // Mensagens da Thread
                if (success) { msg_StatusBar = "Gera PDF: Sucesso"; } else { msg_StatusBar = "Gera PDF: Falha"; }
                if (versao() == null) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
                {
                    stopwatch.Stop();
                    msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                }
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
                if (!success && msg_Falha != "") MessageBox.Show(msg_Falha, "Gera PDF");

                // Configurações finais
                //Globals.ThisAddIn.Application.ScreenUpdating = true;
                iClass_Buttons.muda_imagem("button_gera_pdf", Properties.Resources.icone_pdf2);
                button_gera_pdf.Enabled = true;
            }).Start();
        }

        public string GetLocalPath(string path)
        {
            string localpath;
            if ((path.Substring(0, 4)).ToLower() != "http") //Verifica se está armazenado online
            {
                localpath = path;
            }
            else
            {
                if ((path.ToLower()).IndexOf("sharepoint") != -1) //Verifica se está no OneDrive
                {
                    string path_onedrive = null;
                    string[] onedrive_EV = { Environment.GetEnvironmentVariable("OneDrive"), Environment.GetEnvironmentVariable("OneDriveConsumer"), Environment.GetEnvironmentVariable("OneDriveCommercial") }; // Variáveis de ambiente do onedrive no Windows
                    for (int i = 0; i <= onedrive_EV.Length - 1; i++)
                    {
                        if (path_onedrive == null) { path_onedrive = onedrive_EV[i]; }
                    }
                    string onedrive_subpasta;
                    if (path.IndexOf("/Documents/") != -1)
                    {
                        onedrive_subpasta = path.Substring(path.IndexOf("/Documents/") + ("/Documents/").Length, path.Length - (path.IndexOf("/Documents/") + ("/Documents/").Length));
                    }
                    else { return null; }
                    localpath = path_onedrive + slash + onedrive_subpasta.Replace("/", slash);
                }
                else
                {
                    return null;
                }
            }
            return localpath;
        }

        private void button_redimensiona_imagem_Click(object sender, RibbonControlEventArgs e)
        {
            new Thread(() =>
            {
                // Configurações iniciais
                Stopwatch stopwatch = new Stopwatch(); if (versao() == null) { stopwatch.Start(); } // Inicia o cronômetro para medir o tempo de execução da Thread
                bool success = true;
                string msg_StatusBar = "";
                string msg_Falha = "";
                iClass_Buttons.muda_imagem("button_redimensiona_imagem", Properties.Resources.load_icon_png_7969);
                button_redimensiona_imagem.Enabled = false;
                Globals.ThisAddIn.Application.ScreenUpdating = false;

                if (Globals.ThisAddIn.Application.Selection.InlineShapes.Count < 1)
                {
                    success = false;
                    msg_Falha = "Não há imagens selecionadas.";
                }

                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        InlineShape imagem = ishape;
                        imagem.LockAspectRatio = MsoTriState.msoTrue;

                        //MsoTriState LockAspectRatio_i = imagem.LockAspectRatio;
                        //imagem.LockAspectRatio = (MsoTriState)1;
                        if (checkBox_largura.Checked)
                        {
                            string larg_string = Globals.Ribbons.Ribbon1.editBox_largura.Text;
                            float.TryParse(larg_string, out float larg);
                            imagem.Width = Globals.ThisAddIn.Application.CentimetersToPoints(larg);
                        }

                        if (checkBox_altura.Checked)
                        {
                            string alt_string = Globals.Ribbons.Ribbon1.editBox_altura.Text;
                            float.TryParse(alt_string, out float alt);
                            imagem.Height = Globals.ThisAddIn.Application.CentimetersToPoints(alt);
                        }
                        //imagem.LockAspectRatio = LockAspectRatio_i;
                    }
                }

                // Mensagens da Thread
                if (success) { msg_StatusBar = "Redimensiona: Sucesso"; } else { msg_StatusBar = "Redimensiona: Falha"; }
                if (versao() == null) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
                {
                    stopwatch.Stop();
                    msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                }
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
                if (!success && msg_Falha != "") MessageBox.Show(msg_Falha, "Redimensiona");

                // Configurações finais
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                iClass_Buttons.muda_imagem("button_redimensiona_imagem", Properties.Resources.redimensionar2);
                button_redimensiona_imagem.Enabled = true;
            }).Start();
        }

        private void button_autodimensiona_imagem_Click(object sender, RibbonControlEventArgs e)
        {
            new Thread(() =>
            {
                // Configurações iniciais
                Stopwatch stopwatch = new Stopwatch(); if (versao() == null) { stopwatch.Start(); } // Inicia o cronômetro para medir o tempo de execução da Thread
                bool success = true;
                string msg_StatusBar = "";
                string msg_Falha = "";
                iClass_Buttons.muda_imagem("button_autodimensiona_imagem", Properties.Resources.load_icon_png_7969);
                button_autodimensiona_imagem.Enabled = false;
                Globals.ThisAddIn.Application.ScreenUpdating = false;

                if (Globals.ThisAddIn.Application.Selection.InlineShapes.Count < 1)
                {
                    success = false;
                    msg_Falha = "Não há imagens selecionadas.";
                }

                Dictionary<int, List<InlineShape>> dict_InlineShape_paragraph = new Dictionary<int, List<InlineShape>>();
                foreach (InlineShape iShape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    // Verifica se o parágrafo contém mais de uma InlineShape
                    if (iShape.Range.Paragraphs[1].Range.InlineShapes.Count > 1)
                    {
                        int num_Paragraph = 0;
                        if (iShape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | iShape.Type == WdInlineShapeType.wdInlineShapePicture)
                        {
                            Paragraph iParagraph = iShape.Range.Paragraphs.First;
                            for (int i = 1; i <= Globals.ThisAddIn.Application.ActiveDocument.Paragraphs.Count; i++)
                            {
                                if (Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[i].Range.Start == iParagraph.Range.Start)
                                {
                                    num_Paragraph = i;
                                    break;
                                }
                            }
                        }
                        // Verifica se o dicionário já contém o parágrafo
                        if (!dict_InlineShape_paragraph.ContainsKey(num_Paragraph))
                        {
                            // Se não contém, cria uma nova lista de InlineShapes para esse parágrafo
                            dict_InlineShape_paragraph[num_Paragraph] = new List<InlineShape>();
                        }
                        // Adiciona a InlineShape à lista correspondente ao parágrafo
                        dict_InlineShape_paragraph[num_Paragraph].Add(iShape);
                    }
                    else
                    {
                        float larguraPaginaPts = Globals.ThisAddIn.Application.ActiveDocument.PageSetup.PageWidth;
                        float margemEsquerdaPts = Globals.ThisAddIn.Application.ActiveDocument.PageSetup.LeftMargin;
                        float margemDireitaPts = Globals.ThisAddIn.Application.ActiveDocument.PageSetup.RightMargin;
                        float recuoEsquerdaPts = iShape.Range.Paragraphs[1].Format.LeftIndent;
                        float recuoDireitaPts = iShape.Range.Paragraphs[1].Format.RightIndent;
                        float primeiralinhaPts = iShape.Range.Paragraphs[1].Format.FirstLineIndent;
                        float espacoDigitavelPts = larguraPaginaPts - (margemEsquerdaPts + margemDireitaPts + recuoEsquerdaPts + recuoDireitaPts + primeiralinhaPts);
                        iShape.Width = espacoDigitavelPts;
                    }
                }
                // Itera por cada parágrafo que contém múltiplas InlineShapes
                foreach (var iParagraph in dict_InlineShape_paragraph.Keys)
                {
                    // Verifica se o parágrafo tem exatamente uma linha: caso de aumento das imagens
                    if (((dict_InlineShape_paragraph[iParagraph])[0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines)) == 1)
                    {
                        while (((dict_InlineShape_paragraph[iParagraph])[0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines)) == 1)
                        {
                            foreach (InlineShape iShape in dict_InlineShape_paragraph[iParagraph])
                            {
                                //iShape.Width = (float)(iShape.Width * 1.1);
                                iShape.Width *= 1.1f;
                            }
                        }
                        foreach (InlineShape iShape in dict_InlineShape_paragraph[iParagraph])
                        {
                            //iShape.Width = (float)(iShape.Width * 0.9);
                            iShape.Width *= 0.9f;
                        }
                        while (((dict_InlineShape_paragraph[iParagraph])[0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines)) == 1)
                        {
                            foreach (InlineShape iShape in dict_InlineShape_paragraph[iParagraph])
                            {
                                //iShape.Width = (float)(iShape.Width * 1.01);
                                iShape.Width *= 1.01f;
                            }
                        }
                        while (((dict_InlineShape_paragraph[iParagraph])[0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines)) == 2)
                        {
                            foreach (InlineShape iShape in dict_InlineShape_paragraph[iParagraph])
                            {
                                //iShape.Width = (float)(iShape.Width * 0.99);
                                iShape.Width *= 0.99f;
                            }
                        }
                    }
                    // Verifica se o parágrafo tem mais de uma linha: caso de redução das imagens
                    if (((dict_InlineShape_paragraph[iParagraph])[0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines)) > 1)
                    {
                        // Dicionário para armazenar os tamanhos originais das imagens
                        Dictionary<InlineShape, float> tamanho_original = new Dictionary<InlineShape, float>();

                        // Armazena o tamanho original das imagens no parágrafo
                        foreach (InlineShape iShape in dict_InlineShape_paragraph[iParagraph])
                        {
                            tamanho_original[iShape] = iShape.Width;
                        }

                        // Primeira tentativa de ajustar todas as imagens
                        for (int iteration = 0; iteration < 50; iteration++)
                        {
                            // Verifica se as imagens já cabem em uma linha
                            if (((dict_InlineShape_paragraph[iParagraph])[0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines)) <= 1)
                            {
                                break; // Sai do loop se já estiver ajustado
                            }

                            // Reduz todas as imagens no parágrafo por 10% a cada iteração
                            foreach (InlineShape iShape in dict_InlineShape_paragraph[iParagraph])
                            {
                                iShape.Width *= 0.9f;
                            }
                        }

                        // Após as 50 iterações, verifica se alguma imagem ficou menor que 1 cm e se ainda ocupa mais de uma linha
                        bool algumMenorQue1cm = false;

                        foreach (InlineShape iShape in dict_InlineShape_paragraph[iParagraph])
                        {
                            if (iShape.Width < 28.35) // Verifica se a largura é menor que 1 cm em pontos
                            {
                                algumMenorQue1cm = true;
                                break; // Não precisa verificar mais
                            }
                        }

                        // Se depois de 50 tentativas ainda não couber em uma linha ou imagem ficar muito pequena, desiste de redimensionar.
                        if (((dict_InlineShape_paragraph[iParagraph])[0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines)) > 1 || algumMenorQue1cm)
                        {
                            success = false;
                            msg_Falha = "Alguma(s) imagem(ns) selecionada(s) não cabe(m) em uma única linha.";

                            // Restaura os tamanhos originais das imagens
                            foreach (InlineShape iShape in dict_InlineShape_paragraph[iParagraph])
                            {
                                iShape.Width = tamanho_original[iShape];
                            }
                        }
                        else
                        {
                            // Se as imagens couberem em uma linha, faz o ajuste fino
                            foreach (InlineShape iShape in dict_InlineShape_paragraph[iParagraph])
                            {
                                iShape.Width *= 1.1f; // Aumenta ligeiramente o tamanho

                                // Faz um ajuste final, caso precise
                                while (((dict_InlineShape_paragraph[iParagraph])[0].Range.Paragraphs[1].Range.ComputeStatistics(WdStatistic.wdStatisticLines)) > 1)
                                {
                                    iShape.Width *= 0.99f; // Ajusta em decrementos menores (1%)
                                }
                            }
                        }
                    }
                }

                // Mensagens da Thread
                if (success) { msg_StatusBar = "Autodimensiona: Sucesso"; } else { msg_StatusBar = "Autodimensiona: Falha"; }
                if (versao() == null) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
                {
                    stopwatch.Stop();
                    msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                }
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
                if (!success && msg_Falha != "") MessageBox.Show(msg_Falha, "Autodimensiona");

                // Configurações finais
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                iClass_Buttons.muda_imagem("button_autodimensiona_imagem", Properties.Resources.redimensionar3);
                button_autodimensiona_imagem.Enabled = true;
            }).Start();
        }

        private string get_text(string texto, string inicio = null, string fim = null) //Retona a primeira ocorrência de string entre os strings 'inicio' e 'fim' no string 'texto'.
        {
            if (inicio == null & fim == null) { return null; }

            try
            {
                if (inicio == null)
                {
                    return texto.Substring(0, texto.IndexOf(fim));
                }
                if (fim == null)
                {
                    return texto.Substring(texto.IndexOf(inicio) + inicio.Length);
                }
                //return (texto.Substring(texto.IndexOf(inicio))).Substring(inicio.Length, (texto.Substring(texto.IndexOf(inicio))).IndexOf(fim) - inicio.Length);
                return (texto.Substring(texto.IndexOf(inicio))).Substring(inicio.Length, (texto.Substring(texto.IndexOf(inicio) + inicio.Length)).IndexOf(fim));
            }
            catch { return null; }
        }

        //private void button_remove_formatacao_Click(object sender, RibbonControlEventArgs e)
        //{
        //    Globals.ThisAddIn.Application.ScreenUpdating = false;
        //    foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
        //    {
        //        if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
        //        {
        //            //ishape.Line.Visible = MsoTriState.msoFalse;
        //            ishape.Reset();
        //            ishape.AlternativeText = "";
        //        }
        //    }
        //    Globals.ThisAddIn.Application.ScreenUpdating = true;
        //}

        //private void button_insere_borda_preta_Click(object sender, RibbonControlEventArgs e)
        //{
        //    Globals.ThisAddIn.Application.ScreenUpdating = false;
        //    foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
        //    {
        //        if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
        //        {
        //            ishape.Line.Visible = MsoTriState.msoTrue;
        //            ishape.Line.Weight = (float)0.5;
        //            ishape.Line.ForeColor.RGB = Color.FromArgb(0, 0, 0).ToArgb();
        //        }
        //    }
        //    Globals.ThisAddIn.Application.ScreenUpdating = true;
        //}

        //private void button_insere_borda_vermelha_Click(object sender, RibbonControlEventArgs e)
        //{
        //    Globals.ThisAddIn.Application.ScreenUpdating = false;
        //    foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
        //    {
        //        if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
        //        {
        //            ishape.Line.Visible = MsoTriState.msoTrue;
        //            ishape.Line.Weight = 2;
        //            ishape.Line.ForeColor.RGB = Color.FromArgb(0, 0, 255).ToArgb();
        //        }
        //    }
        //    Globals.ThisAddIn.Application.ScreenUpdating = true;
        //}

        //private void button_insere_borda_amarela_Click(object sender, RibbonControlEventArgs e)
        //{
        //    Globals.ThisAddIn.Application.ScreenUpdating = false;
        //    foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
        //    {
        //        if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
        //        {
        //            ishape.Line.Visible= MsoTriState.msoTrue;
        //            ishape.Line.Weight = 3;
        //            ishape.Line.ForeColor.RGB = Color.FromArgb(0,255,255).ToArgb();
        //        }
        //    }
        //    Globals.ThisAddIn.Application.ScreenUpdating = true;
        //}



        private bool IsLastShapeInParagraph(InlineShape ishape)
        {
            InlineShape lastShape = null;
            foreach (InlineShape inlineShape in ishape.Range.Paragraphs[1].Range.InlineShapes)
            {
                if (inlineShape.Type == WdInlineShapeType.wdInlineShapePicture)
                {
                    lastShape = inlineShape;
                }
            }
            if (lastShape != null && ishape.Range.Information[WdInformation.wdHorizontalPositionRelativeToTextBoundary] >= lastShape.Range.Information[WdInformation.wdHorizontalPositionRelativeToTextBoundary] && ishape.Range.Information[WdInformation.wdVerticalPositionRelativeToTextBoundary] >= lastShape.Range.Information[WdInformation.wdVerticalPositionRelativeToTextBoundary])
            {
                return true;
            }
            return false;
        }

        private void button_confere_preambulo_Click(object sender, RibbonControlEventArgs e)
        {
            new Thread(() =>
            {
                iClass_Buttons.muda_imagem("button_confere_preambulo", Properties.Resources.load_icon_png_7969);
                button_confere_preambulo.Enabled = false;


                string localpath = GetLocalPath(Globals.ThisAddIn.Application.ActiveDocument.Path);
                string download_path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");

                string[] identificadores_laudo = pega_identificadores_laudo();

                string num_laudo = identificadores_laudo[0];
                string ano_laudo = identificadores_laudo[1];
                string unidade_laudo = identificadores_laudo[2];

                if (num_laudo == null | ano_laudo == null | unidade_laudo == null)
                {
                    MessageBox.Show("Referência do laudo não encontrada.");
                    iClass_Buttons.muda_imagem("button_confere_preambulo", Properties.Resources.checklist2);
                    button_confere_preambulo.Enabled = true;
                    return;
                }

                //for (int i = 1; i <= Globals.ThisAddIn.Application.ActiveDocument.Paragraphs.Count; i++)
                //{
                //    string t = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[i].Range.Text;
                //    string t_mod = t.ToLower().Replace(" ", "").Replace(((char)160).ToString(), "").Replace(((char)9).ToString(), "").Replace(((char)8211).ToString(), "-").Replace(((char)176).ToString(), "*").Replace("º", "*").Replace("laudono", "laudon*"); //elimina espaços, espaços inquebráveis e tabs. Ainda troca en-dash por hifen e grau por 'o' sobrescrito.
                //                                                                                                                                                                                                                                                  //MessageBox.Show(t_mod);
                //                                                                                                                                                                                                                                                  //if (t_mod == ((char)13).ToString()) { continue; } //

                //    //string result = "";
                //    //foreach (char c in t_trim) { result += (int)c + " "; }
                //    //MessageBox.Show(result);

                //    if (t_mod.Length > 10)
                //    {
                //        if ((t_mod.Substring(0, 6)).ToLower() == "laudon")
                //        {
                //            num_laudo = get_text(t_mod, "n*", "/");
                //            ano_laudo = get_text(t_mod, "/", "-");
                //            unidade_laudo = get_text(t_mod, "-");
                //            break;
                //            //try { unidade_laudo = t_trim.ToLower().Substring(t_trim.ToLower().IndexOf("- ") + 2); } catch { unidade_laudo = null; }
                //        }
                //    }
                //}
                ////MessageBox.Show(num_laudo + " " + ano_laudo + " " + unidade_laudo);
                //if (num_laudo == null | ano_laudo == null | unidade_laudo == null) { MessageBox.Show("Referência do laudo não encontrada."); return; }

                string asap_path = Path.Combine(localpath, "AsAP_Laudo_" + num_laudo + "-" + ano_laudo + ".asap");
                string asap_downloads_path = Path.Combine(download_path, "AsAP_Laudo_" + num_laudo + "-" + ano_laudo + ".asap");

                // Move o arquivo ASAP de downloads.
                if (File.Exists(asap_downloads_path) & !File.Exists(asap_path))
                {
                    File.Move(asap_downloads_path, asap_path);
                }

                if (File.Exists(asap_path))
                {
                    string ASAP = File.ReadAllText(asap_path, Encoding.Default);
                    //string preambulo = pega_preambulo_laudo();
                    //if (preambulo == null) { MessageBox.Show("preambulo não encontrado."); return; }
                    int paragrafo_do_preambulo = pega_paragrafo_do_preambulo();
                    if (paragrafo_do_preambulo == 0) { MessageBox.Show("preambulo não encontrado."); return; }
                    //string preambulo = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[paragrafo_do_preambulo].Range.Text;
                    string preambulo_padrao = faz_preambulo_padrao(ASAP);
                    //MessageBox.Show(preambulo);
                    compara_preambulo(preambulo_padrao, paragrafo_do_preambulo);
                }
                else
                {
                    DialogResult resultado = MessageBox.Show("Arquivo ASAP não encontrado. Gostaria de baixá-lo?" + System.Environment.NewLine + "(Certificado/Token é necessário)", "", MessageBoxButtons.YesNo);
                    if (resultado == System.Windows.Forms.DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start("https://www.ditec.pf.gov.br:8443/sistemas/criminalistica/controle_documento.php?action=localizar_resultado&d-codigo_tipo_documento=2704&d-numero_documento=" + num_laudo + "&d-ano_documento=" + ano_laudo + "&d-sigla_orgao_emissor-ilike=" + unidade_laudo + "&codigo_unidade_registro_pesquisa=");
                        //System.Diagnostics.Process.Start("https://www.ditec.pf.gov.br/sistemas/criminalistica/documento.php?acao=localizar_registro&tipo_busca=numero_laudo&numero_busca=" + num_laudo + "/" + ano_laudo);
                    }
                }
                iClass_Buttons.muda_imagem("button_confere_preambulo", Properties.Resources.checklist2);
                button_confere_preambulo.Enabled = true;
            }).Start();
        }

        private string[] pega_identificadores_laudo()
        {
            string num_laudo = null;
            string ano_laudo = null;
            string unidade_laudo = null;

            for (int i = 1; i <= Globals.ThisAddIn.Application.ActiveDocument.Paragraphs.Count; i++)
            {
                string t = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[i].Range.Text;
                string t_mod = t.ToLower().Replace(" ", "").Replace(((char)160).ToString(), "").Replace(((char)9).ToString(), "").Replace(((char)8211).ToString(), "-").Replace(((char)176).ToString(), "*").Replace("º", "*").Replace("laudono", "laudon*"); //elimina espaços, espaços inquebráveis e tabs. Ainda troca en-dash por hifen e grau por 'o' sobrescrito.
                                                                                                                                                                                                                                                              //MessageBox.Show(t_mod);
                if (t_mod.Length > 10)
                {
                    if ((t_mod.Substring(0, 6)).ToLower() == "laudon")
                    {
                        num_laudo = get_text(t_mod, "n*", "/");
                        ano_laudo = get_text(t_mod, "/", "-");
                        unidade_laudo = get_text(t_mod, "-");
                        return new string[] { num_laudo, ano_laudo, unidade_laudo };
                    }
                }
            }
            //if (num_laudo == null | ano_laudo == null | unidade_laudo == null) { return null; }
            //else return null;
            return new string[] { null, null, null }; ;
        }

        private int pega_paragrafo_do_preambulo()
        {
            int paragrafo_do_preambulo = 0;
            for (int i = 1; i <= Globals.ThisAddIn.Application.ActiveDocument.Paragraphs.Count; i++)
            {
                string t = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[i].Range.Text;
                string t_trim = t.Trim();
                if (t_trim.Length > 200)
                {
                    if ((t_trim.Substring(0, 2)).ToLower() == "em")
                    {
                        paragrafo_do_preambulo = i;
                        break;
                    }
                }
            }
            return paragrafo_do_preambulo;
        }

        //private string pega_preambulo_laudo()
        //{
        //    string preambulo = null;
        //    for (int i = 1; i <= Globals.ThisAddIn.Application.ActiveDocument.Paragraphs.Count; i++)
        //    {
        //        string t = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[i].Range.Text;
        //        string t_trim = t.Trim();
        //        if (t_trim.Length > 200)
        //        {
        //            if ((t_trim.Substring(0, 2)).ToLower() == "em")
        //            {
        //                preambulo = t;
        //                break;
        //            }
        //        }
        //    }
        //    return preambulo;
        //}


        //string url1 = "https://www.ditec.pf.gov.br:8443/sistemas/criminalistica/controle_documento.php?action=localizar_resultado&d-codigo_tipo_documento=2704&d-numero_documento=" + num_laudo + "&d-ano_documento=" + ano_laudo + "&d-sigla_orgao_emissor-ilike=" + unidade_laudo + "&codigo_unidade_registro_pesquisa=";
        ////url1 = "http://google.com";

        //// Create a web request that points to our SSL-enabled client certificate required web site
        //HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(url1);
        //ServicePointManager.Expect100Continue = true;
        //ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        //ServicePointManager.ServerCertificateValidationCallback += new RemoteCertificateValidationCallback(AlwaysGoodCertificate);

        //// Use the X509Store class to get a handle to the local certificate stores. "My" is the "Personal" store.
        //X509Store store = new X509Store(StoreName.My, StoreLocation.LocalMachine);

        //// Open the store to be able to read from it.
        //store.Open(OpenFlags.ReadOnly);

        //// Use the X509Certificate2Collection class to get a list of certificates that match our criteria (in this case, we should only pull back one).
        //X509Certificate2Collection collection = store.Certificates.Find(X509FindType.FindBySubjectName, "MyClientCert", true);

        //// Associate the certificates with the request
        //request.ClientCertificates = collection;

        //// Make the web request
        //HttpWebResponse response = (HttpWebResponse)request.GetResponse();

        //// Output the stream to a file.
        ////Stream stream = response.GetResponseStream();
        //string resp = "";
        //if (response.StatusCode == HttpStatusCode.OK)
        //{
        //    Stream receiveStream = response.GetResponseStream();
        //    StreamReader readStream = null;
        //    if (response.CharacterSet == null)
        //        readStream = new StreamReader(receiveStream);
        //    else
        //        readStream = new StreamReader(receiveStream, Encoding.GetEncoding(response.CharacterSet));
        //    resp = readStream.ReadToEnd();
        //    MessageBox.Show(resp);
        //}

        //string url1 = "https://www.ditec.pf.gov.br:8443/sistemas/criminalistica/controle_documento.php?action=localizar_resultado&d-codigo_tipo_documento=2704&d-numero_documento=" + num_laudo + "&d-ano_documento=" + ano_laudo + "&d-sigla_orgao_emissor-ilike=" + unidade_laudo + "&codigo_unidade_registro_pesquisa=";
        ////url1 = "http://google.com";

        //ServicePointManager.Expect100Continue = true;
        //ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;



        //X509Store st = new X509Store(StoreName.My, StoreLocation.CurrentUser);
        //st.Open(OpenFlags.MaxAllowed);
        //X509Certificate2Collection collection = X509Certificate2UI.SelectFromCollection(st.Certificates, "Escolha o certificado:", "", X509SelectionFlag.SingleSelection);

        ////Get Cert Chain

        //X509Certificate2 certClient = collection[0];
        //IList<X509Certificate> chain = new List<X509Certificate>();
        //X509Chain x509Chain = new X509Chain();
        //x509Chain.Build(certClient);
        //foreach (X509ChainElement x509ChainElement in x509Chain.ChainElements)
        //{
        //    chain.Add(DotNetUtilities.FromX509Certificate(x509ChainElement.Certificate));
        //}



        ////ServicePointManager.ServerCertificateValidationCallback += new RemoteCertificateValidationCallback(ValidateRemoteCertificate);

        //HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url1);
        //request.ClientCertificates = collection;


        ////HttpWebResponse response = (HttpWebResponse)request.GetResponse();
        ////string resp = "";
        ////if (response.StatusCode == HttpStatusCode.OK)
        ////{
        ////    Stream receiveStream = response.GetResponseStream();
        ////    StreamReader readStream = null;
        ////    if (response.CharacterSet == null)
        ////        readStream = new StreamReader(receiveStream);
        ////    else
        ////        readStream = new StreamReader(receiveStream, Encoding.GetEncoding(response.CharacterSet));
        ////    resp = readStream.ReadToEnd();
        ////    response.Close();
        ////    readStream.Close();
        ////}

        //HttpClient httpClient = new HttpClient();
        //HttpResponseMessage result = httpClient.GetAsync(url1).Result;
        //string str = result.Content.ReadAsStringAsync().Result;
        //MessageBox.Show(str);

        ////MessageBox.Show(resp);


        //      private static bool ValidateRemoteCertificate(
        //object sender,
        //X509Certificate certificate,
        //X509Chain chain,
        //SslPolicyErrors policyErrors)
        //      {
        //          return true;
        //      }

        //private static bool AlwaysGoodCertificate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors policyErrors)
        //{
        //    return true;
        //}

        private string faz_preambulo_padrao(string asap)
        {
            string subtitulo = get_text(asap, "SUBTITULO=", "\n");
            string unidade = get_text(asap, "UNIDADE=", "\n");
            string data = get_text(asap, "DATA=", "\n");
            string dia = data.Substring(0, 2);
            string mes = data.Substring(3, 2);
            mes = numstr_to_month(mes);
            string ano = data.Substring(6, 4);
            string perito1 = get_text(asap, "PERITO1=", "\n");
            string perito2 = get_text(asap, "PERITO2=", "\n");
            string num_ipl = get_text(asap, "NUMERO_IPL=", "\n").Replace("IPL", "Inquérito Policial nº").Replace("RDF", "Registro de Fato nº").Replace("RE", "Registro Especial nº");
            //string documento = get_text(asap, "DOCUMENTO=", "\n").Replace("Of" + (char)65533 + "cio", "Ofício nº"); //caracter desconhecido: losando com interrogação
            string documento = get_text(asap, "DOCUMENTO=", "\n").Replace("Ofício", "Ofício nº").Replace("Despacho", "Despacho nº");
            string data_documento = get_text(asap, "DATA_DOCUMENTO=", "\n");
            string num_sei = get_text(asap, "NUMERO_SIAPRO=", "\n");
            string registro = get_text(asap, "NUMERO_CRIMINALISTICA=", "\n");
            string data_registro = get_text(asap, "DATA_CRIMINALISTICA=", "\n");
            //MessageBox.Show(subtitulo + " " + unidade + " " + data + " " + perito1 + " " + perito2 + " " + num_ipl + " " + documento + " " + data_documento + " " + num_sei + " " + registro + " " + data_registro);
            //string preambulo = null;
            //for (int i = 1; i <= Globals.ThisAddIn.Application.ActiveDocument.Paragraphs.Count; i++)
            //{
            //    string t = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[i].Range.Text;
            //    string t_trim = t.Trim();
            //    if (t_trim.Length > 200)
            //    {
            //        if ((t_trim.Substring(0, 2)).ToLower() == "em")
            //        {
            //            preambulo = t;
            //            break;
            //        }
            //    }
            //}
            //if (preambulo == null) { MessageBox.Show("preambulo não encontrado."); return; }
            //MessageBox.Show(preambulo);
            //string nome_unidade = "Superintendência Regional de Polícia Federal no Maranhão";
            string nome_unidade = unidade_extenso(get_text(unidade, "/"));
            //MessageBox.Show("nome da unidade = " + nome_unidade);
            //string nome_setor_criminalistica = "SETOR TÉCNICO-CIENTÍFICO";
            string ao_sexo_chefe = "o";
            string ao_documento = "o";
            string ao_sexo_perito1 = "o";
            string ao_sexo_perito2 = "o";
            string cargo_chefe;
            if (unidade == "INC/DITEC/PF")
            {
                if (ao_sexo_chefe == "o")
                {
                    cargo_chefe = "Diretor do INSTITUTO NACIONAL DE CRIMINALÍSTICA da Diretoria Técnico-Científica";
                }
                else
                {
                    cargo_chefe = "Diretora do INSTITUTO NACIONAL DE CRIMINALÍSTICA da Diretoria Técnico-Científica";
                }
            }
            else
            {
                cargo_chefe = "Chefe do " + nome_unidade;
            }


            string ao_sexo_peritos;
            string nome_perito1 = get_text(perito1, inicio: null, " (");
            string nome_perito2 = get_text(perito2, inicio: null, " (").Replace("()", "");
            string s_peritos;
            string is_criminais;
            string Elaboraram_oraram;
            if (nome_perito2 == "")
            {
                s_peritos = "";
                is_criminais = "l";
                Elaboraram_oraram = "orou";
                ao_sexo_peritos = ao_sexo_perito1;
            }
            else
            {
                s_peritos = "s";
                is_criminais = "is";
                Elaboraram_oraram = "oraram";
                nome_perito2 = " e " + nome_perito2;
                if (ao_sexo_perito1 == "o" | ao_sexo_perito2 == "o")
                {
                    ao_sexo_peritos = "o";
                }
                else
                {
                    ao_sexo_peritos = "a";
                }
            }


            string preambulo_modelo1 = "Em " + dia + " de " + mes + " de " + ano + ", designad" + ao_sexo_peritos + s_peritos + " pel" + ao_sexo_chefe + " " + cargo_chefe + ", o" + s_peritos + " Perit" + ao_sexo_peritos + s_peritos + " Crimina" + is_criminais + " Federa" + is_criminais + " " + nome_perito1 + nome_perito2 + " elab" + Elaboraram_oraram + " o presente Laudo de Perícia Criminal Federal, no interesse do " + num_ipl + ", a fim de atender ao contido n" + ao_documento + " " + documento + " de " + data_documento + ", protocolado no SEI sob o nº " + num_sei + " e registrado no SISCRIM sob o nº " + registro + ", em " + data_registro + ", descrevendo com verdade e com todas as circunstâncias tudo quanto possa interessar à Justiça e respondendo aos quesitos formulados, abaixo transcritos:" + (char)13;
            //string preambulo_modelo2 = preambulo_modelo1.Replace("respondendo aos quesitos formulados, abaixo transcritos", "atendendo ao abaixo transcrito");
            return preambulo_modelo1;

        }

        private void compara_preambulo(string preambulo_padrao, int paragrafo_do_preambulo)
        {
            string preambulo = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[paragrafo_do_preambulo].Range.Text;

            Globals.ThisAddIn.Application.ActiveDocument.Application.ScreenUpdating = false;
            Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions = true;
            if (preambulo.Replace(" ", "").ToLower().IndexOf("quesito") != -1)
            {
                if (preambulo == preambulo_padrao)
                {
                    MessageBox.Show("Nenhum erro foi encontrado.");
                }
                else
                {
                    //MessageBox.Show("1 OK");
                    //MessageBox.Show(preambulo + System.Environment.NewLine + System.Environment.NewLine + preambulo_padrao);
                    ////MessageBox.Show(preambulo.Length.ToString() + System.Environment.NewLine + System.Environment.NewLine + preambulo_padrao.Length.ToString());
                    //if (preambulo.Replace(((char)176).ToString(), "º").Substring(0, 664) == preambulo_padrao.Substring(0, 664)) { MessageBox.Show("opa"); } //.Replace(((char)176).ToString(), "º")
                    //if (preambulo.Substring(665) == "\n") { MessageBox.Show("eh barra n"); }
                    //string result = "";
                    //foreach (char c in preambulo_padrao.Substring(664)) { result += (int)c + " "; }
                    //MessageBox.Show(result);
                    Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[paragrafo_do_preambulo].Range.InsertBefore("\n");
                    Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[paragrafo_do_preambulo].Range.Text = preambulo_padrao;
                    Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[paragrafo_do_preambulo].Next().Range.Text = "";
                }
            }
            else
            {
                if (preambulo == preambulo_padrao.Replace("respondendo aos quesitos formulados, abaixo transcritos", "atendendo ao abaixo transcrito"))
                {
                    MessageBox.Show("Nenhum erro foi encontrado.");
                }
                else
                {
                    //MessageBox.Show("2 OK");
                    Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[paragrafo_do_preambulo].Range.InsertBefore("\n");
                    Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[paragrafo_do_preambulo].Range.Text = preambulo_padrao.Replace("respondendo aos quesitos formulados, abaixo transcritos", "atendendo ao abaixo transcrito");
                    Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[paragrafo_do_preambulo].Next().Range.Text = "";
                }
            }

            Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions = false;
            Globals.ThisAddIn.Application.ActiveDocument.Application.ScreenUpdating = true;
            //preambulo_padrao = "Gustavo Vieira";
            //Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[paragrafo_do_preambulo].Range.Text = "Otavio";
            //string preambulo = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[paragrafo_do_preambulo].Range.Text;

            //string result = preambulo;



            //MessageBox.Show(result);
        }

        private string unidade_extenso(string un)
        {
            string estado;
            string cidade;
            //MessageBox.Show(get_text(un, inicio: null, "/"));
            switch (get_text(un, inicio: null, "/"))
            {
                case "SR":
                    switch (un.Substring(un.Length - 2))
                    {
                        case "AC":
                            estado = "no Acre";
                            break;
                        case "AL":
                            estado = "em Alagoas";
                            break;
                        case "AP":
                            estado = "no Amapá";
                            break;
                        case "AM":
                            estado = "no Amazonas";
                            break;
                        case "BA":
                            estado = "na Bahia";
                            break;
                        case "CE":
                            estado = "no Ceará";
                            break;
                        case "ES":
                            estado = "no Espírito Santo";
                            break;
                        case "GO":
                            estado = "em Goiás";
                            break;
                        case "MA":
                            estado = "no Maranhão";
                            break;
                        case "MT":
                            estado = "em Mato Grosso";
                            break;
                        case "MS":
                            estado = "em Mato Grosso do Sul";
                            break;
                        case "MG":
                            estado = "em Minas Gerais";
                            break;
                        case "PA":
                            estado = "no Pará";
                            break;
                        case "PB":
                            estado = "na Paraíba";
                            break;
                        case "PR":
                            estado = "no Paraná";
                            break;
                        case "PE":
                            estado = "em Pernambuco";
                            break;
                        case "PI":
                            estado = "no Piauí";
                            break;
                        case "RJ":
                            estado = "no Rio de Janeiro";
                            break;
                        case "RN":
                            estado = "no Rio Grande do Norte";
                            break;
                        case "RS":
                            estado = "no Rio Grande do Sul";
                            break;
                        case "RO":
                            estado = "em Rondônia";
                            break;
                        case "RR":
                            estado = "em Roraima";
                            break;
                        case "SC":
                            estado = "em Santa Catarina";
                            break;
                        case "SP":
                            estado = "em São Paulo";
                            break;
                        case "SE":
                            estado = "em Sergipe";
                            break;
                        case "TO":
                            estado = "no Tocantins";
                            break;
                        case "DF":
                            estado = "no Distrito Federal";
                            break;
                        default:
                            estado = null;
                            break;
                    }
                    return "SETOR TÉCNICO-CIENTÍFICO da Superintendência Regional de Polícia Federal " + estado;
                case "DPF":
                    //MessageBox.Show(get_text(un, "DPF/", "/") + " opa " + un);

                    switch (get_text(un, "DPF/", "/"))
                    {
                        //case "AGA":
                        //    cidade = "";
                        //    break;
                        //case "ANS":
                        //    cidade = "";
                        //    break;
                        //case "AQA":
                        //    cidade = "";
                        //    break;
                        //case "ARS":
                        //    cidade = "";
                        //    break;
                        case "ARU":
                            cidade = "em Araçatuba";
                            break;
                        //case "ATM":
                        //    cidade = "";
                        //    break;
                        //case "BGE":
                        //    cidade = "";
                        //    break;
                        //case "BRA":
                        //    cidade = "";
                        //    break;
                        //case "BRG":
                        //    cidade = "";
                        //    break;
                        //case "BRU":
                        //    cidade = "";
                        //    break;
                        //case "CAC":
                        //    cidade = "";
                        //    break;
                        //case "CAE":
                        //    cidade = "";
                        //    break;
                        case "CAS":
                            cidade = "em Campinas";
                            break;
                        //case "CCM":
                        //    cidade = "";
                        //    break;
                        //case "CGE":
                        //    cidade = "";
                        //    break;
                        //case "CHI":
                        //    cidade = "";
                        //    break;
                        //case "CIT":
                        //    cidade = "";
                        //    break;
                        //case "CRA":
                        //    cidade = "";
                        //    break;
                        //case "CRU":
                        //    cidade = "";
                        //    break;
                        //case "CXA":
                        //    cidade = "";
                        //    break;
                        //case "CXS":
                        //    cidade = "";
                        //    break;
                        //case "CZO":
                        //    cidade = "";
                        //    break;
                        //case "CZS":
                        //    cidade = "";
                        //    break;
                        //case "DCQ":
                        //    cidade = "";
                        //    break;
                        case "DRS":
                            cidade = "em Dourados";
                            break;
                        //case "DVS":
                        //    cidade = "";
                        //    break;
                        //case "EPA":
                        //    cidade = "";
                        //    break;
                        case "FIG":
                            cidade = "em Foz do Iguaçu";
                            break;
                        //case "GMI":
                        //    cidade = "";
                        //    break;
                        //case "GOY":
                        //    cidade = "";
                        //    break;
                        //case "GPB":
                        //    cidade = "";
                        //    break;
                        case "GRA":
                            cidade = "em Guaíra";
                            break;
                        //case "GVS":
                        //    cidade = "";
                        //    break;
                        //case "IJI":
                        //    cidade = "";
                        //    break;
                        //case "ILS":
                        //    cidade = "";
                        //    break;
                        //case "IPN":
                        //    cidade = "";
                        //    break;
                        //case "ITZ":
                        //    cidade = "";
                        //    break;
                        case "JFA":
                            cidade = " em Juiz de Fora";
                            break;
                        //case "JGO":
                        //    cidade = "";
                        //    break;
                        //case "JLS":
                        //    cidade = "";
                        //    break;
                        case "JNE":
                            cidade = "em Juazeiro do Norte";
                            break;
                        //case "JPN":
                        //    cidade = "";
                        //    break;
                        //case "JTI":
                        //    cidade = "";
                        //    break;
                        //case "JVE":
                        //    cidade = "";
                        //    break;
                        case "JZO":
                            cidade = "em Juazeiro";
                            break;
                        case "LDA":
                            cidade = "em Londrina";
                            break;
                        //case "LGE":
                        //    cidade = "";
                        //    break;
                        //case "LIV":
                        //    cidade = "";
                        //    break;
                        //case "MBA":
                        //    cidade = "";
                        //    break;
                        //case "MCE":
                        //    cidade = "";
                        //    break;
                        //case "MGA":
                        //    cidade = "";
                        //    break;
                        case "MII":
                            cidade = "em Marília";
                            break;
                        //case "MOC":
                        //    cidade = "";
                        //    break;
                        //case "MOS":
                        //    cidade = "";
                        //    break;
                        //case "NIG":
                        //    cidade = "";
                        //    break;
                        //case "NRI":
                        //    cidade = "";
                        //    break;
                        //case "NVI":
                        //    cidade = "";
                        //    break;
                        //case "OPE":
                        //    cidade = "";
                        //    break;
                        //case "PAC":
                        //    cidade = "";
                        //    break;
                        //case "PAT":
                        //    cidade = "";
                        //    break;
                        //case "PCA":
                        //    cidade = "";
                        //    break;
                        case "PDE":
                            cidade = "em Presidente Prudente";
                            break;
                        case "PFO":
                            cidade = "em Passo Fundo";
                            break;
                        //case "PGZ":
                        //    cidade = "";
                        //    break;
                        //case "PHB":
                        //    cidade = "";
                        //    break;
                        //case "PNG":
                        //    cidade = "";
                        //    break;
                        //case "PPA":
                        //    cidade = "";
                        //    break;
                        //case "PSO":
                        //    cidade = "";
                        //    break;
                        case "PTS":
                            cidade = "em Pelotas";
                            break;
                        //case "RDO":
                        //    cidade = "";
                        //    break;
                        //case "RGE":
                        //    cidade = "";
                        //    break;
                        //case "ROO":
                        //    cidade = "";
                        //    break;
                        case "RPO":
                            cidade = "em Ribeirão Preto";
                            break;
                        //case "SAG":
                        //    cidade = "";
                        //    break;
                        //case "SBA":
                        //    cidade = "";
                        //    break;
                        //case "SCS":
                        //    cidade = "";
                        //    break;
                        //case "SGO":
                        //    cidade = "";
                        //    break;
                        case "SIC":
                            cidade = "em Sinop";
                            break;
                        //case "SJE":
                        //    cidade = "";
                        //    break;
                        case "SJK":
                            cidade = "em São José dos Campos";
                            break;
                        case "SMA":
                            cidade = "em Santa Maria";
                            break;
                        //case "SMT":
                        //    cidade = "";
                        //    break;
                        case "SNM":
                            cidade = "em Santarém";
                            break;
                        case "SOD":
                            cidade = "em Sorocaba";
                            break;
                        //case "SSB":
                        //    cidade = "";
                        //    break;
                        case "STS":
                            cidade = "em Santos";
                            break;
                        //case "TBA":
                        //    cidade = "";
                        //    break;
                        //case "TLS":
                        //    cidade = "";
                        //    break;
                        case "UDI":
                            cidade = "em Uberlândia";
                            break;
                        //case "UGA":
                        //    cidade = "";
                        //    break;
                        //case "URA":
                        //    cidade = "";
                        //    break;
                        //case "VAG":
                        //    cidade = "";
                        //    break;
                        //case "VDC":
                        //    cidade = "";
                        //    break;
                        case "VLA":
                            cidade = "em Vilhena";
                            break;
                        //case "VRA":
                        //    cidade = "";
                        //    break;
                        //case "XAP":
                        //    cidade = "";
                        //    break;
                        default:
                            cidade = null;
                            break;
                    }
                    return "NÚCLEO TÉCNICO-CIENTÍFICO da Delegacia de Polícia Federal " + cidade;
                default:
                    //MessageBox.Show("hum");
                    return null;
            }

        }

        private string numstr_to_month(string mes)
        {
            switch (mes)
            {
                case "01":
                    return "janeiro";
                case "02":
                    return "fevereiro";
                case "03":
                    return "março";
                case "04":
                    return "abril";
                case "05":
                    return "maio";
                case "06":
                    return "junho";
                case "07":
                    return "julho";
                case "08":
                    return "agosto";
                case "09":
                    return "setembro";
                case "10":
                    return "outubro";
                case "11":
                    return "novembro";
                case "12":
                    return "dezembro";
                default:
                    return null;
            }
        }

        private void button_numera_paragrafos_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button_confere_formatacao_Click(object sender, RibbonControlEventArgs e)
        {
            foreach (Microsoft.Office.Interop.Word.Shape ishape in Globals.ThisAddIn.Application.Selection.Range.ShapeRange)
            {
                MessageBox.Show(ishape.Type.ToString());

            }

        }

        //private void comboBox_insere_TextChanged(object sender, RibbonControlEventArgs e)
        //{
        //    switch (comboBox_insere.Text)
        //    {
        //        case "Borda preta 0,5 pt":
        //            Globals.ThisAddIn.Application.ScreenUpdating = false;
        //            foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
        //            {
        //                if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
        //                {
        //                    ishape.Line.Visible = MsoTriState.msoTrue;
        //                    ishape.Line.Weight = (float)0.5;
        //                    ishape.Line.ForeColor.RGB = Color.FromArgb(0, 0, 0).ToArgb();
        //                }
        //            }
        //            Globals.ThisAddIn.Application.ScreenUpdating = true;
        //            break;
        //        case "Borda vermelha 2 pt":
        //            Globals.ThisAddIn.Application.ScreenUpdating = false;
        //            foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
        //            {
        //                if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
        //                {
        //                    ishape.Line.Visible = MsoTriState.msoTrue;
        //                    ishape.Line.Weight = 2;
        //                    ishape.Line.ForeColor.RGB = Color.FromArgb(0, 0, 255).ToArgb();
        //                }
        //            }
        //            Globals.ThisAddIn.Application.ScreenUpdating = true;
        //            break;
        //        case "Borda amarela 3 pt":
        //            Globals.ThisAddIn.Application.ScreenUpdating = false;
        //            foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
        //            {
        //                if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
        //                {
        //                    ishape.Line.Visible = MsoTriState.msoTrue;
        //                    ishape.Line.Weight = 3;
        //                    ishape.Line.ForeColor.RGB = Color.FromArgb(0, 255, 255).ToArgb();
        //                }
        //            }
        //            Globals.ThisAddIn.Application.ScreenUpdating = true;
        //            break;                    
        //    }
        //}

        //private void comboBox_remove_TextChanged(object sender, RibbonControlEventArgs e)
        //{
        //    switch (comboBox_remove.Text)
        //    {
        //        case "Borda":
        //            Globals.ThisAddIn.Application.ScreenUpdating = false;
        //            foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
        //            {
        //                if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
        //                {
        //                    ishape.Line.Visible = MsoTriState.msoFalse;
        //                }
        //            }
        //            Globals.ThisAddIn.Application.ScreenUpdating = true;
        //            break;
        //        case "Formatação":
        //            Globals.ThisAddIn.Application.ScreenUpdating = false;
        //            foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
        //            {
        //                if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
        //                {
        //                    ishape.Reset();
        //                }
        //            }
        //            Globals.ThisAddIn.Application.ScreenUpdating = true;
        //            break;
        //        case "Forma":
        //            Globals.ThisAddIn.Application.ScreenUpdating = false;
        //            //MessageBox.Show(Globals.ThisAddIn.Application.Selection.ShapeRange.Count.ToString());
        //            //MessageBox.Show(Globals.ThisAddIn.Application.Selection.Range.ShapeRange.Count.ToString());
        //            List<Microsoft.Office.Interop.Word.Shape> listaShapes = new List<Microsoft.Office.Interop.Word.Shape>();
        //            foreach (Microsoft.Office.Interop.Word.Shape ishape in Globals.ThisAddIn.Application.Selection.Range.ShapeRange)
        //            {
        //                //MessageBox.Show(ishape.Type.ToString());
        //                if (ishape.Type == MsoShapeType.msoAutoShape | ishape.Type == MsoShapeType.msoFreeform)
        //                {
        //                    listaShapes.Add(ishape);
        //                }
        //            }
        //            foreach (Microsoft.Office.Interop.Word.Shape ishape in listaShapes)
        //            {
        //                ishape.Delete();
        //            }
        //            Globals.ThisAddIn.Application.ScreenUpdating = true;
        //            break;
        //        case "Texto Alt":
        //            Globals.ThisAddIn.Application.ScreenUpdating = false;
        //            foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
        //            {
        //                if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
        //                {
        //                    ishape.AlternativeText = "";
        //                }
        //            }
        //            Globals.ThisAddIn.Application.ScreenUpdating = true;
        //            break;
        //    }
        //}

        private void button_borda_preta_Click(object sender, RibbonControlEventArgs e)
        {
            new Thread(() =>
            {
                iClass_Buttons.muda_imagem("menu_inserir_imagem", Properties.Resources.load_icon_png_7969);
                menu_inserir_imagem.Enabled = false;
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                //string selectedText = Globals.ThisAddIn.Application.Selection.Range.ToString();
                //int L1 = selectedText.Split('\r').Length;
                //MessageBox.Show(L1.ToString());
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.Line.Visible = MsoTriState.msoTrue;
                        ishape.Line.Weight = (float)0.5;
                        ishape.Line.ForeColor.RGB = Color.FromArgb(0, 0, 0).ToArgb();
                    }
                }
                //selectedText = Globals.ThisAddIn.Application.Selection.Range.ToString();
                //int L2 = selectedText.Split('\r').Length;
                //MessageBox.Show(L2.ToString());
                //if (L2 > L1) 
                //{
                //    MessageBox.Show("opa");
                //}
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                iClass_Buttons.muda_imagem("menu_inserir_imagem", Properties.Resources._);
                menu_inserir_imagem.Enabled = true;
            }).Start();

        }

        private void button_borda_vermelha_Click(object sender, RibbonControlEventArgs e)
        {
            new Thread(() =>
            {
                iClass_Buttons.muda_imagem("menu_inserir_imagem", Properties.Resources.load_icon_png_7969);
                menu_inserir_imagem.Enabled = false;
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.Line.Visible = MsoTriState.msoTrue;
                        ishape.Line.Weight = 2;
                        ishape.Line.ForeColor.RGB = Color.FromArgb(0, 0, 255).ToArgb();
                    }
                }
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                iClass_Buttons.muda_imagem("menu_inserir_imagem", Properties.Resources._);
                menu_inserir_imagem.Enabled = true;
            }).Start();
        }

        private void button_borda_amarela_Click(object sender, RibbonControlEventArgs e)
        {
            new Thread(() =>
            {
                iClass_Buttons.muda_imagem("menu_inserir_imagem", Properties.Resources.load_icon_png_7969);
                menu_inserir_imagem.Enabled = false;
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.Line.Visible = MsoTriState.msoTrue;
                        ishape.Line.Weight = 3;
                        ishape.Line.ForeColor.RGB = Color.FromArgb(0, 255, 255).ToArgb();
                    }
                }
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                iClass_Buttons.muda_imagem("menu_inserir_imagem", Properties.Resources._);
                menu_inserir_imagem.Enabled = true;
            }).Start();
        }

        private void button_legenda_imagem_Click(object sender, RibbonControlEventArgs e)
        {
            new Thread(() =>
            {
                iClass_Buttons.muda_imagem("menu_inserir_imagem", Properties.Resources.load_icon_png_7969);
                menu_inserir_imagem.Enabled = false;
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                Range r = Globals.ThisAddIn.Application.Selection.Range;

                string estilo_nome_baseado = "Legenda";
                Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);

                List<InlineShape> list_InlineShape = new List<InlineShape>();

                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        list_InlineShape.Add(ishape);
                    }
                }
                foreach (InlineShape ishape in list_InlineShape)
                {
                    ishape.Select();
                    //MessageBox.Show(Globals.ThisAddIn.Application.Selection.Paragraphs[1].Next().Range.Characters.Count.ToString());
                    //MessageBox.Show(Globals.ThisAddIn.Application.Selection.Paragraphs[1].Next().Range.Text.Substring(0,7));

                    if (Globals.ThisAddIn.Application.Selection.Paragraphs[1].Next() != null)
                    {
                        if (Globals.ThisAddIn.Application.Selection.Paragraphs[1].Next().Range.Characters.Count >= 7)
                        {
                            if (Globals.ThisAddIn.Application.Selection.Paragraphs[1].Next().Range.Text.Substring(0, 7) == "Figura ")
                            {
                                //r.Select();
                                //Globals.ThisAddIn.Application.ScreenUpdating = true;
                                //return;
                                continue;
                            }
                        }
                    }
                    //if (ishape.Range.Paragraphs[1].Range.InlineShapes.Count > 1)
                    //{
                    //    //MessageBox.Show(ishape.Range.Paragraphs[1].Range.InlineShapes.Count.ToString());
                    //    ////MessageBox.Show(ishape.Range.Text);
                    //    ////MessageBox.Show(ishape.Range.Paragraphs[1].Range.InlineShapes[1].Range.Text);
                    //    //MessageBox.Show(ishape.Equals(ishape.Range.Paragraphs[1].Range.ShapeRange[2]).ToString());
                    //    //if (ishape.Range.Paragraphs[1].Range.InlineShapes[1] == ishape.Range.Paragraphs[1].Range.InlineShapes[1])
                    //    //{
                    //    //    MessageBox.Show("opoppaaa");
                    //    //    //r.Select();
                    //    //    //Globals.ThisAddIn.Application.ScreenUpdating = true;
                    //    //    continue;
                    //    //}

                    //    continue;
                    //}
                    if (IsLastShapeInParagraph(ishape))
                    {
                        bool label_existe = false;
                        foreach (CaptionLabel label in Globals.ThisAddIn.Application.CaptionLabels)
                        {
                            if (label.Name == "Figura") { label_existe = true; }
                        }
                        if (!label_existe) { Globals.ThisAddIn.Application.CaptionLabels.Add("Figura"); }

                        Globals.ThisAddIn.Application.Selection.InsertCaption(Label: "Figura", Title: " " + ((char)8211).ToString(), TitleAutoText: "", Position: WdCaptionPosition.wdCaptionPositionBelow, ExcludeLabel: 0);
                        Globals.ThisAddIn.Application.Selection.set_Style((object)"07 - Legendas de Figuras (PeriTAB)");
                        Globals.ThisAddIn.Application.Selection.InsertAfter(" ");
                        Globals.ThisAddIn.Application.Run("alinha_legenda");
                    }
                }
                r.Select();
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                iClass_Buttons.muda_imagem("menu_inserir_imagem", Properties.Resources._);
                menu_inserir_imagem.Enabled = true;
            }).Start();
        }

        private void button_remove_borda_Click(object sender, RibbonControlEventArgs e)
        {
            new Thread(() =>
            {
                iClass_Buttons.muda_imagem("menu_remover_imagem", Properties.Resources.load_icon_png_7969);
                menu_remover_imagem.Enabled = false;
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.Line.Visible = MsoTriState.msoFalse;
                    }
                }
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                iClass_Buttons.muda_imagem("menu_remover_imagem", Properties.Resources.x);
                menu_remover_imagem.Enabled = true;
            }).Start();
        }

        private void button_remove_formatacao_Click_1(object sender, RibbonControlEventArgs e)
        {
            new Thread(() =>
            {
                iClass_Buttons.muda_imagem("menu_remover_imagem", Properties.Resources.load_icon_png_7969);
                menu_remover_imagem.Enabled = false;
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.Reset();
                    }
                }
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                iClass_Buttons.muda_imagem("menu_remover_imagem", Properties.Resources.x);
                menu_remover_imagem.Enabled = true;
            }).Start();
        }

        private void button_remove_forma_Click(object sender, RibbonControlEventArgs e)
        {
            new Thread(() =>
            {
                iClass_Buttons.muda_imagem("menu_remover_imagem", Properties.Resources.load_icon_png_7969);
                menu_remover_imagem.Enabled = false;
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                //MessageBox.Show(Globals.ThisAddIn.Application.Selection.ShapeRange.Count.ToString());
                //MessageBox.Show(Globals.ThisAddIn.Application.Selection.Range.ShapeRange.Count.ToString());
                List<Microsoft.Office.Interop.Word.Shape> listaShapes = new List<Microsoft.Office.Interop.Word.Shape>();
                foreach (Microsoft.Office.Interop.Word.Shape ishape in Globals.ThisAddIn.Application.Selection.Range.ShapeRange)
                {
                    //MessageBox.Show(ishape.Type.ToString());
                    if (ishape.Type == MsoShapeType.msoAutoShape | ishape.Type == MsoShapeType.msoFreeform | ishape.Type == MsoShapeType.msoLine | ishape.Type == MsoShapeType.msoTextBox)
                    {
                        listaShapes.Add(ishape);
                    }
                }
                foreach (Microsoft.Office.Interop.Word.Shape ishape in listaShapes)
                {
                    ishape.Delete();
                }
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                iClass_Buttons.muda_imagem("menu_remover_imagem", Properties.Resources.x);
                menu_remover_imagem.Enabled = true;
            }).Start();
        }

        private void button_remove_texto_alt_Click(object sender, RibbonControlEventArgs e)
        {
            new Thread(() =>
            {
                iClass_Buttons.muda_imagem("menu_remover_imagem", Properties.Resources.load_icon_png_7969);
                menu_remover_imagem.Enabled = false;
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        ishape.AlternativeText = "";
                    }
                }
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                iClass_Buttons.muda_imagem("menu_remover_imagem", Properties.Resources.x);
                menu_remover_imagem.Enabled = true;
            }).Start();
        }

        private void button_remove_imagem_Click(object sender, RibbonControlEventArgs e)
        {
            new Thread(() =>
            {
                iClass_Buttons.muda_imagem("menu_remover_imagem", Properties.Resources.load_icon_png_7969);
                menu_remover_imagem.Enabled = false;
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                List<InlineShape> listaShapes = new List<InlineShape>();
                foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
                {
                    if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        listaShapes.Add(ishape);
                    }
                }
                foreach (InlineShape ishape in listaShapes)
                {
                    ishape.Delete();
                }
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                iClass_Buttons.muda_imagem("menu_remover_imagem", Properties.Resources.x);
                menu_remover_imagem.Enabled = true;
            }).Start();
        }

        private void button_legenda_tabela_Click(object sender, RibbonControlEventArgs e)
        {
            new Thread(() =>
            {
                iClass_Buttons.muda_imagem("menu_inserir_tabela", Properties.Resources.load_icon_png_7969);
                menu_inserir_tabela.Enabled = false;
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                Range r = Globals.ThisAddIn.Application.Selection.Range;

                string estilo_nome_baseado = "Legenda";
                Globals.ThisAddIn.Application.OrganizerCopy(Ribbon1.Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, estilo_nome_baseado, WdOrganizerObject.wdOrganizerObjectStyles);

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
                r.Select();
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                iClass_Buttons.muda_imagem("menu_inserir_tabela", Properties.Resources._);
                menu_inserir_tabela.Enabled = true;
            }).Start();
        }
        private void button_centralizar_tabela_Click(object sender, RibbonControlEventArgs e)
        {
            new Thread(() =>
            {
                iClass_Buttons.muda_imagem("menu_formatacao_tabela", Properties.Resources.load_icon_png_7969);
                menu_formatacao_tabela.Enabled = false;
                Globals.ThisAddIn.Application.ScreenUpdating = false;

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

                Globals.ThisAddIn.Application.ScreenUpdating = true;
                iClass_Buttons.muda_imagem("menu_formatacao_tabela", Properties.Resources.formatacao2);
                menu_formatacao_tabela.Enabled = true;
            }).Start();
        }

        private void button_minuscula_campos_Click(object sender, RibbonControlEventArgs e)
        {
            new Thread(() =>
            {
                iClass_Buttons.muda_imagem("menu_formatacao_campos", Properties.Resources.load_icon_png_7969);
                menu_formatacao_campos.Enabled = false;
                Globals.ThisAddIn.Application.ScreenUpdating = false;

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
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                iClass_Buttons.muda_imagem("menu_formatacao_campos", Properties.Resources.formatacao2);
                menu_formatacao_campos.Enabled = true;
            }).Start();
        }

        private void button_abre_SISCRIM_Click(object sender, RibbonControlEventArgs e)
        {
            new Thread(() =>
            {
                // Configurações iniciais
                Stopwatch stopwatch = new Stopwatch(); if (versao() == null) { stopwatch.Start(); } // Inicia o cronômetro para medir o tempo de execução da Thread
                bool success = true;
                string msg_StatusBar = "";
                string msg_Falha = "";
                iClass_Buttons.muda_imagem("button_abre_SISCRIM", Properties.Resources.load_icon_png_7969);
                button_abre_SISCRIM.Enabled = false;

                string[] identificadores_laudo = pega_identificadores_laudo();

                string num_laudo = identificadores_laudo[0];
                string ano_laudo = identificadores_laudo[1];
                string unidade_laudo = identificadores_laudo[2];

                string localpath = Globals.Ribbons.Ribbon1.GetLocalPath(Globals.ThisAddIn.Application.ActiveDocument.FullName);

                if (File.Exists(localpath.Substring(0, localpath.LastIndexOf(".")) + ".pdf") | File.Exists(localpath.Substring(0, localpath.LastIndexOf(".")) + "_assinado.pdf"))
                {
                    if (num_laudo == null | ano_laudo == null | unidade_laudo == null)
                    {
                        //MessageBox.Show("Referência do laudo não encontrada.");
                        success = false;
                        msg_Falha = "Referência do laudo não encontrada.";
                    }
                    else
                    {
                        System.Diagnostics.Process.Start("https://www.ditec.pf.gov.br:8443/sistemas/criminalistica/controle_documento.php?action=localizar_resultado&d-codigo_tipo_documento=2704&d-numero_documento=" + num_laudo + "&d-ano_documento=" + ano_laudo + "&d-sigla_orgao_emissor-ilike=" + unidade_laudo + "&codigo_unidade_registro_pesquisa=");
                    }
                }
                else
                {
                    string identificadores_requisicao = null;
                    if (ano_laudo != null & unidade_laudo != null)
                    {
                        identificadores_requisicao = Microsoft.VisualBasic.Interaction.InputBox("O PDF do laudo ainda não foi gerado. Digite o número do registro da requisição:", "", "REGISTRO Nº xxx / " + ano_laudo + " - " + unidade_laudo.ToUpper());
                    }
                    else
                    {
                        identificadores_requisicao = Microsoft.VisualBasic.Interaction.InputBox("O PDF do laudo ainda não foi gerado. Digite o número do registro da requisição:", "", "REGISTRO Nº numero / ano - unidade");
                    }

                    if (identificadores_requisicao != "")
                    {
                        string num_registro = null;
                        string ano_registro = null;
                        string unidade_registro = null;
                        string identificadores_requisicao_mod = identificadores_requisicao.ToLower().Replace(" ", "");

                        num_registro = get_text(identificadores_requisicao_mod, "nº", "/");
                        ano_registro = get_text(identificadores_requisicao_mod, "/", "-");
                        unidade_registro = get_text(identificadores_requisicao_mod, "-");
                        //MessageBox.Show(unidade_registro);
                        int codigo_registro = pega_codigo_registro(unidade_registro);
                        //MessageBox.Show(codigo_registro.ToString());

                        if (num_registro == null | ano_registro == null | unidade_registro == null | !int.TryParse(num_registro, out _) | !int.TryParse(ano_registro, out _) | codigo_registro == 0)
                        {
                            //MessageBox.Show("Número do registro da requisição inválido.");
                            success = false;
                            msg_Falha = "Número do registro da requisição inválido.";
                        }
                        else
                        {
                            //System.Diagnostics.Process.Start("https://www.ditec.pf.gov.br:8443/sistemas/criminalistica/controle_documento.php?action=localizar_resultado&d-numero_registro=" + num_registro + "&d-ano_registro=" + ano_registro + "&codigo_unidade_registro_pesquisa=" + 3347 + "&comando=Procurar"/*unidade_registro + "&codigo_unidade_registro_pesquisa="*/);
                            //System.Diagnostics.Process.Start("https://www.ditec.pf.gov.br:8443/sistemas/criminalistica/controle_documento.php?action=localizar_resultado&d-codigo_tipo_documento=&d-numero_documento=&d-ano_documento=&d-sigla_orgao_emissor-ilike=&d-codigo_subtipo_documento=&p-codigo_tipo_procedimento=&p-numero_procedimento-ilike=&p-sigla_orgao-ilike=&sa-nome_signatario-ilike=&sa-funcao_signatario-ilike=&d-data_emissao-ge=&d-data_emissao-le=&d-numero_siapro=&d-numero_registro_epol=&d-assunto-ilike=&d-operacao-ilike=&dds-nome-ilike=&dc-codigo_tipo_documento_citacao=&dc-nome-ilike=&dc-cpf=&dc-cnpj=&dc-observacao-ilike=&d-marcador-ilike=&numero_registro=" + num_registro + " &ano_registro=" + ano_registro + "&d-data_protocolo-ge=&d-data_protocolo-le=&d-excluido=&d-recebido=&tl-nome-ilike=&sl-nome-ilike=&dm-nome-ilike=&d-nome_sujeito-ilike=&d-codigo_finalidade_documento=&d-codigo_situacao_documento=&soe-codigo_tipo_sujeito=&soe-sigla_uf=&soe-nome-ilike=&codigo_unidade_registro_pesquisa=" + "3347" + "&d-usuario_criacao-ilike=&d-ignorar_registros_adicionais=0&d-codigo_area_exame=&d-urgencia=&d-motivo_urgencia-ilike=&d-data_limite-ge=&d-data_limite-le=&d-sigiloso=&d-observacao-ilike=&d-conteudo-ilike=&oac-indice-tsquery=&d-publicado=N%C3%A3o&d-naopublicado=N%C3%A3o&dcae-codigo_tipo_material=&dcae-medida=&dcae-codigo_unidade_medida=&dccv-renavam-ilike=&dccv-marca-ilike=&dccv-modelo-ilike=&dccv-placa-ilike=&dccv-chassi-ilike=&dccv-ano_fabricacao-ilike=&dccv-ano_modelo-ilike=&dccv-cor-ilike=&dccv-observacoes-ilike=&dcad-data=&dcad-sigla_uf_municipio=&dcad-codigo_municipio=&dcad-codigo_categoria_droga=&dcad-codigo_droga=&dcad-massa=&dcad-codigo_unidade_medida_massa=&dcad-volume=&dcad-codigo_unidade_medida_volume=&dcad-numero_itens=&dcad-massa_media_unitaria=&dcad-codigo_unidade_medida_massa_media_unitaria=&comando=Procurar");
                            //System.Diagnostics.Process.Start("https://www.ditec.pf.gov.br:8443/sistemas/criminalistica/controle_documento.php?action=localizar_resultado&numero_registro=" + num_registro + " &ano_registro=" + ano_registro + "&codigo_unidade_registro_pesquisa=" + "3347" + "&d-ignorar_registros_adicionais=0" + "&comando=Procurar");

                            System.Diagnostics.Process.Start("https://www.ditec.pf.gov.br:8443/sistemas/criminalistica/controle_documento.php?action=localizar_resultado&numero_registro=" + num_registro + " &ano_registro=" + ano_registro + "&codigo_unidade_registro_pesquisa=" + codigo_registro + "&d-ignorar_registros_adicionais=0");
                        }
                    }
                }

                // Mensagens da Thread
                if (success) { msg_StatusBar = "Abre SISCRIM: Sucesso"; } else { msg_StatusBar = "Abre SISCRIM: Falha"; }
                if (versao() == null) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
                {
                    stopwatch.Stop();
                    msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                }
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
                if (!success && msg_Falha != "") MessageBox.Show(msg_Falha, "Abre SISCRIM");

                // Configurações finais
                //Globals.ThisAddIn.Application.ScreenUpdating = true;
                iClass_Buttons.muda_imagem("button_abre_SISCRIM", Properties.Resources.subir2);
                button_abre_SISCRIM.Enabled = true;
            }).Start();
        }

        private int pega_codigo_registro(string unidade_registro)
        {
            switch (unidade_registro.ToUpper().Trim())
            {
                case "DITEC/PF":
                    return 3014;
                case "INC/DITEC/PF":
                    return 3015;
                case "NUTEC/DPF/ARU/SP":
                    return 4597948;
                case "NUTEC/DPF/CAS/SP":
                    return 4598011;
                case "NUTEC/DPF/DRS/MS":
                    return 1064321;
                case "NUTEC/DPF/FIG/PR":
                    return 3571;
                case "NUTEC/DPF/GRA/PR":
                    return 4606718;
                case "NUTEC/DPF/JFA/MG":
                    return 3423927;
                case "NUTEC/DPF/JNE/CE":
                    return 4542507;
                case "NUTEC/DPF/JZO/BA":
                    return 4580096;
                case "NUTEC/DPF/LDA/PR":
                    return 1064330;
                case "NUTEC/DPF/MII/SP":
                    return 4597132;
                case "NUTEC/DPF/PDE/SP":
                    return 4598053;
                case "NUTEC/DPF/PFO/RS":
                    return 6323156;
                case "NUTEC/DPF/PTS/RS":
                    return 6797076;
                case "NUTEC/DPF/RPO/SP":
                    return 1064363;
                case "NUTEC/DPF/SIC/MT":
                    return 4398092;
                case "NUTEC/DPF/SJK/SP":
                    return 6434234;
                case "NUTEC/DPF/SMA/RS":
                    return 3398735;
                case "NUTEC/DPF/SNM/PA":
                    return 4683087;
                case "NUTEC/DPF/SOD/SP":
                    return 4597984;
                case "NUTEC/DPF/STS/SP":
                    return 3849;
                case "NUTEC/DPF/UDI/MG":
                    return 1064366;
                case "NUTEC/DPF/VLA/RO":
                    return 4084597;
                case "SETEC/SR/PF/AC":
                    return 3118;
                case "SETEC/SR/PF/AL":
                    return 3143;
                case "SETEC/SR/PF/AM":
                    return 3168;
                case "SETEC/SR/PF/AP":
                    return 3194;
                case "SETEC/SR/PF/BA":
                    return 3219;
                case "SETEC/SR/PF/CE":
                    return 3244;
                case "SETEC/SR/PF/DF":
                    return 3269;
                case "SETEC/SR/PF/ES":
                    return 3297;
                case "SETEC/SR/PF/GO":
                    return 3322;
                case "SETEC/SR/PF/MA":
                    return 3347;
                case "SETEC/SR/PF/MG":
                    return 3372;
                case "SETEC/SR/PF/MS":
                    return 3397;
                case "SETEC/SR/PF/MT":
                    return 3422;
                case "SETEC/SR/PF/PA":
                    return 3447;
                case "SETEC/SR/PF/PB":
                    return 3472;
                case "SETEC/SR/PF/PE":
                    return 3497;
                case "SETEC/SR/PF/PI":
                    return 3522;
                case "SETEC/SR/PF/PR":
                    return 3547;
                case "SETEC/SR/PF/RJ":
                    return 3587;
                case "SETEC/SR/PF/RN":
                    return 3641;
                case "SETEC/SR/PF/RO":
                    return 3666;
                case "SETEC/SR/PF/RR":
                    return 3691;
                case "SETEC/SR/PF/RS":
                    return 3716;
                case "SETEC/SR/PF/SC":
                    return 3743;
                case "SETEC/SR/PF/SE":
                    return 3768;
                case "SETEC/SR/PF/SP":
                    return 3797;
                case "SETEC/SR/PF/TO":
                    return 3859;
                case "UTEC/DPF/ITZ/MA":
                    return 4803309;
                case "UTEC/DPF/MBA/PA":
                    return 4682890;
                case "UTEC/DPF/ROO/MT":
                    return 4398074;
                case "UTEC/DPF/SGO/PE":
                    return 7154495;
                default:
                    return 0;
            }
        }













        //private string get_text2(string texto, string inicio = null, string fim = null) //Retona a primeira ocorrência de string entre os strings 'inicio' e 'fim' no string 'texto'.
        //{
        //    if (inicio == null & fim == null) { return null; }

        //    //try
        //    //{
        //        if (inicio == null)
        //        {
        //            return texto.Substring(0, texto.IndexOf(fim));
        //        }
        //        if (fim == null)
        //        {
        //            return texto.Substring(texto.IndexOf(inicio) + inicio.Length);
        //        }
        //    string a1 = (texto.Substring(texto.IndexOf(inicio)));
        //    int a2 = inicio.Length;
        //    int a3 = (texto.Substring(texto.IndexOf(inicio) + inicio.Length)).IndexOf(fim);
        //    return (texto.Substring(texto.IndexOf(inicio))).Substring(inicio.Length, (texto.Substring(texto.IndexOf(inicio) + inicio.Length)).IndexOf(fim));
        //    //}
        //    //catch { return null; }
        //}

    }
}

        //private void button1_Click(object sender, RibbonControlEventArgs e)
        //{
        //    var client = new WebClient();
        //    string url = "https://www.ditec.pf.gov.br/sistemas/criminalistica/controle_documento.php?action=gerar_ini_asap&id=63785170";
        //    string filename = "AsAP_Laudo_191-2023.asap";
        //    client.DownloadFile(url, filename);
        //}


        //    var httpClient = new HttpClient();
            


        //    using (var stream = await httpClient.GetStreamAsync(url))
        //    {
        //        using (var fileStream = new FileStream("300.png", FileMode.CreateNew))
        //        {
        //            await stream.CopyToAsync(fileStream);
        //        }
        //    }
        //}
    //}
        //}