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

namespace PeriTAB
{
    public partial class Ribbon1
    {
        public class Variables
        {
            private static string var1 = Path.GetTempPath() + "PeriTAB_Template_tmp.dotm";
            private static string var2 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PeriTAB");
            private static string var3, var4;
            private static X509Certificate2 var_cert = null;
            private static IExternalSignature var_sig = null;
            public static string caminho_template { get { return var1; } set { } }
            public static string caminho_AppData_Roaming_PeriTAB { get { return var2; } set { } }
            public static string editBox_largura_Text { get { return var3; } set { var3 = value; } }
            public static string editBox_altura_Text { get { return var4; } set { var4 = value; } }
            public static X509Certificate2 cert { get { return var_cert; } set { var_cert = value; } }
            public static IExternalSignature sig { get { return var_sig; } set { var_sig = value; } }
        }

        const string quote = "\"";
        const string slash = @"\";

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //Escreve o Template na pasta tmp e adiciona ela como suplemento.
            try { File.WriteAllBytes(Variables.caminho_template, Properties.Resources.Normal); } catch (IOException ex) { MessageBox.Show("PeriTAB_Template_tmp.dotm em uso"); Globals.ThisAddIn.Application.Quit(); return; }
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

        private void toggleButton_estilos_Click(object sender, RibbonControlEventArgs e)
        {
            var botao_toggle = (Microsoft.Office.Tools.Ribbon.RibbonToggleButton)sender;
            if (botao_toggle.Checked == true) Globals.ThisAddIn.TaskPane1.Visible = true;
            if (botao_toggle.Checked == false) Globals.ThisAddIn.TaskPane1.Visible = false;
        }

        private void button_inserir_sumario_Click(object sender, RibbonControlEventArgs e)
        {
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
                        string extensao = (pathfile[i].Substring(pathfile[i].Length - 4)).ToLower();
                        if (extensao == ".jpg" | extensao == "jpeg" | extensao == ".png" | extensao == ".bmp" | extensao == ".gif" | extensao == "tiff") //Se tem extensao de imagem
                        {
                            Array.Resize(ref pathfile2, n + 1);
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

                        InlineShape imagem = Globals.ThisAddIn.Application.Selection.InlineShapes.AddPicture(pathfile2[i], link, save);
                        MsoTriState LockAspectRatio_i = imagem.LockAspectRatio;
                        imagem.LockAspectRatio = (MsoTriState)1;
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
                        imagem.LockAspectRatio = LockAspectRatio_i;

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
            if (Variables.editBox_largura_Text == null) { Variables.editBox_largura_Text = Class_Buttons.preferences.largura; }
            
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
            if (Variables.editBox_altura_Text == null) { Variables.editBox_altura_Text = Class_Buttons.preferences.altura; }

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
            if (checkBox_referencia.Checked)
            {
                System.Windows.Forms.MessageBox.Show("Cuidado! Excluir/mover/renomear o arquivo da imagem causará perda de referência.");
            }
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
            Globals.ThisAddIn.Application.Run("renomeia_documento");
        }

        private void button_gerar_pdf_Click(object sender, RibbonControlEventArgs e)
        {
            string path = Globals.ThisAddIn.Application.ActiveDocument.FullName;
            string localpath = GetLocalPath(path);
            if (localpath == "") { MessageBox.Show("Não foi possível gerar o PDF."); return; }
            Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(localpath.Substring(0, localpath.LastIndexOf(".")), WdExportFormat.wdExportFormatPDF, UseISO19005_1: true);

            string path_pdf = localpath.Substring(0, localpath.LastIndexOf(".")) + ".pdf";
            if (Globals.Ribbons.Ribbon1.checkBox_assinar.Checked)
            {
                string path_pdf_assinado = localpath.Substring(0, localpath.LastIndexOf(".")) + "_assinado.pdf";

                X509Certificate2 certClient = null;
                X509Store st = new X509Store(StoreName.My, StoreLocation.CurrentUser);
                st.Open(OpenFlags.MaxAllowed);
                IExternalSignature s;
                foreach (X509Certificate2 c in st.Certificates)
                {
                    if (c.Verify() == false) { st.Remove(c); continue; } //Elimina certificado não validados
                    try { s = new X509Certificate2Signature(c, "SHA-256"); } catch { st.Remove(c); } //Elimina certificado que não se pode pegar a assinatura
                }

                switch (st.Certificates.Count)
                {
                    case 0:
                        MessageBox.Show("Nenhum certificado válido encontrado.");
                        return;
                    case 1:
                        certClient = st.Certificates[0];
                        break;
                    default:
                        X509Certificate2Collection collection = X509Certificate2UI.SelectFromCollection(st.Certificates, "Escolha o certificado:", "", X509SelectionFlag.SingleSelection);
                        if (collection.Count > 0)
                        {
                            certClient = collection[0];
                        }
                        else
                        {
                            MessageBox.Show("Nenhum certificado foi selecionado.");
                            return;
                        }
                        break;
                }

                Variables.cert = certClient;

                st.Close();

                //Get Cert Chain
                IList<X509Certificate> chain = new List<X509Certificate>();
                X509Chain x509Chain = new X509Chain();
                x509Chain.Build(certClient);
                foreach (X509ChainElement x509ChainElement in x509Chain.ChainElements)
                {
                    chain.Add(DotNetUtilities.FromX509Certificate(x509ChainElement.Certificate));
                }

                PdfReader inputPdf = new PdfReader(path_pdf);

                FileStream signedPdf = new FileStream(path_pdf_assinado, FileMode.Create);

                PdfStamper pdfStamper = PdfStamper.CreateSignature(inputPdf, signedPdf, '\0');

                IExternalSignature externalSignature = new X509Certificate2Signature(certClient, "SHA-256");

                PdfSignatureAppearance signatureAppearance = pdfStamper.SignatureAppearance;

                //signatureAppearance.SignatureGraphic = Image.GetInstance(pathToSignatureImage);
                //signatureAppearance.SetVisibleSignature(new iTextSharp.text.Rectangle(0, 00, 250, 150), inputPdf.NumberOfPages, "Signature");
                signatureAppearance.SignatureRenderingMode = PdfSignatureAppearance.RenderingMode.DESCRIPTION;

                
                //RSACryptoServiceProvider rsa = new RSACryptoServiceProvider();
                //CspParameters cspp = new CspParameters();
                //cspp.KeyContainerName = rsa.CspKeyContainerInfo.KeyContainerName;
                //cspp.ProviderName = rsa.CspKeyContainerInfo.ProviderName;
                //cspp.ProviderType = rsa.CspKeyContainerInfo.ProviderType;
                //cspp.Flags = CspProviderFlags.NoPrompt;
                //RSACryptoServiceProvider rsa2 = new RSACryptoServiceProvider(cspp);
                //rsa.PersistKeyInCsp = true;

                (new RSACryptoServiceProvider()).PersistKeyInCsp = true; //Define chave persistente. Só pede a senha da primeira vez.


                try { MakeSignature.SignDetached(signatureAppearance, externalSignature, chain, null, null, null, 0, CryptoStandard.CMS); } catch (CryptographicException ex) { return; }
                inputPdf.Close();
                pdfStamper.Close();

                if (File.Exists(path_pdf_assinado))
                {
                    Globals.ThisAddIn.Application.DisplayStatusBar = true; Globals.ThisAddIn.Application.StatusBar = "PDF gerado com sucesso.";
                    if (File.Exists(path_pdf)) { File.Delete(path_pdf); }
                    if (Globals.Ribbons.Ribbon1.checkBox_abrir.Checked) { System.Diagnostics.Process.Start(path_pdf_assinado); }
                }
                else
                {
                    Globals.ThisAddIn.Application.DisplayStatusBar = true; Globals.ThisAddIn.Application.StatusBar = "A geração do PDF falhou.";
                }
            }
            else
            {
                if (File.Exists(path_pdf)) 
                {
                    Globals.ThisAddIn.Application.DisplayStatusBar = true; Globals.ThisAddIn.Application.StatusBar = "PDF gerado com sucesso.";
                    if (Globals.Ribbons.Ribbon1.checkBox_abrir.Checked) { System.Diagnostics.Process.Start(path_pdf); }
                }
                else
                {
                    Globals.ThisAddIn.Application.DisplayStatusBar = true; Globals.ThisAddIn.Application.StatusBar = "A geração do PDF falhou.";
                }
            }
        }

        private string GetLocalPath(string path)
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
                    else { return ""; }
                    localpath = path_onedrive + slash + onedrive_subpasta.Replace("/", slash);
                }
                else
                {
                    return "";
                }
            }
            return localpath;
        }
    }
}
