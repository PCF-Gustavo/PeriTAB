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
using Aspose.Pdf.Operators;
using static System.Windows.Forms.LinkLabel;
using System.Net;
using System.Text;
//using static System.Net.WebRequestMethods;
using System.Net.Http;
using Thinktecture.IdentityModel.Client;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using System.Drawing;
using System.Windows.Controls;
using System.Net.Security;
using Org.BouncyCastle.Crypto.Tls;
using System.Security.Authentication;
using System.Runtime.InteropServices;

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
            Globals.ThisAddIn.Application.ScreenUpdating = false;

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

            Globals.ThisAddIn.Application.ScreenUpdating = true;
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
            PdfReader inputPdf = null;
            bool inputPdf_open = false;
            string path = Globals.ThisAddIn.Application.ActiveDocument.FullName;
            string localpath = GetLocalPath(path);
            if (localpath == null) { MessageBox.Show("Não foi possível gerar o PDF."); return; }
            string path_pdf = localpath.Substring(0, localpath.LastIndexOf(".")) + ".pdf";
            //Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(localpath.Substring(0, localpath.LastIndexOf(".")), WdExportFormat.wdExportFormatPDF, UseISO19005_1: true);

            //try { Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(localpath.Substring(0, localpath.LastIndexOf(".")), WdExportFormat.wdExportFormatPDF, UseISO19005_1: true); } catch (COMException ex) { MessageBox.Show("O PDF está aberto. Feche-o para gerar um novo PDF."); return; }
            Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(Path.Combine(Path.GetTempPath(),"tmp_pdf_PeriTAB"), WdExportFormat.wdExportFormatPDF, UseISO19005_1: true);


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
                        goto del_temp;
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
                            goto del_temp;
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

                //PdfReader inputPdf = new PdfReader(path_pdf);
                //PdfReader inputPdf = new PdfReader(Path.Combine(Path.GetTempPath(), "tmp_pdf_PeriTAB.pdf"));
                inputPdf = new PdfReader(Path.Combine(Path.GetTempPath(), "tmp_pdf_PeriTAB.pdf"));
                inputPdf_open = true;

                FileStream signedPdf = null;
                try { signedPdf = new FileStream(path_pdf_assinado, FileMode.Create); } catch (IOException ex) { MessageBox.Show("O PDF está aberto. Feche-o para gerar um novo PDF."); goto del_temp; }

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

                try { MakeSignature.SignDetached(signatureAppearance, externalSignature, chain, null, null, null, 0, CryptoStandard.CMS); } catch (CryptographicException ex)
                {
                    //Cancelamento da senha do token
                    signedPdf.Close();
                    File.Delete(path_pdf_assinado);
                    goto del_temp; 
                }
                //inputPdf.Close();
                pdfStamper.Close();

                if (File.Exists(path_pdf_assinado))
                {
                    Globals.ThisAddIn.Application.DisplayStatusBar = true; Globals.ThisAddIn.Application.StatusBar = "PDF gerado com sucesso.";
                    //if (File.Exists(path_pdf)) { File.Delete(path_pdf); }
                    if (Globals.Ribbons.Ribbon1.checkBox_abrir.Checked) { System.Diagnostics.Process.Start(path_pdf_assinado); }
                }
                else
                {
                    Globals.ThisAddIn.Application.DisplayStatusBar = true; Globals.ThisAddIn.Application.StatusBar = "A geração do PDF falhou.";
                }
            }
            else
            {
                if (File.Exists(Path.Combine(Path.GetTempPath(), "tmp_pdf_PeriTAB.pdf")))
                {
                    if (File.Exists(path_pdf)) { try { File.Delete(path_pdf); } catch (IOException ex) { MessageBox.Show("O PDF está aberto. Feche-o para gerar um novo PDF."); goto del_temp; } }
                    File.Move(Path.Combine(Path.GetTempPath(), "tmp_pdf_PeriTAB.pdf"), path_pdf);
                    Globals.ThisAddIn.Application.DisplayStatusBar = true; Globals.ThisAddIn.Application.StatusBar = "PDF gerado com sucesso.";
                    if (Globals.Ribbons.Ribbon1.checkBox_abrir.Checked) { System.Diagnostics.Process.Start(path_pdf); }
                }
                else
                { 
                    MessageBox.Show("Não foi possível gerar o PDF.");
                }

            }
        del_temp:
            if (inputPdf_open) inputPdf.Close();
            if (File.Exists(Path.Combine(Path.GetTempPath(), "tmp_pdf_PeriTAB.pdf"))) { File.Delete(Path.Combine(Path.GetTempPath(), "tmp_pdf_PeriTAB.pdf")); } //Deleta tmp.pdf
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
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
            {
                if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                {
                    InlineShape imagem = ishape;

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
                }
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private string get_text(string texto, string inicio = null, string fim = null ) //Retona a primeira ocorrência de string entre os strings 'inicio' e 'fim' no string 'texto'.
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

        private void button_destaca_imagem_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            foreach (InlineShape ishape in Globals.ThisAddIn.Application.Selection.InlineShapes)
            {
                if (ishape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture | ishape.Type == WdInlineShapeType.wdInlineShapePicture)
                {
                    ishape.Line.Visible= MsoTriState.msoTrue;
                    ishape.Line.Weight = 3;
                    ishape.Line.ForeColor.RGB = Color.FromArgb(0,255,255).ToArgb();
                }
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void button_confere_preambulo_Click(object sender, RibbonControlEventArgs e)
        {
            string localpath = GetLocalPath(Globals.ThisAddIn.Application.ActiveDocument.Path);
            string download_path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");

            string num_laudo = null;
            string ano_laudo = null;
            string unidade_laudo = null;


            for (int i = 1; i <= Globals.ThisAddIn.Application.ActiveDocument.Paragraphs.Count; i++)
            {
                string t = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[i].Range.Text;
                string t_mod = t.ToLower().Replace(" ", "").Replace(((char)160).ToString(), "").Replace(((char)9).ToString(), "").Replace(((char)8211).ToString(), "-").Replace(((char)176).ToString(), "º"); //elimina espaços, espaços inquebráveis e tabs. Ainda troca en-dash por hifen e grau por 'o' sobrescrito.
                //MessageBox.Show(t_mod);
                //if (t_mod == ((char)13).ToString()) { continue; } //

                //string result = "";
                //foreach (char c in t_trim) { result += (int)c + " "; }
                //MessageBox.Show(result);

                if (t_mod.Length > 10)
                {
                    if ((t_mod.Substring(0, 6)).ToLower() == "laudon")
                    {
                        num_laudo = get_text(t_mod, "nº", "/");
                        ano_laudo = get_text(t_mod, "/", "-");
                        unidade_laudo = get_text(t_mod, "-");
                        break;
                        //try { unidade_laudo = t_trim.ToLower().Substring(t_trim.ToLower().IndexOf("- ") + 2); } catch { unidade_laudo = null; }
                    }
                }
            }
            //MessageBox.Show(num_laudo + " " + ano_laudo + " " + unidade_laudo);
            if (num_laudo == null | ano_laudo == null | unidade_laudo == null) { MessageBox.Show("Referência do laudo não encontrada."); return; }

            string asap_path = Path.Combine(localpath, "AsAP_Laudo_" + num_laudo + "-" + ano_laudo + ".asap");
            string asap_downloads_path = Path.Combine(download_path, "AsAP_Laudo_" + num_laudo + "-" + ano_laudo + ".asap");

            // Move o arquivo ASAP de downloads.
            if (File.Exists(asap_downloads_path) & !File.Exists(asap_path))
            {
                File.Move(asap_downloads_path, asap_path);
            }

            if (File.Exists(asap_path))
            {
                string ASAP = File.ReadAllText(asap_path);
                confere_preambulo(ASAP);
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



            }

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

        private void confere_preambulo(string asap)
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
            string num_ipl = get_text(asap, "NUMERO_IPL=", "\n").Replace("IPL", "Inquérito Policial nº").Replace("RDF","Registro de Fato nº").Replace("RE", "Registro Especial nº");
            string documento = get_text(asap, "DOCUMENTO=", "\n").Replace("Of" + (char)65533 + "cio","Ofício nº"); //caracter desconhecido: losando com interrogação
            string data_documento = get_text(asap, "DATA_DOCUMENTO=", "\n");
            string num_sei = get_text(asap, "NUMERO_SIAPRO=", "\n");
            string registro = get_text(asap, "NUMERO_CRIMINALISTICA=", "\n");
            string data_registro = get_text(asap, "DATA_CRIMINALISTICA=", "\n");

            MessageBox.Show(subtitulo + " " + unidade + " " + data + " " + perito1 + " " + perito2 + " " + num_ipl + " " + documento + " " + data_documento + " " + num_sei + " " + registro + " " + data_registro);

            string preambulo = null;
            for (int i = 1; i <= Globals.ThisAddIn.Application.ActiveDocument.Paragraphs.Count; i++)
            {
                string t = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[i].Range.Text;
                string t_trim = t.Trim();
                if (t_trim.Length > 200)
                {
                    if ((t_trim.Substring(0, 2)).ToLower() == "em")
                    {
                        preambulo = t;
                        break;
                    }
                }
            }
            if (preambulo == null) { MessageBox.Show("preambulo não encontrado."); return; }
            //MessageBox.Show(preambulo);
            //string nome_unidade = "Superintendência Regional de Polícia Federal no Maranhão";
            string nome_unidade = unidade_extenso(get_text(unidade,"/"));
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
            string nome_perito2 = get_text(perito2, inicio: null, " (").Replace("()","");
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

            

            //string dia = "10";
            //string mes = "abril";
            //string ano = "3333";
            

            string preambulo_modelo1 = "Em " + dia + " de " + mes + " de " + ano + ", designad" + ao_sexo_peritos + s_peritos + " pel" + ao_sexo_chefe + " " + cargo_chefe + ", o" + s_peritos + " Perit" + ao_sexo_peritos + s_peritos + " Crimina" + is_criminais + " Federa" + is_criminais + " " + nome_perito1 + nome_perito2 + " elab" + Elaboraram_oraram + " o presente Laudo de Perícia Criminal Federal, no interesse do " + num_ipl + ", a fim de atender ao contido n" + ao_documento + " " + documento + " de " + data_documento + ", protocolado no SEI sob o nº " + num_sei + " e registrado no SISCRIM sob o nº " + registro + ", em " + data_registro + ", descrevendo com verdade e com todas as circunstâncias tudo quanto possa interessar à Justiça e respondendo aos quesitos formulados, abaixo transcritos:";
            string preambulo_modelo2 = preambulo_modelo1.Replace("respondendo aos quesitos formulados, abaixo transcritos", "atendendo ao abaixo transcrito");
            MessageBox.Show(preambulo_modelo2);




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

                    switch (get_text(un, "DPF/","/"))
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