using iTextSharp.text.pdf;
using iTextSharp.text.pdf.security;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Org.BouncyCastle.Security;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Threading.Tasks;
using X509Certificate = Org.BouncyCastle.X509.X509Certificate;
using Task = System.Threading.Tasks.Task;

namespace PeriTAB
{
    public partial class Ribbon
    {
        private async void Button_renomeia_documento_Click(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon_com_UI_responsiva(sender, e, async progress =>
            {

                string nome_doc_completo = Globals.ThisAddIn.Application.ActiveDocument.FullName;
                string caminho_doc = Globals.ThisAddIn.Application.ActiveDocument.Path;
                string nome_doc_antigo = Globals.ThisAddIn.Application.ActiveDocument.Name;
                string nome_doc = null;

                nome_doc_completo = GetLocalPath(nome_doc_completo);
                nome_doc = Microsoft.VisualBasic.Interaction.InputBox("Novo nome do documento:", "", nome_doc_antigo.Substring(0, nome_doc_antigo.LastIndexOf(".")));

                // Expressão regular para validar nome de arquivo no Windows
                string regex_Windows = @"^[^\\\/\:\*\?\""<>\|]+$";

                if (nome_doc == "") { throw new Exception(""); }
                else if (!Regex.IsMatch(nome_doc, regex_Windows) || string.IsNullOrWhiteSpace(nome_doc)) throw new Exception("Nome inválido.");
                else if (nome_doc == nome_doc_antigo.Substring(0, nome_doc_antigo.LastIndexOf("."))) { }
                else
                {
                    Globals.ThisAddIn.Application.ActiveDocument.SaveAs2(FileName: Path.Combine(caminho_doc, nome_doc + ".docx"), FileFormat: WdSaveFormat.wdFormatDocumentDefault);

                    try { File.Delete(nome_doc_completo); }
                    catch { Variables.Lista_arquivos_para_excluir.Add(nome_doc_completo); }
                }
                await Task.CompletedTask;
            });
        }

        private async void Button_gerar_pdf_Click(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon_com_UI_responsiva(sender, e, async progress =>
            {
                PdfReader inputPdf = null;
                bool inputPdf_open = false;

                string path = Globals.ThisAddIn.Application.ActiveDocument.FullName;
                string localpath = GetLocalPath(path);

                if (localpath == null) throw new Exception("Não foi possível gerar o PDF.");

                string path_pdf = Path.ChangeExtension(localpath, ".pdf");
                string path_pdf_assinado = Path.ChangeExtension(localpath, null) + "_assinado.pdf";


                string path_tmp_pdf = Path.Combine(Path.GetTempPath(), "tmp_pdf_PeriTAB.pdf");
                string path_tmp_pdf_assinado = Path.Combine(Path.GetTempPath(), "tmp_pdf_assinado_PeriTAB.pdf");

                try
                {
                    Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(
                        path_tmp_pdf,
                        WdExportFormat.wdExportFormatPDF,
                        UseISO19005_1: true);

                    if (Globals.Ribbons.Ribbon.CheckBox_assinar.Checked)
                    {
                        bool assinado = Assina_PDF(path_tmp_pdf, path_tmp_pdf_assinado);

                        if (!assinado) return;

                        Substitui_PDF_Final(path_tmp_pdf_assinado,path_pdf_assinado);

                        if (Globals.Ribbons.Ribbon.CheckBox_abrir.Checked) System.Diagnostics.Process.Start(path_pdf_assinado);
                    }
                    else
                    {
                        Substitui_PDF_Final(path_tmp_pdf, path_pdf);

                        if (Globals.Ribbons.Ribbon.CheckBox_abrir.Checked) System.Diagnostics.Process.Start(path_pdf);
                    }

                    
                }
                finally
                {
                    if (inputPdf_open)
                        inputPdf.Close();

                    if (File.Exists(path_tmp_pdf))
                        File.Delete(path_tmp_pdf);

                    if (File.Exists(path_tmp_pdf_assinado))
                        File.Delete(path_tmp_pdf_assinado);
                }
                await Task.CompletedTask;
            });
        }

        private static bool Assina_PDF(string pathPdfEntrada, string pathPdfSaida)
        {
            X509Certificate2 certClient = null;

            // ================= SELEÇÃO DO CERTIFICADO =================
            using (X509Store st = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                st.Open(OpenFlags.MaxAllowed);

                List<X509Certificate2> certificadosValidos = new List<X509Certificate2>();

                foreach (X509Certificate2 c in st.Certificates)
                {
                    if (!c.HasPrivateKey) continue;
                    if (!c.Verify()) continue;
                    if (!Checar_Se_ICP_Brasil(c)) continue;

                    certificadosValidos.Add(c);
                }

                if (certificadosValidos.Count == 0)
                    throw new Exception("Não foi possível gerar o PDF.");

                if (certificadosValidos.Count == 1)
                {
                    certClient = certificadosValidos[0];
                }
                else
                {
                    X509Certificate2Collection collection =
                        X509Certificate2UI.SelectFromCollection(
                            new X509Certificate2Collection(certificadosValidos.ToArray()),
                            "Escolha o certificado:",
                            "",
                            X509SelectionFlag.SingleSelection);

                    if (collection.Count == 0) throw new Exception("Nenhum certificado válido encontrado.");

                    certClient = collection[0];
                }
            }

            // ================= CADEIA DE CERTIFICAÇÃO =================
            IList<X509Certificate> chain = new List<X509Certificate>();
            X509Chain x509Chain = new X509Chain();
            x509Chain.Build(certClient);

            foreach (X509ChainElement el in x509Chain.ChainElements)
                chain.Add(DotNetUtilities.FromX509Certificate(el.Certificate));

            // ================= ASSINATURA DO PDF =================
            PdfReader inputPdf = null;
            PdfStamper pdfStamper = null;
            bool assinaturaConcluida = false;

            try
            {
                inputPdf = new PdfReader(pathPdfEntrada);

                using (FileStream signedPdf = new FileStream(pathPdfSaida, FileMode.Create))
                {
                    pdfStamper = PdfStamper.CreateSignature(inputPdf, signedPdf, '\0');

                    try
                    {
                        IExternalSignature externalSignature =
                            new X509Certificate2Signature(certClient, "SHA-256");

                        MakeSignature.SignDetached(
                            pdfStamper.SignatureAppearance,
                            externalSignature,
                            chain,
                            null,
                            null,
                            null,
                            0,
                            CryptoStandard.CMS);

                        assinaturaConcluida = true;
                        return true;
                    }
                    catch (CryptographicException)
                    {
                        return false;
                    }
                }
            }
            finally
            {
                if (assinaturaConcluida && pdfStamper != null)
                    pdfStamper.Close();

                if (inputPdf != null)
                    inputPdf.Close();
            }
        }

        private static void Substitui_PDF_Final(string pathTmp, string pathFinal)
        {
            if (!File.Exists(pathTmp))
                throw new Exception("Não foi possível gerar o PDF.");

            if (File.Exists(pathFinal))
            {
                try
                {
                    File.Delete(pathFinal);
                }
                catch (IOException)
                {
                    throw new Exception("O PDF está aberto. Feche-o para gerar um novo PDF.");
                }
            }

            File.Move(pathTmp, pathFinal);

        }


        private static bool Checar_Se_ICP_Brasil(X509Certificate2 cert)
        {
            using (var chain = new X509Chain())
            {
                chain.ChainPolicy.RevocationMode = X509RevocationMode.NoCheck;
                chain.ChainPolicy.VerificationFlags = X509VerificationFlags.NoFlag;

                if (!chain.Build(cert))
                    return false;

                // Root CA = último elemento da cadeia
                X509Certificate2 root = chain.ChainElements
                    [chain.ChainElements.Count - 1]
                    .Certificate;

                string cnRoot = root.GetNameInfo(X509NameType.SimpleName, false);

                return cnRoot.StartsWith(
                    "Autoridade Certificadora Raiz Brasileira",
                    StringComparison.OrdinalIgnoreCase
                );
            }
        }

        //private async void Button_gerar_pdf_Click(object sender, RibbonControlEventArgs e)
        //{
        //    // Atualiza a UI na Thread principal
        //    RibbonButton RibbonButton = (RibbonButton)sender;
        //    RibbonButton.Image = Properties.Resources.load_icon_png_7969;
        //    RibbonButton.Enabled = false;

        //    await Task.Run(() =>
        //    {
        //        // Configurações iniciais
        //        bool success = true;
        //        string msg_StatusBar = "";
        //        string msg_Falha = "";

        //        PdfReader inputPdf = null;
        //        bool inputPdf_open = false;
        //        string path = Globals.ThisAddIn.Application.ActiveDocument.FullName;
        //        string localpath = GetLocalPath(path);
        //        if (localpath == null)
        //        {
        //            success = false;
        //            msg_Falha = "Não foi possível gerar o PDF.";
        //            goto saida;
        //            //iClass_Buttons.muda_imagem("button_gera_pdf", Properties.Resources.icone_pdf2); MessageBox.Show("Não foi possível gerar o PDF."); button_gera_pdf.Enabled = true; return; 
        //        }
        //        string path_pdf = localpath.Substring(0, localpath.LastIndexOf(".")) + ".pdf";
        //        string path_pdf_assinado = localpath.Substring(0, localpath.LastIndexOf(".")) + "_assinado.pdf";
        //        string path_tmp_pdf = Path.Combine(Path.GetTempPath(), "tmp_pdf_PeriTAB.pdf");
        //        string path_tmp_pdf_assinado = Path.Combine(Path.GetTempPath(), "tmp_pdf_assinado_PeriTAB.pdf");

        //        Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(path_tmp_pdf, WdExportFormat.wdExportFormatPDF, UseISO19005_1: true);

        //        if (Globals.Ribbons.Ribbon.CheckBox_assinar.Checked)
        //        {
        //            X509Certificate2 certClient = null;
        //            X509Store st = new X509Store(StoreName.My, StoreLocation.CurrentUser);
        //            st.Open(OpenFlags.MaxAllowed);
        //            //IExternalSignature s = null;

        //            List<X509Certificate2> certificadosValidos = new List<X509Certificate2>();
        //            foreach (X509Certificate2 c in st.Certificates)
        //            {
        //                if (!c.HasPrivateKey) continue;
        //                if (!c.Verify()) continue;
        //                //try { _ = new X509Certificate2Signature(c, "SHA-256"); } catch { continue; }

        //                //MessageBox.Show(c.GetNameInfo(X509NameType.SimpleName, true));

        //                //if (c.GetNameInfo(X509NameType.SimpleName, true).StartsWith("PF SUBORDINATE CA", StringComparison.OrdinalIgnoreCase)) continue;

        //                if (!CertificadoEhICPBrasil(c)) continue;

        //                certificadosValidos.Add(c);

        //                //if (c.Verify() == false) { st.Remove(c); continue; } //Elimina certificado não validados
        //                //try { s = new X509Certificate2Signature(c, "SHA-256"); } catch { st.Remove(c); } //Elimina certificado que não se pode pegar a assinatura
        //            }
        //            //switch (st.Certificates.Count)
        //            switch (certificadosValidos.Count)
        //            {
        //                case 0:
        //                    success = false;
        //                    msg_Falha = "Nenhum certificado válido encontrado.";
        //                    goto saida;
        //                case 1:
        //                    //certClient = st.Certificates[0];
        //                    certClient = certificadosValidos[0];
        //                    break;
        //                default:
        //                    //X509Certificate2Collection collection = X509Certificate2UI.SelectFromCollection(st.Certificates, "Escolha o certificado:", "", X509SelectionFlag.SingleSelection);
        //                    X509Certificate2Collection collection = X509Certificate2UI.SelectFromCollection(new X509Certificate2Collection(certificadosValidos.ToArray()), "Escolha o certificado:", "", X509SelectionFlag.SingleSelection);
        //                    if (collection.Count > 0)
        //                    {
        //                        certClient = collection[0];
        //                    }
        //                    else
        //                    {
        //                        success = false;
        //                        goto saida;
        //                    }
        //                    break;
        //            }
        //            st.Close();

        //            IList<X509Certificate> chain = new List<X509Certificate>();

        //            X509Chain x509Chain = new X509Chain();
        //            x509Chain.Build(certClient);

        //            foreach (X509ChainElement x509ChainElement in x509Chain.ChainElements)
        //            {
        //                chain.Add(DotNetUtilities.FromX509Certificate(x509ChainElement.Certificate));
        //            }

        //            inputPdf = new PdfReader(path_tmp_pdf);
        //            inputPdf_open = true;

        //            //FileStream signedPdf = null;
        //            //try
        //            //{
        //            //signedPdf = new FileStream(path_pdf_assinado, FileMode.Create);
        //            FileStream signedPdf = new FileStream(path_tmp_pdf_assinado, FileMode.Create);
        //            //}
        //            //catch (IOException)
        //            //{
        //            //    success = false;
        //            //    msg_Falha = "O PDF está aberto. Feche-o para gerar um novo PDF.";
        //            //    goto saida;
        //            //}

        //            PdfStamper pdfStamper = PdfStamper.CreateSignature(inputPdf, signedPdf, '\0');

        //            // Desativa a persistência da chave no CSP, garantindo que a senha seja solicitada sempre
        //            //RSACryptoServiceProvider rsa = (RSACryptoServiceProvider)certClient.PrivateKey;
        //            //rsa.PersistKeyInCsp = false; // Força a solicitação da senha

        //            //RSACryptoServiceProvider rsa2 = new RSACryptoServiceProvider();

        //            //RSACryptoServiceProvider rsa = certClient.PrivateKey as RSACryptoServiceProvider;*********************************************************
        //            //rsa.PersistKeyInCsp = false;

        //            //rsa2.PersistKeyInCsp = true;
        //            //MessageBox.Show(rsa2.PersistKeyInCsp.ToString());

        //            IExternalSignature externalSignature = new X509Certificate2Signature(certClient, "SHA-256");

        //            PdfSignatureAppearance signatureAppearance = pdfStamper.SignatureAppearance;





        //            //signatureAppearance.SignatureGraphic = Image.GetInstance(pathToSignatureImage);
        //            //signatureAppearance.SetVisibleSignature(new iTextSharp.text.Rectangle(0, 00, 250, 150), inputPdf.NumberOfPages, "Signature");
        //            //signatureAppearance.SignatureRenderingMode = PdfSignatureAppearance.RenderingMode.DESCRIPTION;


        //            //RSACryptoServiceProvider rsa = new RSACryptoServiceProvider();
        //            //CspParameters cspp = new CspParameters();
        //            //cspp.KeyContainerName = rsa.CspKeyContainerInfo.KeyContainerName;
        //            //cspp.ProviderName = rsa.CspKeyContainerInfo.ProviderName;
        //            //cspp.ProviderType = rsa.CspKeyContainerInfo.ProviderType;
        //            //cspp.Flags = CspProviderFlags.NoPrompt;
        //            //RSACryptoServiceProvider rsa2 = new RSACryptoServiceProvider(cspp);
        //            //rsa.PersistKeyInCsp = true;

        //            //(new RSACryptoServiceProvider()).PersistKeyInCsp = true; //Define chave persistente. Só pede a senha da primeira vez.

        //            //RSACryptoServiceProvider rsa = new RSACryptoServiceProvider();
        //            //rsa.PersistKeyInCsp = false;

        //            //(new RSACryptoServiceProvider()).PersistKeyInCsp = false;

        //            //CspParameters cspp = new CspParameters();
        //            //cspp.KeyContainerName = "MyKeyContainer";
        //            //RSACryptoServiceProvider rsa = new RSACryptoServiceProvider(cspp);

        //            //if (Globals.Ribbons.Ribbon.checkBox_senha.Checked)
        //            //{
        //            //    //(new RSACryptoServiceProvider()).PersistKeyInCsp = true; //Define chave persistente. Só pede a senha da primeira vez.
        //            //    rsa.PersistKeyInCsp = true;
        //            //    //MessageBox.Show("1");
        //            //    //rsa.Clear();
        //            //}
        //            //if (!Globals.Ribbons.Ribbon.checkBox_senha.Checked)
        //            //{
        //            //    //(new RSACryptoServiceProvider()).PersistKeyInCsp = false;
        //            //    rsa.PersistKeyInCsp = false;
        //            //    //MessageBox.Show("2");
        //            //    rsa.Clear();
        //            //}


        //            try
        //            {
        //                MakeSignature.SignDetached(signatureAppearance, externalSignature, chain, null, null, null, 0, CryptoStandard.CMS);
        //            }
        //            catch (CryptographicException)
        //            {
        //                //Cancelamento da senha do token
        //                signedPdf.Close();
        //                File.Delete(path_tmp_pdf_assinado);
        //                success = false;
        //                goto saida;
        //            }
        //            pdfStamper.Close();

        //            if (File.Exists(path_tmp_pdf_assinado))
        //            {
        //                if (File.Exists(path_pdf_assinado))
        //                {
        //                    try { File.Delete(path_pdf_assinado); }
        //                    catch (IOException)
        //                    {
        //                        success = false;
        //                        msg_Falha = "O PDF está aberto. Feche-o para gerar um novo PDF.";
        //                        goto saida;
        //                    }
        //                }
        //                File.Move(path_tmp_pdf_assinado, path_pdf_assinado);
        //                if (Globals.Ribbons.Ribbon.CheckBox_abrir.Checked) { System.Diagnostics.Process.Start(path_pdf_assinado); }
        //            }
        //            else
        //            {
        //                success = false;
        //                msg_Falha = "Não foi possível gerar o PDF.";
        //                goto saida;
        //            }
        //        }
        //        else
        //        {
        //            if (File.Exists(path_tmp_pdf))
        //            {
        //                if (File.Exists(path_pdf))
        //                {
        //                    try { File.Delete(path_pdf); }
        //                    catch (IOException)
        //                    {
        //                        success = false;
        //                        msg_Falha = "O PDF está aberto. Feche-o para gerar um novo PDF.";
        //                        goto saida;
        //                    }
        //                }
        //                File.Move(path_tmp_pdf, path_pdf);
        //                if (Globals.Ribbons.Ribbon.CheckBox_abrir.Checked) { System.Diagnostics.Process.Start(path_pdf); }
        //            }
        //            else
        //            {
        //                success = false;
        //                msg_Falha = "Não foi possível gerar o PDF.";
        //                goto saida;
        //            }
        //        }

        //    saida:
        //        if (inputPdf_open) inputPdf.Close();
        //        if (File.Exists(Path.Combine(Path.GetTempPath(), "tmp_pdf_PeriTAB.pdf"))) { File.Delete(Path.Combine(Path.GetTempPath(), "tmp_pdf_PeriTAB.pdf")); } //Deleta tmp.pdf
        //        if (File.Exists(Path.Combine(Path.GetTempPath(), "tmp_pdf_assinado_PeriTAB.pdf"))) { File.Delete(Path.Combine(Path.GetTempPath(), "tmp_pdf_assinado_PeriTAB.pdf")); } //Deleta tmp.pdf asinado

        //        // Mensagens da Thread
        //        if (success) { msg_StatusBar = "Gera PDF: Sucesso"; } else { msg_StatusBar = "Gera PDF: Falha"; }
        //        Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
        //        if (!success && msg_Falha != "") MessageBox.Show(msg_Falha, "Gera PDF");

        //    });

        //    RibbonButton.Image = Properties.Resources.icone_pdf;
        //    RibbonButton.Enabled = true;
        //}

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

        //private string Get_text(string texto, string inicio = null, string fim = null) //Retona a primeira ocorrência de string entre os strings 'inicio' e 'fim' no string 'texto'.
        //{
        //    if (inicio == null & fim == null) { return null; }

        //    try
        //    {
        //        if (inicio == null)
        //        {
        //            return texto.Substring(0, texto.IndexOf(fim));
        //        }
        //        if (fim == null)
        //        {
        //            return texto.Substring(texto.IndexOf(inicio) + inicio.Length);
        //        }
        //        return (texto.Substring(texto.IndexOf(inicio))).Substring(inicio.Length, (texto.Substring(texto.IndexOf(inicio) + inicio.Length)).IndexOf(fim));
        //    }
        //    catch { return null; }
        //}
    }
}