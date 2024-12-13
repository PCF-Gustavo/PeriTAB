using iTextSharp.text.pdf;
using iTextSharp.text.pdf.security;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Org.BouncyCastle.Security;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Tarefa = System.Threading.Tasks.Task;
using X509Certificate = Org.BouncyCastle.X509.X509Certificate;

namespace PeriTAB
{
    public partial class Ribbon
    {
        private async void button_abre_SISCRIM_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;

            await Tarefa.Run(() =>
            {
                // Configurações iniciais
                Stopwatch stopwatch = new Stopwatch(); if (Variables.debugging) { stopwatch.Start(); } // Inicia o cronômetro para medir o tempo de execução da Thread
                bool success = true;
                string msg_StatusBar = "";
                string msg_Falha = "";


                string[] identificadores_laudo = pega_identificadores_laudo();

                string num_laudo = identificadores_laudo[0];
                string ano_laudo = identificadores_laudo[1];
                string unidade_laudo = identificadores_laudo[2];

                string localpath = Globals.Ribbons.Ribbon.GetLocalPath(Globals.ThisAddIn.Application.ActiveDocument.FullName);

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
                if (Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
                {
                    stopwatch.Stop();
                    msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                }
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
                if (!success && msg_Falha != "") MessageBox.Show(msg_Falha, "Abre SISCRIM");

                // Configurações finais
                //Globals.ThisAddIn.Application.ScreenUpdating = true;

                /*}).Start();*/
            });

            RibbonButton.Image = Properties.Resources.subir2;
            RibbonButton.Enabled = true;
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
        private async void button_renomeia_documento_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;

            await Tarefa.Run(() =>
            {
                // Configurações iniciais
                Stopwatch stopwatch = new Stopwatch(); if (Variables.debugging) { stopwatch.Start(); } // Inicia o cronômetro para medir o tempo de execução da Thread
                bool success = true;
                string msg_StatusBar = "";
                string msg_Falha = "";

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
                if (Variables.debugging) { stopwatch.Stop(); }
                nome_doc = Microsoft.VisualBasic.Interaction.InputBox("Novo nome do documento:", "", nome_doc_antigo.Substring(0, nome_doc_antigo.LastIndexOf(".")));
                if (Variables.debugging) { stopwatch.Start(); }

                // Expressão regular para validar nome de arquivo no Windows
                string regex_Windows = @"^[^\\\/\:\*\?\""<>\|]+$";

                // Usa Regex.IsMatch para validar o nome do arquivo
                //bool nomeValido = 

                //MessageBox.Show(nome_doc);
                if (nome_doc == "") { success = false; }
                //else if (nome_doc == null) { success = false; }
                else if (/*nome_doc == "" || */!Regex.IsMatch(nome_doc, regex_Windows) || string.IsNullOrWhiteSpace(nome_doc))
                {
                    //MessageBox.Show("ok");
                    //return;
                    success = false;
                    msg_Falha = "Nome inválido.";
                }
                else if (nome_doc == nome_doc_antigo.Substring(0, nome_doc_antigo.LastIndexOf("."))) { }
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


                        foreach (var process in Process.GetProcessesByName("WINWORD"))
                        {
                            string windowTitle = process.MainWindowTitle;
                            string processId = process.Id.ToString();
                            MessageBox.Show($"Processo: {process.ProcessName} | ID: {processId} | Título da Janela: {windowTitle}");
                            if (process.MainWindowTitle.Contains(nome_doc_antigo))
                            {
                                process.Kill(); // Força o encerramento do processo específico
                            }
                        }
                        try { File.Delete(nome_doc_completo); } catch { MessageBox.Show("Falha ao deletar o documento antigo 7."); }

                    }
                }

                // Mensagens da Thread
                if (success) { msg_StatusBar = "Renomeia documento: Sucesso"; } else { msg_StatusBar = "Renomeia documento: Falha"; }
                if (Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
                {
                    stopwatch.Stop();
                    msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                }
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
                if (!success && msg_Falha != "") MessageBox.Show(msg_Falha, "Renomeia documento");

                // Configurações finais
                //Globals.ThisAddIn.Application.ScreenUpdating = true;

                /*}).Start();*/
            });
            RibbonButton.Image = Properties.Resources.abc;
            RibbonButton.Enabled = true;
        }

        private async void button_gerar_pdf_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;
            //Globals.ThisAddIn.Application.DisplayStatusBar = false;

            await Tarefa.Run(() =>
            {

                // Configurações iniciais
                Stopwatch stopwatch = new Stopwatch(); if (Variables.debugging) { stopwatch.Start(); } // Inicia o cronômetro para medir o tempo de execução da Thread
                bool success = true;
                string msg_StatusBar = "";
                string msg_Falha = "";


                //iClass_Buttons.button_gera_pdf_image(load: true);
                PdfReader inputPdf = null;
                bool inputPdf_open = false;
                string path = Globals.ThisAddIn.Application.ActiveDocument.FullName;
                string localpath = GetLocalPath(path);
                if (localpath == null)
                {
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

                if (Globals.Ribbons.Ribbon.checkBox_assinar.Checked)
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
                            if (Variables.debugging) { stopwatch.Stop(); }
                            X509Certificate2Collection collection = X509Certificate2UI.SelectFromCollection(st.Certificates, "Escolha o certificado:", "", X509SelectionFlag.SingleSelection);
                            if (Variables.debugging) { stopwatch.Start(); }
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
                    try
                    {
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

                    //RSACryptoServiceProvider rsa = certClient.PrivateKey as RSACryptoServiceProvider;*********************************************************
                    //rsa.PersistKeyInCsp = false;

                    //rsa2.PersistKeyInCsp = true;
                    //MessageBox.Show(rsa2.PersistKeyInCsp.ToString());

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

                    //if (Globals.Ribbons.Ribbon.checkBox_senha.Checked)
                    //{
                    //    //(new RSACryptoServiceProvider()).PersistKeyInCsp = true; //Define chave persistente. Só pede a senha da primeira vez.
                    //    rsa.PersistKeyInCsp = true;
                    //    //MessageBox.Show("1");
                    //    //rsa.Clear();
                    //}
                    //if (!Globals.Ribbons.Ribbon.checkBox_senha.Checked)
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
                    //certClient.Dispose();





                    if (File.Exists(path_pdf_assinado))
                    {
                        //iClass_Buttons.muda_imagem("button_gera_pdf", Properties.Resources.icone_pdf2);
                        //button_gera_pdf.Enabled = true;
                        //Globals.ThisAddIn.Application.DisplayStatusBar = true;
                        //Globals.ThisAddIn.Application.StatusBar = "PDF gerado com sucesso.";
                        //if (File.Exists(path_pdf)) { File.Delete(path_pdf); }
                        if (Globals.Ribbons.Ribbon.checkBox_abrir.Checked) { System.Diagnostics.Process.Start(path_pdf_assinado); }
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
                        if (File.Exists(path_pdf))
                        {
                            try { File.Delete(path_pdf); }
                            catch (IOException)
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
                        if (Globals.Ribbons.Ribbon.checkBox_abrir.Checked) { System.Diagnostics.Process.Start(path_pdf); }
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
                if (Variables.debugging) // Se estiver no modo Debugging, mostra o tempo de execução na barra de status
                {
                    stopwatch.Stop();
                    msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                }
                Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
                if (!success && msg_Falha != "") MessageBox.Show(msg_Falha, "Gera PDF");

                // Configurações finais
                //Globals.ThisAddIn.Application.ScreenUpdating = true;


                /*}).Start();*/
            });

            RibbonButton.Image = Properties.Resources.icone_pdf;
            RibbonButton.Enabled = true;
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
    }
}