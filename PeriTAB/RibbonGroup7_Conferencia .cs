using iTextSharp.text.pdf.security;
using iTextSharp.text.pdf;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Windows.Forms;
using Org.BouncyCastle.Security;
using System.Security.Cryptography.X509Certificates;
using X509Certificate = Org.BouncyCastle.X509.X509Certificate;
using System.Text;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Tarefa = System.Threading.Tasks.Task;
using System.Text.RegularExpressions;


namespace PeriTAB
{
    public partial class Ribbon
    {
        private void button_confere_formatacao_Click(object sender, RibbonControlEventArgs e)
        {
        }

        private async void button_confere_preambulo_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;

            await Tarefa.Run(() =>
            {

                string localpath = GetLocalPath(Globals.ThisAddIn.Application.ActiveDocument.Path);
                string download_path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");

                string[] identificadores_laudo = pega_identificadores_laudo();

                string num_laudo = identificadores_laudo[0];
                string ano_laudo = identificadores_laudo[1];
                string unidade_laudo = identificadores_laudo[2];

                if (num_laudo == null | ano_laudo == null | unidade_laudo == null)
                {
                    MessageBox.Show("Referência do laudo não encontrada.");
                    RibbonButton.Image = Properties.Resources.checklist2;
                    RibbonButton.Enabled = true;
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

                /*}).Start();*/
            });

            RibbonButton.Image = Properties.Resources.checklist2;
            RibbonButton.Enabled = true;
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


        private async void button_confere_num_legenda_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;

            // Executa as tarefas em segundo plano
            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.Run("atualiza_todos_campos");
                Globals.ThisAddIn.Application.Run("confere_numeracao_legendas");
            });

            // Após a execução das tarefas, atualiza a UI na Thread principal
            RibbonButton.Image = Properties.Resources.lupa;
            RibbonButton.Enabled = true;
        }
    }
}