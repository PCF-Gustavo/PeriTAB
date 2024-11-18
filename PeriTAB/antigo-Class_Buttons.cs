using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Controls;
using static System.Net.Mime.MediaTypeNames;

namespace PeriTAB
{
    internal class Class_Buttons
    {


        public void DefaultAll()
        {
            button_confere_num_legenda_Default();
            button_alinha_legenda_Default();
            button_atualiza_campos_Default();
            button_moeda_Default();
            button_inteiro_Default();
            //button_importa_estilos_Default();
            button_limpa_estilos_Default();
            toggleButton_painel_de_estilos_Default();
            button_cola_imagem_Default();
            checkBox_altura_Default();
            checkBox_largura_Default();
            editBox_largura_Default();
            editBox_altura_Default();
            //dropDown_ordem_Default();
            dropDown_separador_Default();
            checkBox_assinar_Default();
            checkBox_abrir_Default();
            button_renomeia_documento_Default();
            button_gera_pdf_Default();
            button_abre_SISCRIM_Default();
        }

        public void button_confere_num_legenda_Default()
        {
            Globals.Ribbons.Ribbon.button_confere_num_legenda.Enabled = true;
            Globals.Ribbons.Ribbon.button_confere_num_legenda.ScreenTip = "Macro confere_numeracao_legendas";
            Globals.Ribbons.Ribbon.button_confere_num_legenda.SuperTip = "Corrige erros de numeração em Figuras, Tabelas etc.";
        }
        public void button_alinha_legenda_Default()
        {
            Globals.Ribbons.Ribbon.button_alinha_legenda.Enabled = true;
            Globals.Ribbons.Ribbon.button_alinha_legenda.ScreenTip = "Macro alinha_legenda";
            Globals.Ribbons.Ribbon.button_alinha_legenda.SuperTip = "Alinha legenda de Figuras, Tabelas etc.";
        }

        public void button_atualiza_campos_Default()
        {
            Globals.Ribbons.Ribbon.button_atualiza_campos.Enabled = true;
            Globals.Ribbons.Ribbon.button_atualiza_campos.ScreenTip = "Macro atualiza_todos_campos";
            Globals.Ribbons.Ribbon.button_atualiza_campos.SuperTip = "Atualiza todos os campos (MS Word Fields) do documento.";
        }
        public void button_moeda_Default()
        {
            Globals.Ribbons.Ribbon.button_moeda.Enabled = true;
            Globals.Ribbons.Ribbon.button_moeda.ScreenTip = "Macro moeda_por_extenso";
            Globals.Ribbons.Ribbon.button_moeda.SuperTip = "Escreve por extenso o valor em Reais. Posicione o cursor ao final do número.";
        }
        public void button_inteiro_Default()
        {
            Globals.Ribbons.Ribbon.button_inteiro.Enabled = true;
            Globals.Ribbons.Ribbon.button_inteiro.ScreenTip = "Macro moeda_por_extenso";
            Globals.Ribbons.Ribbon.button_inteiro.SuperTip = "Escreve por extenso o número inteiro. Posicione o cursor ao final do número.";
        }
        //public void button_importa_estilos_Default()
        //{
        //    Globals.Ribbons.Ribbon.button_importa_estilos.Enabled = true;
        //    Globals.Ribbons.Ribbon.button_importa_estilos.SuperTip = "Importa Estilos de parágrafos.";
        //}
        public void button_limpa_estilos_Default()
        {
            Globals.Ribbons.Ribbon.button_limpa_estilos.Enabled = true;
            Globals.Ribbons.Ribbon.button_limpa_estilos.ScreenTip = "Macro limpa_estilos";
            Globals.Ribbons.Ribbon.button_limpa_estilos.SuperTip = "Remove Estilos de parágrafos não utilizados.";
        }

        public void toggleButton_painel_de_estilos_Default()
        {            
            Globals.Ribbons.Ribbon.toggleButton_painel_de_estilos.Checked = bool.Parse(GetPreference("painel_de_estilos")); 
            //Globals.ThisAddIn.TaskPane1.Visible = bool.Parse(preferences.painel_de_estilos);
        }

        public void button_cola_imagem_Default()
        {
            Globals.Ribbons.Ribbon.button_cola_imagem.Enabled = true;
            Globals.Ribbons.Ribbon.button_cola_imagem.ScreenTip = "";
            Globals.Ribbons.Ribbon.button_cola_imagem.SuperTip = "Cola imagens do Clipboard em ordem alfabética.";
        }

        public void checkBox_largura_Default()
        {
            Globals.Ribbons.Ribbon.checkBox_largura.Checked = bool.Parse(GetPreference("largura_checked"));
        }
        public void editBox_largura_Default()
        {
            Globals.Ribbons.Ribbon.editBox_largura.Enabled = bool.Parse(GetPreference("largura_checked"));
            //if (Globals.Ribbons.Ribbon.editBox_largura.Enabled) { Globals.Ribbons.Ribbon.editBox_largura.Text = preferences.largura; }
            if (Globals.Ribbons.Ribbon.checkBox_largura.Checked) { Globals.Ribbons.Ribbon.editBox_largura.Text = GetPreference("largura"); }

        }

        public void checkBox_altura_Default()
        {
            Globals.Ribbons.Ribbon.checkBox_altura.Checked = !bool.Parse(GetPreference("largura_checked"));
        }

        public void editBox_altura_Default()
        {
            Globals.Ribbons.Ribbon.editBox_altura.Enabled = !bool.Parse(GetPreference("largura_checked"));
            //if (Globals.Ribbons.Ribbon.checkBox_altura.Enabled) { Globals.Ribbons.Ribbon.editBox_altura.Text = preferences.altura; }
            if (Globals.Ribbons.Ribbon.checkBox_altura.Checked) { Globals.Ribbons.Ribbon.editBox_altura.Text = GetPreference("altura"); }
        }

        //public void dropDown_ordem_Default()
        //{
        //    int index = -1;
        //    if (preferences.ordem == "Alfabética") { index = 0; }
        //    if (preferences.ordem == "Seleção") { index = 1; }
        //    Globals.Ribbons.Ribbon.dropDown_ordem.SelectedItemIndex = index;
        //}

        public void dropDown_separador_Default()
        {
            int index = -1;
            if (GetPreference("separador") == "Nenhum") { index = 0; }
            if (GetPreference("separador") == "Espaço") { index = 1; }
            if (GetPreference("separador") == "Parágrafo") { index = 2; }
            if (GetPreference("separador") == "Parágrafo + 3pt") { index = 3; }
            Globals.Ribbons.Ribbon.dropDown_separador.SelectedItemIndex = index;
        }
        public void checkBox_assinar_Default()
        {
            Globals.Ribbons.Ribbon.checkBox_assinar.Checked = bool.Parse(GetPreference("assinar_pdf"));
        }
        public void checkBox_abrir_Default()
        {
            Globals.Ribbons.Ribbon.checkBox_abrir.Checked = bool.Parse(GetPreference("abrir_pdf"));
        }
        public void button_renomeia_documento_Default()
        {
            Globals.Ribbons.Ribbon.button_renomeia_documento.Enabled = true;
            Globals.Ribbons.Ribbon.button_renomeia_documento.ScreenTip = "";
            Globals.Ribbons.Ribbon.button_renomeia_documento.SuperTip = "Renomeia o documento atual.";
        }
        public void button_gera_pdf_Default()
        {
            Globals.Ribbons.Ribbon.button_gera_pdf.Image = Properties.Resources.icone_pdf_chave;
            Globals.Ribbons.Ribbon.button_gera_pdf.Enabled = true;
            Globals.Ribbons.Ribbon.button_gera_pdf.ScreenTip = "";
            Globals.Ribbons.Ribbon.button_gera_pdf.SuperTip = "Gera o PDF do documento na pasta onde está salvo.";
        }
        public void button_abre_SISCRIM_Default()
        {
            Globals.Ribbons.Ribbon.button_abre_SISCRIM.Enabled = true;
            Globals.Ribbons.Ribbon.button_abre_SISCRIM.ScreenTip = "";
            Globals.Ribbons.Ribbon.button_abre_SISCRIM.SuperTip = "Abre SISCRIM na página do Laudo ou da Requisição.";
            //Globals.Ribbons.Ribbon.button_abre_SISCRIM.Enabled = false; 
            //Globals.Ribbons.Ribbon.button_abre_SISCRIM.ScreenTip = "Desabilitado"; 
            //Globals.Ribbons.Ribbon.button_abre_SISCRIM.SuperTip = "O PDF do laudo ainda não foi gerado.";
        }

        //public void button_gera_pdf_image(bool load)
        //{
        //    if (load) Globals.Ribbons.Ribbon.button_gera_pdf.Image = Properties.Resources.load_icon_png_7969;
        //    else Globals.Ribbons.Ribbon.button_gera_pdf.Image = Properties.Resources.icone_pdf2;
        //}

        //public void muda_imagem(Microsoft.Office.Tools.Ribbon.RibbonButton botao, System.Drawing.Bitmap imagem)
        //{
        //    botao.Image = imagem;
        //}

        //public void muda_imagem(string botao, System.Drawing.Bitmap imagem)
        //{
        //    switch (botao)
        //    {
        //        case "button_atualiza_campos":
        //            Globals.Ribbons.Ribbon.button_atualiza_campos.Image = imagem;
        //            break;
        //        case "button_redimensiona_imagem":
        //            Globals.Ribbons.Ribbon.button_redimensiona_imagem.Image = imagem;
        //            break;
        //        case "button_cola_imagem":
        //            Globals.Ribbons.Ribbon.button_cola_imagem.Image = imagem;
        //            break;
        //        case "button_autodimensiona_imagem":
        //            Globals.Ribbons.Ribbon.button_autodimensiona_imagem.Image = imagem;
        //            break;
        //        case "button_confere_preambulo":
        //            Globals.Ribbons.Ribbon.button_confere_preambulo.Image = imagem;
        //            break;
        //        case "button_confere_num_legenda":
        //            Globals.Ribbons.Ribbon.button_confere_num_legenda.Image = imagem;
        //            break;
        //        case "button_abre_SISCRIM":
        //            Globals.Ribbons.Ribbon.button_abre_SISCRIM.Image = imagem;
        //            break;
        //        case "button_renomeia_documento":
        //            Globals.Ribbons.Ribbon.button_renomeia_documento.Image = imagem;
        //            break;
        //        case "button_gera_pdf":
        //            Globals.Ribbons.Ribbon.button_gera_pdf.Image = imagem;
        //            break;
        //        case "menu_inserir_imagem":
        //            Globals.Ribbons.Ribbon.menu_inserir_imagem.Image = imagem;
        //            break;
        //        case "menu_remover_imagem":
        //            Globals.Ribbons.Ribbon.menu_remover_imagem.Image = imagem;
        //            break;
        //        case "menu_formatacao_imagem":
        //            Globals.Ribbons.Ribbon.menu_formatacao_imagem.Image = imagem;
        //            break;
        //        case "menu_inserir_tabela":
        //            Globals.Ribbons.Ribbon.menu_inserir_tabela.Image = imagem;
        //            break;
        //        case "menu_remover_tabela":
        //            Globals.Ribbons.Ribbon.menu_remover_tabela.Image = imagem;
        //            break;
        //        case "menu_formatacao_tabela":
        //            Globals.Ribbons.Ribbon.menu_formatacao_tabela.Image = imagem;
        //            break;
        //        case "menu_formatacao_campos":
        //            Globals.Ribbons.Ribbon.menu_formatacao_campos.Image = imagem;
        //            break;
        //            //default:
        //            //    break;
        //    }
        //}


        //public class preferences
        //{
        //    private static string private_largura, private_altura, private_largura_checked, private_separador, private_painel_de_estilos, private_assinar_pdf, private_abrir_pdf;
        //    static preferences() // Bloco estático para definir o valor inicial das variáveis (Leitura das preferencias)
        //    {
        //        //MessageBox.Show("preferences");
        //        if (File.Exists(Ribbon.Variables.caminho_preferences))
        //        {
        //            XmlDocument xmlDoc = new XmlDocument();
        //            xmlDoc.Load(Ribbon.Variables.caminho_preferences);
        //            XmlElement root = xmlDoc.DocumentElement;

        //            // Lendo as preferências
        //            private_largura = GetElementValue(root, "largura", "10");
        //            private_altura = GetElementValue(root, "altura", "10");
        //            private_largura_checked = GetElementValue(root, "largura_checked", "true");
        //            private_separador = GetElementValue(root, "separador", "Nenhum");
        //            private_painel_de_estilos = GetElementValue(root, "painel_de_estilos", "false");
        //            private_assinar_pdf = GetElementValue(root, "assinar_pdf", "true");
        //            private_abrir_pdf = GetElementValue(root, "abrir_pdf", "true");
        //        }
        //        else
        //        {
        //            // Preferências iniciais
        //            private_largura = "10";
        //            private_altura = "10";
        //            private_largura_checked = "true";
        //            private_separador = "Nenhum";
        //            private_painel_de_estilos = "false";
        //            private_assinar_pdf = "true";
        //            private_abrir_pdf = "true";
        //        }

        //        // Ajuste inicial dos valores de editBox_largura_Text e editBox_altura_Text
        //        //Ribbon.Variables.editBox_largura_Text = private_largura;
        //        //Ribbon.Variables.editBox_altura_Text = private_altura;
        //    }

        //    public static string largura { get { return private_largura; } set { private_largura = value; } }
        //    public static string altura { get { return private_altura; } set { private_altura = value; } }
        //    public static string largura_checked { get { return private_largura_checked; } set { private_largura_checked = value; } }
        //    public static string separador { get { return private_separador; } set { private_separador = value; } }
        //    public static string painel_de_estilos { get { return private_painel_de_estilos; } set { private_painel_de_estilos = value; } }
        //    public static string assinar_pdf { get { return private_assinar_pdf; } set { private_assinar_pdf = value; } }
        //    public static string abrir_pdf { get { return private_abrir_pdf; } set { private_abrir_pdf = value; } }

        //    // Método auxiliar para obter o valor de um elemento XML
        //    private static string GetElementValue(XmlElement root, string tagName, string defaultValue)
        //    {
        //        XmlNode node = root.SelectSingleNode(tagName);
        //        return node != null ? node.InnerText : defaultValue;
        //    }

        //}


        //public class preferences
        private static Dictionary<string, string> dict_preferences_campo_e_valor = new Dictionary<string, string>()
        {
            { "largura", "10" },
            { "altura", "10" },
            { "largura_checked", "true" },
            { "separador", "Nenhum" },
            { "painel_de_estilos", "false" },
            { "assinar_pdf", "true" },
            { "abrir_pdf", "true" }
        };

        // Função para obter o valor de uma preferência
        public static string GetPreference(string key)
        {
            // Verifica se a chave existe e retorna o valor, ou retorna null
            return dict_preferences_campo_e_valor.ContainsKey(key) ? dict_preferences_campo_e_valor[key] : null;
        }

        // Função para alterar o valor de uma preferência
        public static void ChangePreference(string key, string value)
        {
            // Verifica se a chave existe no dicionário
            if (dict_preferences_campo_e_valor.ContainsKey(key))
            {
                // Altera o valor da chave existente
                dict_preferences_campo_e_valor[key] = value;
            }
            //else
            //{
            //    // Se a chave não existir, cria uma nova chave com o valor
            //    dict_preferences_campo_e_valor.Add(key, value);
            //}
        }

        public void le_preferencias(string caminho_preferences)
        {
            if (File.Exists(Ribbon.Variables.caminho_preferences))
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Ribbon.Variables.caminho_preferences);
                XmlElement root = xmlDoc.DocumentElement;

                List<String> list_key = dict_preferences_campo_e_valor.Keys.ToList();
                foreach (var key in list_key)
                {
                    dict_preferences_campo_e_valor[key] = GetElementValue(root, key, dict_preferences_campo_e_valor[key]);
                }
            }
        }
        // Método auxiliar para obter o valor de um elemento XML
        private static string GetElementValue(XmlElement root, string tagName, string defaultValue)
        {
            XmlNode node = root.SelectSingleNode(tagName);
            return node != null ? node.InnerText : defaultValue;
        }

        //private static void preferences() // Bloco estático para definir o valor inicial das variáveis (Leitura das preferencias)
        //    {
        //         string private_largura, private_altura, private_largura_checked, private_separador, private_painel_de_estilos, private_assinar_pdf, private_abrir_pdf;
        //        //MessageBox.Show("preferences");
        //        if (File.Exists(Ribbon.Variables.caminho_preferences))
        //        {
        //            XmlDocument xmlDoc = new XmlDocument();
        //            xmlDoc.Load(Ribbon.Variables.caminho_preferences);
        //            XmlElement root = xmlDoc.DocumentElement;

        //        // Lendo as preferências
        //        foreach (var key in dict_preferences_campo_e_valor.Keys)
        //        {
        //            preferences[key] = GetElementValue(root, key, defaultPreferences[key]);
        //        }
        //        //private_largura = GetElementValue(root, "largura", "10");
        //        //    private_altura = GetElementValue(root, "altura", "10");
        //        //    private_largura_checked = GetElementValue(root, "largura_checked", "true");
        //        //    private_separador = GetElementValue(root, "separador", "Nenhum");
        //        //    private_painel_de_estilos = GetElementValue(root, "painel_de_estilos", "false");
        //        //    private_assinar_pdf = GetElementValue(root, "assinar_pdf", "true");
        //        //    private_abrir_pdf = GetElementValue(root, "abrir_pdf", "true");
        //        }
        //        else
        //        {
        //            // Preferências iniciais
        //            private_largura = "10";
        //            private_altura = "10";
        //            private_largura_checked = "true";
        //            private_separador = "Nenhum";
        //            private_painel_de_estilos = "false";
        //            private_assinar_pdf = "true";
        //            private_abrir_pdf = "true";
        //        }

        //        // Ajuste inicial dos valores de editBox_largura_Text e editBox_altura_Text
        //        //Ribbon.Variables.editBox_largura_Text = private_largura;
        //        //Ribbon.Variables.editBox_altura_Text = private_altura;
        //    }

        //    // Método auxiliar para obter o valor de um elemento XML
        //    private static string GetElementValue(XmlElement root, string tagName, string defaultValue)
        //    {
        //        XmlNode node = root.SelectSingleNode(tagName);
        //        return node != null ? node.InnerText : defaultValue;
        //    }

        //}



    }

}
