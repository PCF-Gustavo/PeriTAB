using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

namespace PeriTAB
{
    internal class Class_RibbonControls
    {
        public void Configura_Valores_iniciais()
        {
            button_confere_num_legenda_valorinicial();
            button_alinha_legenda_valorinicial();
            button_atualiza_campos_valorinicial();
            button_moeda_valorinicial();
            button_inteiro_valorinicial();
            button_limpa_estilos_valorinicial();
            toggleButton_painel_de_estilos_valorinicial();
            button_cola_imagem_valorinicial();
            checkBox_altura_valorinicial();
            checkBox_largura_valorinicial();
            editBox_largura_valorinicial();
            editBox_altura_valorinicial();
            dropDown_separador_valorinicial();
            checkBox_assinar_valorinicial();
            checkBox_abrir_valorinicial();
            button_renomeia_documento_valorinicial();
            button_gera_pdf_valorinicial();
            button_abre_SISCRIM_valorinicial();
        }

        public void button_confere_num_legenda_valorinicial()
        {
            Globals.Ribbons.Ribbon.button_confere_num_legenda.Enabled = true;
            Globals.Ribbons.Ribbon.button_confere_num_legenda.ScreenTip = "Macro confere_numeracao_legendas";
            Globals.Ribbons.Ribbon.button_confere_num_legenda.SuperTip = "Corrige erros de numeração em Figuras, Tabelas etc.";
        }
        public void button_alinha_legenda_valorinicial()
        {
            Globals.Ribbons.Ribbon.button_alinha_legenda.Enabled = true;
            Globals.Ribbons.Ribbon.button_alinha_legenda.ScreenTip = "Macro alinha_legenda";
            Globals.Ribbons.Ribbon.button_alinha_legenda.SuperTip = "Alinha legenda de Figuras, Tabelas etc.";
        }

        public void button_atualiza_campos_valorinicial()
        {
            Globals.Ribbons.Ribbon.button_atualiza_campos.Enabled = true;
            Globals.Ribbons.Ribbon.button_atualiza_campos.ScreenTip = "Macro atualiza_todos_campos";
            Globals.Ribbons.Ribbon.button_atualiza_campos.SuperTip = "Atualiza todos os campos (MS Word Fields) do documento.";
        }
        public void button_moeda_valorinicial()
        {
            Globals.Ribbons.Ribbon.button_moeda.Enabled = true;
            Globals.Ribbons.Ribbon.button_moeda.ScreenTip = "Macro moeda_por_extenso";
            Globals.Ribbons.Ribbon.button_moeda.SuperTip = "Escreve por extenso o valor em Reais. Posicione o cursor ao final do número.";
        }
        public void button_inteiro_valorinicial()
        {
            Globals.Ribbons.Ribbon.button_inteiro.Enabled = true;
            Globals.Ribbons.Ribbon.button_inteiro.ScreenTip = "Macro inteiro_por_extenso";
            Globals.Ribbons.Ribbon.button_inteiro.SuperTip = "Escreve por extenso o número inteiro. Posicione o cursor ao final do número.";
        }

        public void button_limpa_estilos_valorinicial()
        {
            Globals.Ribbons.Ribbon.button_limpa_estilos.Enabled = true;
            Globals.Ribbons.Ribbon.button_limpa_estilos.ScreenTip = "Macro limpa_estilos";
            Globals.Ribbons.Ribbon.button_limpa_estilos.SuperTip = "Remove Estilos de parágrafos não utilizados.";
        }

        public void toggleButton_painel_de_estilos_valorinicial()
        {
            Globals.Ribbons.Ribbon.toggleButton_painel_de_estilos.Checked = bool.Parse(GetPreference("painel_de_estilos"));
        }

        public void button_cola_imagem_valorinicial()
        {
            Globals.Ribbons.Ribbon.button_cola_imagem.Enabled = true;
            Globals.Ribbons.Ribbon.button_cola_imagem.ScreenTip = "";
            Globals.Ribbons.Ribbon.button_cola_imagem.SuperTip = "Cola imagens do Clipboard em ordem alfabética.";
        }

        public void checkBox_largura_valorinicial()
        {
            Globals.Ribbons.Ribbon.checkBox_largura.Checked = bool.Parse(GetPreference("largura_checked"));
        }
        public void editBox_largura_valorinicial()
        {
            Globals.Ribbons.Ribbon.editBox_largura.Enabled = bool.Parse(GetPreference("largura_checked"));
            if (Globals.Ribbons.Ribbon.checkBox_largura.Checked) { Globals.Ribbons.Ribbon.editBox_largura.Text = GetPreference("largura"); }

        }

        public void checkBox_altura_valorinicial()
        {
            Globals.Ribbons.Ribbon.checkBox_altura.Checked = !bool.Parse(GetPreference("largura_checked"));
        }

        public void editBox_altura_valorinicial()
        {
            Globals.Ribbons.Ribbon.editBox_altura.Enabled = !bool.Parse(GetPreference("largura_checked"));
            if (Globals.Ribbons.Ribbon.checkBox_altura.Checked) { Globals.Ribbons.Ribbon.editBox_altura.Text = GetPreference("altura"); }
        }

        public void dropDown_separador_valorinicial()
        {
            Globals.Ribbons.Ribbon.dropDown_separador.SelectedItem = GetDropDownItemFromLabel(Globals.Ribbons.Ribbon.dropDown_separador, GetPreference("separador"));
        }

        private RibbonDropDownItem GetDropDownItemFromLabel(RibbonDropDown Control, string Label)
        {
            foreach (RibbonDropDownItem item in Control.Items)
            {
                if (item.Label == Label)
                {
                    return item;
                }
            }
            return null;  // Retorna null se não encontrar o item
        }
        public void checkBox_assinar_valorinicial()
        {
            Globals.Ribbons.Ribbon.checkBox_assinar.Checked = bool.Parse(GetPreference("assinar_pdf"));
        }
        public void checkBox_abrir_valorinicial()
        {
            Globals.Ribbons.Ribbon.checkBox_abrir.Checked = bool.Parse(GetPreference("abrir_pdf"));
        }
        public void button_renomeia_documento_valorinicial()
        {
            Globals.Ribbons.Ribbon.button_renomeia_documento.Enabled = true;
            Globals.Ribbons.Ribbon.button_renomeia_documento.ScreenTip = "";
            Globals.Ribbons.Ribbon.button_renomeia_documento.SuperTip = "Renomeia o documento atual.";
        }
        public void button_gera_pdf_valorinicial()
        {
            Globals.Ribbons.Ribbon.button_gera_pdf.Image = Properties.Resources.icone_pdf_chave;
            Globals.Ribbons.Ribbon.button_gera_pdf.Enabled = true;
            Globals.Ribbons.Ribbon.button_gera_pdf.ScreenTip = "";
            Globals.Ribbons.Ribbon.button_gera_pdf.SuperTip = "Gera o PDF do documento na pasta onde está salvo.";
        }
        public void button_abre_SISCRIM_valorinicial()
        {
            Globals.Ribbons.Ribbon.button_abre_SISCRIM.Enabled = true;
            Globals.Ribbons.Ribbon.button_abre_SISCRIM.ScreenTip = "";
            Globals.Ribbons.Ribbon.button_abre_SISCRIM.SuperTip = "Abre SISCRIM na página do Laudo ou da Requisição.";
        }

        //Dicionario de preferencias com valores iniciais padrao
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

        // Função para atualizar o dicionario de preferencias com as informacoes do arquivo preferences.xml
        public static void ChangePreference(string key, string value)
        {
            // Verifica se a chave existe no dicionário
            if (dict_preferences_campo_e_valor.ContainsKey(key))
            {
                // Altera o valor da chave existente
                dict_preferences_campo_e_valor[key] = value;
            }
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

    }

}
