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
            button_renomeia_documento_valorinicial();
            button_gera_pdf_valorinicial();
            button_abre_SISCRIM_valorinicial();

            // Atribuição das preferências
            dropDown_unidade_valorinicial();
            dropDown_precisao_valorinicial();
            toggleButton_painel_de_estilos_valorinicial();
            checkBox_largura_valorinicial();
            checkBox_altura_valorinicial();
            editBox_largura_valorinicial();
            editBox_altura_valorinicial();
            dropDown_separador_valorinicial();
            checkBox_assinar_valorinicial();
            checkBox_abrir_valorinicial();
        }

        public void toggleButton_painel_de_estilos_valorinicial()
        {
            Globals.Ribbons.Ribbon.toggleButton_painel_de_estilos.Checked = bool.Parse(GetPreference("painel_de_estilos"));
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
        public void dropDown_precisao_valorinicial()
        {
            Globals.Ribbons.Ribbon.dropDown_precisao.SelectedItem = GetDropDownItemFromLabel(Globals.Ribbons.Ribbon.dropDown_precisao, GetPreference("precisao"));
        }        
        public void dropDown_unidade_valorinicial()
        {
            Globals.Ribbons.Ribbon.dropDown_unidade.SelectedItem = GetDropDownItemFromLabel(Globals.Ribbons.Ribbon.dropDown_unidade, GetPreference("unidade"));
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
             { "unidade", "quilograma (kg)" }
            ,{ "precisao", "0,00" }
            ,{ "painel_de_estilos", "false" }
            ,{ "largura_checked", "true" }
            ,{ "largura", "10" }
            ,{ "altura", "10" }
            ,{ "separador", "Nenhum" }
            ,{ "assinar_pdf", "true" }
            ,{ "abrir_pdf", "true" }
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
