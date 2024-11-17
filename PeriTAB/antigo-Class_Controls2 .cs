using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.IO;

namespace PeriTAB
{
    internal class Class_Controls2 // Classe para maninupulação dos controles
    {
        // Define valor inicial para todos os controles
        public void Configura_Valores_iniciais()
        {
            SetControl(Control: Globals.Ribbons.Ribbon.button_confere_num_legenda, Enabled: true, ScreenTip: "Macro confere_numeracao_legendas", SuperTip: "Corrige erros de numeração em Figuras, Tabelas etc.");
            SetControl(Control: Globals.Ribbons.Ribbon.button_alinha_legenda, Enabled: true, ScreenTip: "Macro alinha_legenda", SuperTip: "Alinha legenda de Figuras, Tabelas etc.");
            SetControl(Control: Globals.Ribbons.Ribbon.button_atualiza_campos, Enabled: true, ScreenTip: "Macro atualiza_todos_campos", SuperTip: "Atualiza todos os campos (MS Word Fields) do documento.");
            SetControl(Control: Globals.Ribbons.Ribbon.button_moeda, Enabled: true, ScreenTip: "Macro moeda_por_extenso", SuperTip: "Escreve por extenso o valor em Reais. Posicione o cursor ao final do número.");
            SetControl(Control: Globals.Ribbons.Ribbon.button_inteiro, Enabled: true, ScreenTip: "Macro moeda_por_extenso", SuperTip: "Escreve por extenso o número inteiro. Posicione o cursor ao final do número.");
            SetControl(Control: Globals.Ribbons.Ribbon.button_limpa_estilos, Enabled: true, ScreenTip: "Macro limpa_estilos", SuperTip: "Remove Estilos de parágrafos não utilizados.");
            SetControl(Control: Globals.Ribbons.Ribbon.toggleButton_painel_de_estilos, Checked: bool.Parse(GetPreference("painel_de_estilos")));
            SetControl(Control: Globals.Ribbons.Ribbon.button_cola_imagem, Enabled: true, SuperTip: "Cola imagens do Clipboard em ordem alfabética.");
            SetControl(Control: Globals.Ribbons.Ribbon.checkBox_largura, Checked: bool.Parse(GetPreference("largura_checked")));
            SetControl(Control: Globals.Ribbons.Ribbon.editBox_largura, Checked: bool.Parse(GetPreference("largura_checked")));
            SetControl(Control: Globals.Ribbons.Ribbon.checkBox_altura, Checked: !bool.Parse(GetPreference("largura_checked")));
            SetControl(Control: Globals.Ribbons.Ribbon.editBox_altura, Checked: !bool.Parse(GetPreference("largura_checked")));
            if (bool.Parse(GetPreference("largura_checked"))){ SetControl(Control: Globals.Ribbons.Ribbon.editBox_largura, Text: GetPreference("largura")); }
            if (!bool.Parse(GetPreference("largura_checked"))){ SetControl(Control: Globals.Ribbons.Ribbon.editBox_altura, Text: GetPreference("altura")); }
            SetControl(Control: Globals.Ribbons.Ribbon.dropDown_separador, SelectedItem: GetDropDownItemFromLabel(Globals.Ribbons.Ribbon.dropDown_separador, GetPreference("separador")));
            SetControl(Control: Globals.Ribbons.Ribbon.checkBox_assinar, Checked: bool.Parse(GetPreference("assinar_pdf")));
            SetControl(Control: Globals.Ribbons.Ribbon.checkBox_abrir, Checked: bool.Parse(GetPreference("abrir_pdf")));
            SetControl(Control: Globals.Ribbons.Ribbon.button_renomeia_documento, Enabled: true, SuperTip: "Renomeia o documento atual.");
            SetControl(Control: Globals.Ribbons.Ribbon.button_gera_pdf, Enabled: true, SuperTip: "Gera o PDF do documento na pasta onde está salvo.");
            SetControl(Control: Globals.Ribbons.Ribbon.button_abre_SISCRIM, Enabled: true, SuperTip: "Abre SISCRIM na página do Laudo ou da Requisição.");
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

        public void SetControl(RibbonControl Control, bool Enabled = true, bool Visible = true, string ScreenTip = "", string SuperTip = "", bool Checked = false, string Text = "", RibbonDropDownItem SelectedItem = null, System.Drawing.Bitmap Image = null)
        {
            // Definir propriedades comuns
            Control.Enabled = Enabled;
            Control.Visible = Visible;
            // Ações específicas para cada tipo de controle
            switch (Control)
            {
                case RibbonButton Button:
                    Button.ScreenTip = ScreenTip;
                    Button.SuperTip = SuperTip;
                    Button.Image = Image;
                    break;

                case RibbonCheckBox CheckBox:
                    CheckBox.ScreenTip = ScreenTip;
                    CheckBox.SuperTip = SuperTip;
                    CheckBox.Checked = Checked;
                    break;

                case RibbonToggleButton ToggleButton:
                    ToggleButton.ScreenTip = ScreenTip;
                    ToggleButton.SuperTip = SuperTip;
                    ToggleButton.Checked = Checked;
                    break;

                case RibbonEditBox EditBox:
                    EditBox.ScreenTip = ScreenTip;
                    EditBox.SuperTip = SuperTip;
                    EditBox.Text = Text;
                    break;

                case RibbonDropDown DropDown:
                    DropDown.ScreenTip = ScreenTip;
                    DropDown.SuperTip = SuperTip;
                    DropDown.SelectedItem = SelectedItem;
                    break;
            }
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

        // Função para alterar o valor de uma preferência
        public static void ChangePreference(string key, string value)
        {
            // Verifica se a chave existe no dicionário
            if (dict_preferences_campo_e_valor.ContainsKey(key))
            {
                // Altera o valor da chave existente
                dict_preferences_campo_e_valor[key] = value;
            }
        }

        // Função para atualizar o dicionario de preferencias com aquelas do arquivo preferences.xml
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
