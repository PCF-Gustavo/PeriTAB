using Microsoft.Office.Interop.Word;
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
            Button_renomeia_documento_valorinicial();
            Button_gera_pdf_valorinicial();

            // Atribuição das preferências
            DropDown_unidade_valorinicial();
            DropDown_precisao_valorinicial();
            ToggleButton_painel_de_estilos_valorinicial();
            CheckBox_largura_valorinicial();
            CheckBox_altura_valorinicial();
            EditBox_largura_valorinicial();
            EditBox_altura_valorinicial();
            DropDown_separador_valorinicial();
            CheckBox_assinar_valorinicial();
            CheckBox_abrir_valorinicial();
        }

        private void ToggleButton_painel_de_estilos_valorinicial()
        {
            Globals.Ribbons.Ribbon.ToggleButton_painel_de_estilos.Checked = bool.Parse(Retorna_preferencia("painel_de_estilos"));
        }

        private void CheckBox_largura_valorinicial()
        {
            Globals.Ribbons.Ribbon.CheckBox_largura.Checked = bool.Parse(Retorna_preferencia("largura_checked"));
        }
        private void EditBox_largura_valorinicial()
        {
            Globals.Ribbons.Ribbon.EditBox_largura.Enabled = bool.Parse(Retorna_preferencia("largura_checked"));
            if (Globals.Ribbons.Ribbon.CheckBox_largura.Checked) { Globals.Ribbons.Ribbon.EditBox_largura.Text = Retorna_preferencia("largura"); }

        }

        private void CheckBox_altura_valorinicial()
        {
            Globals.Ribbons.Ribbon.CheckBox_altura.Checked = !bool.Parse(Retorna_preferencia("largura_checked"));
        }

        private void EditBox_altura_valorinicial()
        {
            Globals.Ribbons.Ribbon.EditBox_altura.Enabled = !bool.Parse(Retorna_preferencia("largura_checked"));
            if (Globals.Ribbons.Ribbon.CheckBox_altura.Checked) { Globals.Ribbons.Ribbon.EditBox_altura.Text = Retorna_preferencia("altura"); }
        }

        private void DropDown_separador_valorinicial()
        {
            Globals.Ribbons.Ribbon.DropDown_separador.SelectedItem = GetDropDownItemFromLabel(Globals.Ribbons.Ribbon.DropDown_separador, Retorna_preferencia("separador"));
        }
        private void DropDown_precisao_valorinicial()
        {
            Globals.Ribbons.Ribbon.DropDown_precisao.SelectedItem = GetDropDownItemFromLabel(Globals.Ribbons.Ribbon.DropDown_precisao, Retorna_preferencia("precisao"));
        }
        private void DropDown_unidade_valorinicial()
        {
            Globals.Ribbons.Ribbon.DropDown_unidade.SelectedItem = GetDropDownItemFromLabel(Globals.Ribbons.Ribbon.DropDown_unidade, Retorna_preferencia("unidade"));
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
        private void CheckBox_assinar_valorinicial()
        {
            Globals.Ribbons.Ribbon.CheckBox_assinar.Checked = bool.Parse(Retorna_preferencia("assinar_pdf"));
        }
        private void CheckBox_abrir_valorinicial()
        {
            Globals.Ribbons.Ribbon.CheckBox_abrir.Checked = bool.Parse(Retorna_preferencia("abrir_pdf"));
        }
        public void Button_renomeia_documento_valorinicial()
        {
            Globals.Ribbons.Ribbon.Button_renomeia_documento.Enabled = true;
            Globals.Ribbons.Ribbon.Button_renomeia_documento.ScreenTip = "";
            Globals.Ribbons.Ribbon.Button_renomeia_documento.SuperTip = "Renomeia o documento atual.";
        }
        public void Button_gera_pdf_valorinicial()
        {
            Globals.Ribbons.Ribbon.Button_gera_pdf.Enabled = true;
            Globals.Ribbons.Ribbon.Button_gera_pdf.ScreenTip = "";
            Globals.Ribbons.Ribbon.Button_gera_pdf.SuperTip = "Gera o PDF do documento na pasta onde está salvo.";
        }

        //Dicionario de preferencias com valores iniciais padrao
        private static readonly Dictionary<string, string> Dicionario_preferencias_campo_e_valor = new Dictionary<string, string>()
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
        public static string Retorna_preferencia(string key)
        {
            // Verifica se a chave existe e retorna o valor, ou retorna null
            return Dicionario_preferencias_campo_e_valor.ContainsKey(key) ? Dicionario_preferencias_campo_e_valor[key] : null;
        }

        // Função para atualizar o dicionario de preferencias com as informacoes do arquivo preferences.xml
        public static void Muda_preferencia(string key, string value)
        {
            // Verifica se a chave existe no dicionário
            if (Dicionario_preferencias_campo_e_valor.ContainsKey(key))
            {
                // Altera o valor da chave existente
                Dicionario_preferencias_campo_e_valor[key] = value;
            }
        }

        public void Le_preferencias(string caminho_preferences)
        {
            if (File.Exists(Ribbon.Variables.Caminho_preferences))
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Ribbon.Variables.Caminho_preferences);
                XmlElement root = xmlDoc.DocumentElement;

                List<String> list_key = Dicionario_preferencias_campo_e_valor.Keys.ToList();
                foreach (var key in list_key)
                {
                    Dicionario_preferencias_campo_e_valor[key] = GetElementValue(root, key, Dicionario_preferencias_campo_e_valor[key]);
                }
            }
        }
        // Método auxiliar para obter o valor de um elemento XML
        private static string GetElementValue(XmlElement root, string tagName, string defaultValue)
        {
            XmlNode node = root.SelectSingleNode(tagName);
            return node != null ? node.InnerText : defaultValue;
        }

        public void Atualiza_Habilitacao(Microsoft.Office.Tools.Ribbon.RibbonControl oRibbonControl) 
        {
            try
            {
                if (false) { }
                else if (ReferenceEquals(oRibbonControl, Globals.Ribbons.Ribbon.Button_renomeia_documento))
                {
                    Button_renomeia_documento_valorinicial();
                    if (Globals.ThisAddIn.Application.ActiveDocument.Path == "")
                    {
                        Globals.Ribbons.Ribbon.Button_renomeia_documento.Enabled = false;
                        Globals.Ribbons.Ribbon.Button_renomeia_documento.ScreenTip = "Desabilitado";
                        Globals.Ribbons.Ribbon.Button_renomeia_documento.SuperTip = "Este documento ainda não foi salvo.";
                    }
                    return;
                }
                else if (ReferenceEquals(oRibbonControl, Globals.Ribbons.Ribbon.Button_gera_pdf))
                {
                    Button_gera_pdf_valorinicial();
                    if (Globals.ThisAddIn.Application.ActiveDocument.Path == "")
                    {
                        Globals.Ribbons.Ribbon.Button_gera_pdf.Enabled = false;
                        Globals.Ribbons.Ribbon.Button_gera_pdf.ScreenTip = "Desabilitado";
                        Globals.Ribbons.Ribbon.Button_gera_pdf.SuperTip = "Este documento ainda não foi salvo.";
                    }
                    return;
                }
                else if (ReferenceEquals(oRibbonControl, Globals.Ribbons.Ribbon.CheckBox_destaca_campos))
                {
                    if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)1) Globals.Ribbons.Ribbon.CheckBox_destaca_campos.Checked = true;
                    if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)0 | Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)2)
                        Globals.Ribbons.Ribbon.CheckBox_destaca_campos.Checked = false;
                    return;
                }
                else if (ReferenceEquals(oRibbonControl, Globals.Ribbons.Ribbon.CheckBox_mostra_indicadores))
                {
                    if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowBookmarks == true) Globals.Ribbons.Ribbon.CheckBox_mostra_indicadores.Checked = true;
                    if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowBookmarks == false) Globals.Ribbons.Ribbon.CheckBox_mostra_indicadores.Checked = false;
                    return;
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            { }
            }

    }

}
