using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System.Windows;
using System.Diagnostics;
using System.Xml;

namespace PeriTAB
{
    public partial class ThisAddIn
    {
        public MyUserControl iMyUserControl;
        public Dictionary<Microsoft.Office.Interop.Word.Document, MyUserControl> Dicionario_Doc_e_UserControl = new Dictionary<Microsoft.Office.Interop.Word.Document, MyUserControl>();
        //public Microsoft.Office.Tools.CustomTaskPane TaskPane1;

        //Class_New_or_Open_Event iClass_New_or_Open_Event;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            //MessageBox.Show("ThisAddIn_Startup");
            //le_preferencias(Ribbon1.Variables.caminho_preferences);

            // Ajuste dos valores de editBox_largura_Text e
            //Ribbon1.Variables.editBox_largura_Text = Class_Buttons.preferences.largura;
            //Ribbon1.Variables.editBox_altura_Text = Class_Buttons.preferences.altura;

            //Configura o Task Pane
            //iMyUserControl = new MyUserControl();
            //TaskPane1 = Globals.ThisAddIn.CustomTaskPanes.Add(iMyUserControl, "Painel de Estilos (PeriTAB)");
            //TaskPane1.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
            //TaskPane1.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
            //TaskPane1.Height = 90;
            //TaskPane1.VisibleChanged += MyCustomTaskPane_VisibleChanged;



            //Inicia Eventos
            //iClass_New_or_Open_Event = new Class_New_or_Open_Event(); iClass_New_or_Open_Event.Evento_New_or_Open();
            Class_New_or_Open_Event iClass_New_or_Open_Event = new Class_New_or_Open_Event(); iClass_New_or_Open_Event.Evento_New_or_Open();
            //MessageBox.Show(Globals.ThisAddIn.CustomTaskPanes.Count.ToString());
            //MessageBox.Show(Globals.ThisAddIn.Application.Documents.Count.ToString());
            if (Globals.ThisAddIn.Application.Documents.Count == 1) { /*MessageBox.Show("sss");*/ iClass_New_or_Open_Event.Metodo_New_or_Open(Globals.ThisAddIn.Application.ActiveDocument); }
            Class_DocumentBeforeClose_Event iClass_DocumentBeforeClose_Event = new Class_DocumentBeforeClose_Event(); iClass_DocumentBeforeClose_Event.Evento_DocumentBeforeClose();
            //Class_AnyButtonClick_Event iClass_AnyButtonClick_Event = new Class_AnyButtonClick_Event(); iClass_AnyButtonClick_Event.Evento_AnyButtonClick();
            Class_Buttons iClass_Buttons = new Class_Buttons(); iClass_Buttons.DefaultAll();
            Class_DocSave_Event iClass_DocSave_Event = new Class_DocSave_Event(); iClass_DocSave_Event.Evento_DocSave();            
            Class_SelectionChange_Event iClass_SelectionChange_Event = new Class_SelectionChange_Event(); iClass_SelectionChange_Event.Evento_SelectionChange();
            Class_WindowActivate_Event iClass_WindowActivate_Event = new Class_WindowActivate_Event(); iClass_WindowActivate_Event.Evento_WindowActivate();
            Class_WindowDeactivate_Event iClass_WindowDeactivate_Event = new Class_WindowDeactivate_Event(); iClass_WindowDeactivate_Event.Evento_WindowDeactivate();

            //iClass_New_or_Open_Event.Metodo_TaskPane2_Visible(true);

            //Configura o primeiro Task Pane
            //iClass_New_or_Open_Event.Metodo_New_or_Open(null);

            //if (Globals.ThisAddIn.Application.Documents.Count == 1) { /*MessageBox.Show("sss");*/ iClass_New_or_Open_Event.Metodo_New_or_Open(null); }
        }



        //private void MyCustomTaskPane_VisibleChanged(object sender, EventArgs e)
        //{
        //    //if (Globals.ThisAddIn.TaskPane1.Visible == false & Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked == true) { Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked = false; }
        //}

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (File.Exists(Ribbon1.Variables.caminho_template))
            {
                Globals.ThisAddIn.Application.AddIns.Unload(true);
                try { File.Delete(Ribbon1.Variables.caminho_template); } catch (IOException) { }
                try { escreve_preferencias(Ribbon1.Variables.caminho_preferences); } catch (IOException) { }
            }
        }

        //private void escreve_preferencias_antiga()
        //{
        //    if (!Directory.Exists(Ribbon1.Variables.caminho_AppData_Roaming_PeriTAB)) { Directory.CreateDirectory(Ribbon1.Variables.caminho_AppData_Roaming_PeriTAB); } //Cria a pasta AppData/Roaming/PeriTAB caso não exista

        //    string preferences_path = Path.Combine(Ribbon1.Variables.caminho_AppData_Roaming_PeriTAB, "preferences");

        //    string preferences = "";
            
        //    if (Globals.Ribbons.Ribbon1.editBox_largura.Text != "" & Globals.Ribbons.Ribbon1.editBox_largura.Text != null) 
        //    { 
        //        preferences += "<largura>" + Globals.Ribbons.Ribbon1.editBox_largura.Text + "</largura>" + System.Environment.NewLine;
        //        //MessageBox.Show("1L");
        //    } 
        //    else if (Ribbon1.Variables.editBox_largura_Text != "" & Ribbon1.Variables.editBox_largura_Text != null) 
        //    { 
        //        preferences += "<largura>" + Ribbon1.Variables.editBox_largura_Text + "</largura>" + System.Environment.NewLine;
        //        //MessageBox.Show("2L");
        //    } 
        //    else 
        //    { 
        //        preferences += "<largura>" + Class_Buttons.preferences.largura + "</largura>" + System.Environment.NewLine;
        //        //MessageBox.Show("3L");
        //    }

        //    if (Globals.Ribbons.Ribbon1.editBox_altura.Text != "" & Globals.Ribbons.Ribbon1.editBox_altura.Text != null)
        //    { 
        //        preferences += "<altura>" + Globals.Ribbons.Ribbon1.editBox_altura.Text + "</altura>" + System.Environment.NewLine;
        //        //MessageBox.Show("1A");
        //    } 
        //    else if (Ribbon1.Variables.editBox_altura_Text != "" & Ribbon1.Variables.editBox_altura_Text != null) 
        //    {
        //        preferences += "<altura>" + Ribbon1.Variables.editBox_altura_Text + "</altura>" + System.Environment.NewLine;
        //        //MessageBox.Show("2A");
        //    } 
        //    else 
        //    {
        //        preferences += "<altura>" + Class_Buttons.preferences.altura + "</altura>" + System.Environment.NewLine;
        //        //MessageBox.Show("3A");
        //    }
            
        //    preferences += "<largura_checked>" + Globals.Ribbons.Ribbon1.checkBox_largura.Checked.ToString() + "</largura_checked>" + System.Environment.NewLine;
        //    //preferences += "<ordem>" + Globals.Ribbons.Ribbon1.dropDown_ordem.SelectedItem.Label + "</ordem>" + System.Environment.NewLine;
        //    preferences += "<separador>" + Globals.Ribbons.Ribbon1.dropDown_separador.SelectedItem.Label + "</separador>" + System.Environment.NewLine;
        //    preferences += "<painel_de_estilos>" + Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked.ToString() + "</painel_de_estilos>" + System.Environment.NewLine;
        //    preferences += "<assinar_pdf>" + Globals.Ribbons.Ribbon1.checkBox_assinar.Checked.ToString() + "</assinar_pdf>" + System.Environment.NewLine;
        //    preferences += "<abrir_pdf>" + Globals.Ribbons.Ribbon1.checkBox_abrir.Checked.ToString() + "</abrir_pdf>" + System.Environment.NewLine;
        //    File.WriteAllText(preferences_path, preferences);
        //    //MessageBox.Show(preferences);
        //}

        //private void le_preferencias_antiga()
        //{
        //    string preferences_path = Path.Combine(Ribbon1.Variables.caminho_AppData_Roaming_PeriTAB, "preferences");

        //    if (File.Exists(preferences_path))
        //    {
        //        string preferences_text = File.ReadAllText(preferences_path);

        //        if (procura(preferences_text, "largura") != null)
        //        {
        //            Class_Buttons.preferences.largura = procura(preferences_text, "largura");
        //        }
        //        else
        //        {
        //            Class_Buttons.preferences.largura = "10";
        //        }

        //        if (procura(preferences_text, "altura") != null)
        //        {
        //            Class_Buttons.preferences.altura = procura(preferences_text, "altura");
        //        }
        //        else
        //        {
        //            Class_Buttons.preferences.altura = "10";
        //        }

        //        if (procura(preferences_text, "largura_checked") != null)
        //        {
        //            Class_Buttons.preferences.largura_checked = procura(preferences_text, "largura_checked");
        //        }
        //        else
        //        {
        //            Class_Buttons.preferences.largura_checked = "true";
        //        }

        //        if (procura(preferences_text, "separador") != null)
        //        {
        //            Class_Buttons.preferences.separador = procura(preferences_text, "separador");
        //        }
        //        else
        //        {
        //            Class_Buttons.preferences.separador = "Nenhum";
        //        }

        //        if (procura(preferences_text, "painel_de_estilos") != null)
        //        {
        //            Class_Buttons.preferences.painel_de_estilos = procura(preferences_text, "painel_de_estilos");
        //        }
        //        else
        //        {
        //            Class_Buttons.preferences.painel_de_estilos = "false";
        //        }

        //        if (procura(preferences_text, "assinar_pdf") != null)
        //        {
        //            Class_Buttons.preferences.assinar_pdf = procura(preferences_text, "assinar_pdf");
        //        }
        //        else
        //        {
        //            Class_Buttons.preferences.assinar_pdf = "true";
        //        }

        //        if (procura(preferences_text, "abrir_pdf") != null)
        //        {
        //            Class_Buttons.preferences.abrir_pdf = procura(preferences_text, "abrir_pdf");
        //        }
        //        else
        //        {
        //            Class_Buttons.preferences.abrir_pdf = "true";
        //        }

        //        //Class_Buttons.preferences.largura = procura(preferences_text, "largura");
        //        //Class_Buttons.preferences.altura = procura(preferences_text, "altura");
        //        //Class_Buttons.preferences.largura_checked = procura(preferences_text, "largura_checked");
        //        //Class_Buttons.preferences.ordem = procura(preferences_text, "ordem");
        //        //Class_Buttons.preferences.separador = procura(preferences_text, "separador");
        //        //Class_Buttons.preferences.painel_de_estilos = procura(preferences_text, "painel_de_estilos");
        //        //Class_Buttons.preferences.assinar_pdf = procura(preferences_text, "assinar_pdf");
        //        //Class_Buttons.preferences.abrir_pdf = procura(preferences_text, "abrir_pdf");
        //    }
        //    else
        //    { // Preferências iniciais
        //        Class_Buttons.preferences.largura = "10";
        //        Class_Buttons.preferences.altura = "10";
        //        Class_Buttons.preferences.largura_checked = "true";
        //        //Class_Buttons.preferences.ordem = "Alfabética";
        //        Class_Buttons.preferences.separador = "Nenhum";
        //        Class_Buttons.preferences.painel_de_estilos = "false";
        //        Class_Buttons.preferences.assinar_pdf = "true";
        //        Class_Buttons.preferences.abrir_pdf = "true";
        //    }
        //}

        //private string procura(string texto, string valor) 
        //{
        //    string str1 = "<" + valor + ">";
        //    string str2 = "</" + valor + ">";


        //    if (texto.IndexOf(str1) > -1 & texto.IndexOf(str2) > -1)
        //    {
        //        return texto.Substring((texto.IndexOf(str1) + (str1).Length), texto.IndexOf(str2) - (texto.IndexOf(str1) + (str1).Length));
        //    }
        //    else
        //    {
        //        return null;
        //    }
        //}

   
        private void escreve_preferencias(string caminho_preferences)
        {
            if (!Directory.Exists(Ribbon1.Variables.caminho_AppData_Roaming_PeriTAB))
            {
                Directory.CreateDirectory(Ribbon1.Variables.caminho_AppData_Roaming_PeriTAB);
            }

            //string preferences_path = Path.Combine(Ribbon1.Variables.caminho_AppData_Roaming_PeriTAB, "preferences.xml");

            // Cria um dicionário de preferências
            var preferencias = new Dictionary<string, string>
    {
        { "largura", string.IsNullOrEmpty(Globals.Ribbons.Ribbon1.editBox_largura.Text) ? Ribbon1.Variables.editBox_largura_Text : Globals.Ribbons.Ribbon1.editBox_largura.Text }, // Verifica e define valores para largura se for vazio ou null (Ribbon1.Variables.editBox_largura_Text)
        { "altura", string.IsNullOrEmpty(Globals.Ribbons.Ribbon1.editBox_altura.Text) ? Ribbon1.Variables.editBox_altura_Text : Globals.Ribbons.Ribbon1.editBox_altura.Text },// Verifica e define valores para altura se for vazio ou null (Ribbon1.Variables.editBox_altura_Text)
        { "largura_checked", Globals.Ribbons.Ribbon1.checkBox_largura.Checked.ToString() },
        { "separador", Globals.Ribbons.Ribbon1.dropDown_separador.SelectedItem.Label },
        { "painel_de_estilos", Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked.ToString() },
        { "assinar_pdf", Globals.Ribbons.Ribbon1.checkBox_assinar.Checked.ToString() },
        { "abrir_pdf", Globals.Ribbons.Ribbon1.checkBox_abrir.Checked.ToString() }
    };
            //if (string.IsNullOrEmpty(preferencias["altura"]) || string.IsNullOrEmpty(preferencias["largura"])) 
            //{
            //    MessageBox.Show("erro na preferencia");
            //}
            //MessageBox.Show(preferencias["altura"]);
            // Criar um XmlDocument para gerar o XML
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement rootElement = xmlDoc.CreateElement("preferencias");

            // Adicionar as preferências no XML
            foreach (var pref in preferencias)
            {
                XmlElement prefElement = xmlDoc.CreateElement(pref.Key);
                prefElement.InnerText = pref.Value;
                rootElement.AppendChild(prefElement);
            }

            // Adicionar o rootElement ao documento
            xmlDoc.AppendChild(rootElement);

            // Salvar o XML no arquivo
            xmlDoc.Save(caminho_preferences);
        }

        //private void le_preferencias(string caminho_preferences)
        //{
        //    //string preferences_path = Path.Combine(Ribbon1.Variables.caminho_AppData_Roaming_PeriTAB, "preferences.xml");

        //    if (File.Exists(caminho_preferences))
        //    {
        //        XmlDocument xmlDoc = new XmlDocument();
        //        xmlDoc.Load(caminho_preferences);
        //        XmlElement root = xmlDoc.DocumentElement;

        //        // Lendo as preferências
        //        Class_Buttons.preferences.largura = GetElementValue(root, "largura", "10");
        //        Class_Buttons.preferences.altura = GetElementValue(root, "altura", "10");
        //        Class_Buttons.preferences.largura_checked = GetElementValue(root, "largura_checked", "true");
        //        Class_Buttons.preferences.separador = GetElementValue(root, "separador", "Nenhum");
        //        Class_Buttons.preferences.painel_de_estilos = GetElementValue(root, "painel_de_estilos", "false");
        //        Class_Buttons.preferences.assinar_pdf = GetElementValue(root, "assinar_pdf", "true");
        //        Class_Buttons.preferences.abrir_pdf = GetElementValue(root, "abrir_pdf", "true");
        //    }
        //    else
        //    {
        //        // Preferências iniciais
        //        Class_Buttons.preferences.largura = "10";
        //        Class_Buttons.preferences.altura = "10";
        //        Class_Buttons.preferences.largura_checked = "true";
        //        Class_Buttons.preferences.separador = "Nenhum";
        //        Class_Buttons.preferences.painel_de_estilos = "false";
        //        Class_Buttons.preferences.assinar_pdf = "true";
        //        Class_Buttons.preferences.abrir_pdf = "true";
        //    }
        //}

        //// Método auxiliar para obter o valor de um elemento XML
        //private string GetElementValue(XmlElement root, string tagName, string defaultValue)
        //{
        //    XmlNode node = root.SelectSingleNode(tagName);
        //    return node != null ? node.InnerText : defaultValue;
        //}

        #region Código gerado por VSTO

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
