using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace PeriTAB
{
    public partial class ThisAddIn
    {
        // Cria instância das classes
        Class_New_or_Open_Event iClass_New_or_Open_Event = new Class_New_or_Open_Event();
        Class_DocumentClose_Event iClass_DocumentClose_Event = new Class_DocumentClose_Event();
        Class_DocSave_Event iClass_DocSave_Event = new Class_DocSave_Event();
        Class_SelectionChange_Event iClass_SelectionChange_Event = new Class_SelectionChange_Event();
        Class_WindowActivate_Event iClass_WindowActivate_Event = new Class_WindowActivate_Event();
        Class_WindowDeactivate_Event iClass_WindowDeactivate_Event = new Class_WindowDeactivate_Event();

        public MyUserControl iMyUserControl;
        Class_RibbonControls iClass_RibbonControls = new Class_RibbonControls();

        public Dictionary<Microsoft.Office.Interop.Word.Document, MyUserControl> Dicionario_Doc_e_UserControl = new Dictionary<Microsoft.Office.Interop.Word.Document, MyUserControl>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Lê e configura preferências
            iClass_RibbonControls.le_preferencias(Ribbon.Variables.caminho_preferences);
            iClass_RibbonControls.Configura_Valores_iniciais();

            //Inicia Eventos
            iClass_New_or_Open_Event.Evento_New_or_Open();
            if (Globals.ThisAddIn.Application.Documents.Count == 1) iClass_New_or_Open_Event.Metodo_New_or_Open(Globals.ThisAddIn.Application.ActiveDocument);
            iClass_DocumentClose_Event.Evento_DocumentClose();
            iClass_DocSave_Event.Evento_DocSave();
            iClass_SelectionChange_Event.Evento_SelectionChange();
            iClass_WindowActivate_Event.Evento_WindowActivate();
            iClass_WindowDeactivate_Event.Evento_WindowDeactivate();


            
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (File.Exists(Ribbon.Variables.caminho_template))
            {
                Globals.ThisAddIn.Application.AddIns.Unload(true);
                try { File.Delete(Ribbon.Variables.caminho_template); } catch (IOException) { }
                try { escreve_preferencias(Ribbon.Variables.caminho_preferences); } catch (IOException) { }
            }
        }
        private void escreve_preferencias(string caminho_preferences)
        {
            if (!Directory.Exists(Ribbon.Variables.caminho_AppData_Roaming_PeriTAB))
            {
                Directory.CreateDirectory(Ribbon.Variables.caminho_AppData_Roaming_PeriTAB);
            }

            // Cria um dicionário de preferências
            Dictionary<string, string> preferencias = new Dictionary<string, string>
            {
                 { "largura", string.IsNullOrEmpty(Globals.Ribbons.Ribbon.editBox_largura.Text) ? Class_RibbonControls.GetPreference("largura") : Globals.Ribbons.Ribbon.editBox_largura.Text } // Verifica e define valores para largura se for vazio ou null (Ribbon.Variables.editBox_largura_Text)
                ,{ "altura", string.IsNullOrEmpty(Globals.Ribbons.Ribbon.editBox_altura.Text) ? Class_RibbonControls.GetPreference("altura") : Globals.Ribbons.Ribbon.editBox_altura.Text }// Verifica e define valores para altura se for vazio ou null (Ribbon.Variables.editBox_altura_Text)
                ,{ "largura_checked", Globals.Ribbons.Ribbon.checkBox_largura.Checked.ToString() }
                ,{ "separador", Globals.Ribbons.Ribbon.dropDown_separador.SelectedItem.Label }
                ,{ "painel_de_estilos", Globals.Ribbons.Ribbon.toggleButton_painel_de_estilos.Checked.ToString() }
                ,{ "assinar_pdf", Globals.Ribbons.Ribbon.checkBox_assinar.Checked.ToString() }
                ,{ "abrir_pdf", Globals.Ribbons.Ribbon.checkBox_abrir.Checked.ToString() }
            };

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
