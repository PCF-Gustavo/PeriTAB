using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace PeriTAB
{
    public partial class ThisAddIn
    {
        // Cria instância das classes
        //Class_New_or_Open_Event iClass_New_or_Open_Event = new Class_New_or_Open_Event();
        //Class_DocumentClose_Event iClass_DocumentClose_Event = new Class_DocumentClose_Event();
        private readonly Class_DocSave_Event2 iClass_DocSave_Event = new Class_DocSave_Event2();
        private readonly Class_SelectionChange_Event iClass_SelectionChange_Event = new Class_SelectionChange_Event();
        private readonly Class_WindowActivate_Event iClass_WindowActivate_Event = new Class_WindowActivate_Event();
        private readonly Class_WindowDeactivate_Event iClass_WindowDeactivate_Event = new Class_WindowDeactivate_Event();

        //public MyUserControl iMyUserControl;
        private readonly Class_RibbonControls iClass_RibbonControls = new Class_RibbonControls();

        //public Dictionary<Microsoft.Office.Interop.Word.Document, MyUserControl> Dicionario_Doc_e_UserControl = new Dictionary<Microsoft.Office.Interop.Word.Document, MyUserControl>();

        //public Dictionary<Microsoft.Office.Interop.Word.Window, Microsoft.Office.Interop.Word.Document> Dicionario_Window_e_Doc = new Dictionary<Microsoft.Office.Interop.Word.Window, Microsoft.Office.Interop.Word.Document>();
        public Dictionary<Microsoft.Office.Interop.Word.Window, MyUserControl> Dicionario_Window_e_UserControl = new Dictionary<Microsoft.Office.Interop.Word.Window, MyUserControl>();
        //public Dictionary<Microsoft.Office.Interop.Word.Window, Microsoft.Office.Tools.CustomTaskPane> Dicionario_Window_e_TaskPane = new Dictionary<Microsoft.Office.Interop.Word.Window, Microsoft.Office.Tools.CustomTaskPane>();
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //System.Windows.Forms.MessageBox.Show("ThisAddIn_Startup");

            //Lê e configura preferências
            iClass_RibbonControls.Le_preferencias(Ribbon.Variables.Caminho_preferences);
            iClass_RibbonControls.Configura_Valores_iniciais();

            //Inicia Eventos
            //iClass_New_or_Open_Event.Evento_New_or_Open();
            //if (Globals.ThisAddIn.Application.Documents.Count == 1) iClass_New_or_Open_Event.Metodo_New_or_Open(Globals.ThisAddIn.Application.ActiveDocument);
            //iClass_DocumentClose_Event.Evento_DocumentClose();
            iClass_DocSave_Event.Evento_DocSave();
            iClass_SelectionChange_Event.Evento_SelectionChange();
            iClass_WindowActivate_Event.Evento_WindowActivate();
            if (Globals.ThisAddIn.Application.Documents.Count > 0) iClass_WindowActivate_Event.Metodo_WindowActivate(null, Globals.ThisAddIn.Application.ActiveWindow); // Para adicionar a Taskpane quando abro o Word direto no documento (sem passar pelo BackStage)
            iClass_WindowDeactivate_Event.Evento_WindowDeactivate();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (File.Exists(Ribbon.Variables.Caminho_template))
            {
                Globals.ThisAddIn.Application.AddIns.Unload(true);
                try { File.Delete(Ribbon.Variables.Caminho_template); } catch (IOException) { }
                try { Escreve_preferencias(Ribbon.Variables.Caminho_preferences); } catch (IOException) { }
            }

            Excluir_arquivos_da_lista(Ribbon.Variables.Lista_arquivos_para_excluir);
            Atualiza_txt_lista_de_arquivos_para_excluir(Ribbon.Variables.Lista_arquivos_para_excluir);
        }

        public static void Excluir_arquivos_da_lista(List<string> lista_arquivos)
        {
            foreach (var arquivo in new List<string>(lista_arquivos))
            {
                if (File.Exists(arquivo))
                {
                    try
                    {
                        File.Delete(arquivo);
                        lista_arquivos.Remove(arquivo);
                    }
                    catch (IOException) { }
                }
                else
                {
                    lista_arquivos.Remove(arquivo);
                }
            }
        }

       private void Atualiza_txt_lista_de_arquivos_para_excluir(List<string> lista_arquivos)
        {
            if (lista_arquivos.Count == 0)
            {
                File.Delete(Ribbon.Variables.Caminho_lista_de_arquivos_para_excluir);
            }
            else
            {
                File.WriteAllLines(Ribbon.Variables.Caminho_lista_de_arquivos_para_excluir, lista_arquivos);
            }
        }

        private void Escreve_preferencias(string caminho_preferences)
        {
            if (!Directory.Exists(Ribbon.Variables.Caminho_AppData_Roaming_PeriTAB))
            {
                Directory.CreateDirectory(Ribbon.Variables.Caminho_AppData_Roaming_PeriTAB);
            }

            // Cria um dicionário de preferências
            Dictionary<string, string> preferencias = new Dictionary<string, string>
            {
                 { "unidade", Globals.Ribbons.Ribbon.DropDown_unidade.SelectedItem.Label }
                ,{ "precisao", Globals.Ribbons.Ribbon.DropDown_precisao.SelectedItem.Label }
                ,{ "painel_de_estilos", Globals.Ribbons.Ribbon.ToggleButton_painel_de_estilos.Checked.ToString() }
                ,{ "largura_checked", Globals.Ribbons.Ribbon.CheckBox_largura.Checked.ToString() }
                ,{ "largura", string.IsNullOrEmpty(Globals.Ribbons.Ribbon.EditBox_largura.Text) ? Class_RibbonControls.Retorna_preferencia("largura") : Globals.Ribbons.Ribbon.EditBox_largura.Text } // Verifica e define valores para largura se for vazio ou null (Ribbon.Variables.editBox_largura_Text)
                ,{ "altura", string.IsNullOrEmpty(Globals.Ribbons.Ribbon.EditBox_altura.Text) ? Class_RibbonControls.Retorna_preferencia("altura") : Globals.Ribbons.Ribbon.EditBox_altura.Text }// Verifica e define valores para altura se for vazio ou null (Ribbon.Variables.editBox_altura_Text)
                ,{ "separador", Globals.Ribbons.Ribbon.DropDown_separador.SelectedItem.Label }
                ,{ "assinar_pdf", Globals.Ribbons.Ribbon.CheckBox_assinar.Checked.ToString() }
                ,{ "abrir_pdf", Globals.Ribbons.Ribbon.CheckBox_abrir.Checked.ToString() }
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
