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
            //MessageBox.Show("Startup");
            le_preferencias();

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



        private void MyCustomTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            //if (Globals.ThisAddIn.TaskPane1.Visible == false & Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked == true) { Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked = false; }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (File.Exists(Ribbon1.Variables.caminho_template))
            {
                Globals.ThisAddIn.Application.AddIns.Unload(true);
                try { File.Delete(Ribbon1.Variables.caminho_template); } catch (IOException) { }
                try { escreve_preferencias(); } catch (IOException) { }
            }
        }

        private void escreve_preferencias()
        {
            if (!Directory.Exists(Ribbon1.Variables.caminho_AppData_Roaming_PeriTAB)) { Directory.CreateDirectory(Ribbon1.Variables.caminho_AppData_Roaming_PeriTAB); } //Cria a pasta AppData/Roaming/PeriTAB caso não exista

            string preferences_path = Path.Combine(Ribbon1.Variables.caminho_AppData_Roaming_PeriTAB, "preferences");

            string preferences = "";
            
            if (Globals.Ribbons.Ribbon1.editBox_largura.Text != "" & Globals.Ribbons.Ribbon1.editBox_largura.Text != null) 
            { 
                preferences += "<largura>" + Globals.Ribbons.Ribbon1.editBox_largura.Text + "</largura>" + System.Environment.NewLine;
                //MessageBox.Show("1L");
            } 
            else if (Ribbon1.Variables.editBox_largura_Text != "" & Ribbon1.Variables.editBox_largura_Text != null) 
            { 
                preferences += "<largura>" + Ribbon1.Variables.editBox_largura_Text + "</largura>" + System.Environment.NewLine;
                //MessageBox.Show("2L");
            } 
            else 
            { 
                preferences += "<largura>" + Class_Buttons.preferences.largura + "</largura>" + System.Environment.NewLine;
                //MessageBox.Show("3L");
            }

            if (Globals.Ribbons.Ribbon1.editBox_altura.Text != "" & Globals.Ribbons.Ribbon1.editBox_altura.Text != null)
            { 
                preferences += "<altura>" + Globals.Ribbons.Ribbon1.editBox_altura.Text + "</altura>" + System.Environment.NewLine;
                //MessageBox.Show("1A");
            } 
            else if (Ribbon1.Variables.editBox_altura_Text != "" & Ribbon1.Variables.editBox_altura_Text != null) 
            {
                preferences += "<altura>" + Ribbon1.Variables.editBox_altura_Text + "</altura>" + System.Environment.NewLine;
                //MessageBox.Show("2A");
            } 
            else 
            {
                preferences += "<altura>" + Class_Buttons.preferences.altura + "</altura>" + System.Environment.NewLine;
                //MessageBox.Show("3A");
            }
            
            preferences += "<largura_checked>" + Globals.Ribbons.Ribbon1.checkBox_largura.Checked.ToString() + "</largura_checked>" + System.Environment.NewLine;
            //preferences += "<ordem>" + Globals.Ribbons.Ribbon1.dropDown_ordem.SelectedItem.Label + "</ordem>" + System.Environment.NewLine;
            preferences += "<separador>" + Globals.Ribbons.Ribbon1.dropDown_separador.SelectedItem.Label + "</separador>" + System.Environment.NewLine;
            preferences += "<painel_de_estilos>" + Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked.ToString() + "</painel_de_estilos>" + System.Environment.NewLine;
            preferences += "<assinar_pdf>" + Globals.Ribbons.Ribbon1.checkBox_assinar.Checked.ToString() + "</assinar_pdf>" + System.Environment.NewLine;
            preferences += "<abrir_pdf>" + Globals.Ribbons.Ribbon1.checkBox_abrir.Checked.ToString() + "</abrir_pdf>" + System.Environment.NewLine;
            File.WriteAllText(preferences_path, preferences);
            //MessageBox.Show(preferences);
        }

        private void le_preferencias()
        {
            string preferences_path = Path.Combine(Ribbon1.Variables.caminho_AppData_Roaming_PeriTAB, "preferences");

            if (File.Exists(preferences_path))
            {
                string preferences_text = File.ReadAllText(preferences_path);

                Class_Buttons.preferences.largura = procura(preferences_text, "largura");
                Class_Buttons.preferences.altura = procura(preferences_text, "altura");
                Class_Buttons.preferences.largura_checked = procura(preferences_text, "largura_checked");
                //Class_Buttons.preferences.ordem = procura(preferences_text, "ordem");
                Class_Buttons.preferences.separador = procura(preferences_text, "separador");
                Class_Buttons.preferences.painel_de_estilos = procura(preferences_text, "painel_de_estilos");
                Class_Buttons.preferences.assinar_pdf = procura(preferences_text, "assinar_pdf");
                Class_Buttons.preferences.abrir_pdf = procura(preferences_text, "abrir_pdf");
            }
            else
            { // Preferências iniciais
                Class_Buttons.preferences.largura = "10";
                Class_Buttons.preferences.altura = "10";
                Class_Buttons.preferences.largura_checked = "true";
                //Class_Buttons.preferences.ordem = "Alfabética";
                Class_Buttons.preferences.separador = "Nenhum";
                Class_Buttons.preferences.painel_de_estilos = "false";
                Class_Buttons.preferences.assinar_pdf = "true";
                Class_Buttons.preferences.abrir_pdf = "true";
            }
        }

        private string procura(string texto, string valor) 
        {
            string str1 = "<" + valor + ">";
            string str2 = "</" + valor + ">";
            return texto.Substring((texto.IndexOf(str1) + (str1).Length), texto.IndexOf(str2) - (texto.IndexOf(str1) + (str1).Length));
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
