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
        public UserControl1 iUserControl1;
        public Microsoft.Office.Tools.CustomTaskPane TaskPane1;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            le_preferencias();

            //Configura o Task Pane
            iUserControl1 = new UserControl1();
            TaskPane1 = Globals.ThisAddIn.CustomTaskPanes.Add(iUserControl1, "Painel de Estilos (PeriTAB)");
            TaskPane1.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
            TaskPane1.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
            TaskPane1.Height = 90;
            TaskPane1.VisibleChanged += MyCustomTaskPane_VisibleChanged;


            //Inicia Eventos            
            Class_New_or_Open_Event iClass_New_or_Open_Event = new Class_New_or_Open_Event(); iClass_New_or_Open_Event.Evento_New_or_Open();
            Class_AnyButtonClick_Event iClass_AnyButtonClick_Event = new Class_AnyButtonClick_Event(); iClass_AnyButtonClick_Event.Evento_AnyButtonClick();
            Class_Buttons iClass_Buttons = new Class_Buttons(); iClass_Buttons.DefaultAll();
            Class_DocSave_Event iClass_DocSave_Event = new Class_DocSave_Event(); iClass_DocSave_Event.Evento_DocSave();            
            Class_SelectionChange_Event iClass_SelectionChange_Event = new Class_SelectionChange_Event(); iClass_SelectionChange_Event.Evento_SelectionChange();
            Class_WindowActivate_Event iClass_WindowActivate_Event = new Class_WindowActivate_Event(); iClass_WindowActivate_Event.Evento_WindowActivate();
            Class_WindowDeactivate_Event iClass_WindowDeactivate_Event = new Class_WindowDeactivate_Event(); iClass_WindowDeactivate_Event.Evento_WindowDeactivate();

        }

        private void MyCustomTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.TaskPane1.Visible == false & Globals.Ribbons.Ribbon1.toggleButton_estilos.Checked == true) { Globals.Ribbons.Ribbon1.toggleButton_estilos.Checked = false; }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            escreve_preferencias();
        }

        private void escreve_preferencias()
        {
            if (!Directory.Exists(Ribbon1.Variables.caminho_AppData_Roaming_PeriTAB)) { Directory.CreateDirectory(Ribbon1.Variables.caminho_AppData_Roaming_PeriTAB); } //Cria a pasta AppData/Roaming/PeriTAB caso não exista

            string preferences_path = Path.Combine(Ribbon1.Variables.caminho_AppData_Roaming_PeriTAB, "preferences");

            string preferences = "";
            
            if (Globals.Ribbons.Ribbon1.editBox_largura.Text != "") { preferences += "<largura>" + Globals.Ribbons.Ribbon1.editBox_largura.Text + "</largura>" + System.Environment.NewLine; } else if (Ribbon1.Variables.editBox_largura_Text != "") { preferences += "<largura>" + Ribbon1.Variables.editBox_largura_Text + "</largura>" + System.Environment.NewLine; } else { preferences += "<largura>" + Class_Buttons.preferences.largura + "</largura>" + System.Environment.NewLine; }
            if (Globals.Ribbons.Ribbon1.editBox_altura.Text != "") { preferences += "<altura>" + Globals.Ribbons.Ribbon1.editBox_altura.Text + "</altura>" + System.Environment.NewLine; } else if (Ribbon1.Variables.editBox_altura_Text != "") { preferences += "<altura>" + Ribbon1.Variables.editBox_altura_Text + "</altura>" + System.Environment.NewLine; } else { preferences += "<altura>" + Class_Buttons.preferences.altura + "</altura>" + System.Environment.NewLine; }
            preferences += "<largura_checked>" + Globals.Ribbons.Ribbon1.checkBox_largura.Checked.ToString() + "</largura_checked>" + System.Environment.NewLine;
            preferences += "<ordem>" + Globals.Ribbons.Ribbon1.dropDown_ordem.SelectedItem.Label + "</ordem>" + System.Environment.NewLine;
            preferences += "<separador>" + Globals.Ribbons.Ribbon1.dropDown_separador.SelectedItem.Label + "</separador>" + System.Environment.NewLine;
            preferences += "<painel_de_estilos>" + Globals.Ribbons.Ribbon1.toggleButton_estilos.Checked.ToString() + "</painel_de_estilos>" + System.Environment.NewLine;

            File.WriteAllText(preferences_path, preferences);
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
                Class_Buttons.preferences.ordem = procura(preferences_text, "ordem");
                Class_Buttons.preferences.separador = procura(preferences_text, "separador");
                Class_Buttons.preferences.painel_de_estilos = procura(preferences_text, "painel_de_estilos");
            }
            else
            { // Preferências iniciais
                Class_Buttons.preferences.largura = "10";
                Class_Buttons.preferences.altura = "10";
                Class_Buttons.preferences.largura_checked = "true";
                Class_Buttons.preferences.ordem = "Alfabética";
                Class_Buttons.preferences.separador = "Nenhum";
                Class_Buttons.preferences.painel_de_estilos = "false";
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
