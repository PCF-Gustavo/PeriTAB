using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using System.Configuration;

using System.ComponentModel;


namespace PeriTAB
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Evento_SettingChanging();
            Class1.Metodo_DocumentOpen();
        }

        public void SettingChanging(object sender, System.EventArgs e)
        {
            MessageBox.Show("SettingChanging");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public void Evento_SettingChanging()
        {
            Properties.Settings.Default.SettingChanging += new System.Configuration.SettingChangingEventHandler(SettingChanging);
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
