using Microsoft.Office.Interop.Word;
using System.Reflection;
using System;
using System.Runtime.Versioning;

namespace PeriTAB{

    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {

        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Designer de Componentes

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab_default = this.Factory.CreateRibbonTab();
            this.tab = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button7 = this.Factory.CreateRibbonButton();
            this.button8 = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.button9 = this.Factory.CreateRibbonButton();
            this.button10 = this.Factory.CreateRibbonButton();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.label3 = this.Factory.CreateRibbonLabel();
            this.tab_default.SuspendLayout();
            this.tab.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group5.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab_default
            // 
            this.tab_default.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab_default.Label = "TabAddIns";
            this.tab_default.Name = "tab_default";
            // 
            // tab
            // 
            this.tab.Groups.Add(this.group1);
            this.tab.Groups.Add(this.group2);
            this.tab.Groups.Add(this.group3);
            this.tab.Groups.Add(this.group4);
            this.tab.Groups.Add(this.group5);
            this.tab.Label = "PeriTAB";
            this.tab.Name = "tab";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.button2);
            this.group1.Items.Add(this.button3);
            this.group1.Label = "Macros";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::PeriTAB.Properties.Resources.lupa;
            this.button1.Label = "Confere numeração das legendas";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Anybutton_Click);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = global::PeriTAB.Properties.Resources.esquadro;
            this.button2.Label = "Alinha legenda";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Anybutton_Click);
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Image = global::PeriTAB.Properties.Resources.abc;
            this.button3.Label = "Renomeia documento";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Anybutton_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.button4);
            this.group2.Items.Add(this.button5);
            this.group2.Items.Add(this.button6);
            this.group2.Label = "Campos";
            this.group2.Name = "group2";
            // 
            // button4
            // 
            this.button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button4.Image = global::PeriTAB.Properties.Resources.radiografia;
            this.button4.Label = "Alterna visualização dos campos";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Anybutton_Click);
            // 
            // button5
            // 
            this.button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button5.Image = global::PeriTAB.Properties.Resources.aquarela;
            this.button5.Label = "Alterna destaque dos campos";
            this.button5.Name = "button5";
            this.button5.ShowImage = true;
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Anybutton_Click);
            // 
            // button6
            // 
            this.button6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button6.Image = global::PeriTAB.Properties.Resources.atualizar;
            this.button6.Label = "Atualiza todos os campos";
            this.button6.Name = "button6";
            this.button6.ShowImage = true;
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Anybutton_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.button7);
            this.group3.Items.Add(this.button8);
            this.group3.Label = "Por Extenso";
            this.group3.Name = "group3";
            // 
            // button7
            // 
            this.button7.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button7.Image = global::PeriTAB.Properties.Resources.dinheiro;
            this.button7.Label = "Moeda";
            this.button7.Name = "button7";
            this.button7.ShowImage = true;
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Anybutton_Click);
            // 
            // button8
            // 
            this.button8.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button8.Image = global::PeriTAB.Properties.Resources.numero;
            this.button8.Label = "Número inteiro";
            this.button8.Name = "button8";
            this.button8.ShowImage = true;
            this.button8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Anybutton_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.button9);
            this.group4.Items.Add(this.button10);
            this.group4.Label = "Estilos";
            this.group4.Name = "group4";
            // 
            // button9
            // 
            this.button9.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button9.Image = global::PeriTAB.Properties.Resources.download2;
            this.button9.Label = "Importa estilos";
            this.button9.Name = "button9";
            this.button9.ShowImage = true;
            this.button9.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Anybutton_Click);
            // 
            // button10
            // 
            this.button10.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button10.Image = global::PeriTAB.Properties.Resources.vassoura;
            this.button10.Label = "Limpa estilos";
            this.button10.Name = "button10";
            this.button10.ShowImage = true;
            this.button10.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Anybutton_Click);
            // 
            // group5
            // 
            this.group5.Items.Add(this.label1);
            this.group5.Items.Add(this.label2);
            this.group5.Items.Add(this.label3);
            this.group5.Label = "Sobre";
            this.group5.Name = "group5";
            // 
            // label1
            // 
            this.label1.Label = "PeriTAB 1.0.0";
            this.label1.Name = "label1";
            // 
            // label2
            // 
            this.label2.Label = "Criado por PCF Gustavo";
            this.label2.Name = "label2";
            // 
            // label3
            // 
            this.label3.Label = "gustavo.gvs@pf.gov.br";
            this.label3.Name = "label3";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab_default);
            this.Tabs.Add(this.tab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab_default.ResumeLayout(false);
            this.tab_default.PerformLayout();
            this.tab.ResumeLayout(false);
            this.tab.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab_default;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button10;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
