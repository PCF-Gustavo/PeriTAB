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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            this.tab_default = this.Factory.CreateRibbonTab();
            this.tab = this.Factory.CreateRibbonTab();
            this.group_porextenso = this.Factory.CreateRibbonGroup();
            this.group_formatacao = this.Factory.CreateRibbonGroup();
            this.group_estilos = this.Factory.CreateRibbonGroup();
            this.group_campos = this.Factory.CreateRibbonGroup();
            this.group_imagem = this.Factory.CreateRibbonGroup();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.checkBox_referencia = this.Factory.CreateRibbonCheckBox();
            this.box1 = this.Factory.CreateRibbonBox();
            this.checkBox_largura = this.Factory.CreateRibbonCheckBox();
            this.editBox_largura = this.Factory.CreateRibbonEditBox();
            this.box2 = this.Factory.CreateRibbonBox();
            this.checkBox_altura = this.Factory.CreateRibbonCheckBox();
            this.editBox_altura = this.Factory.CreateRibbonEditBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.label_multiplas_imagens = this.Factory.CreateRibbonLabel();
            this.dropDown_ordem = this.Factory.CreateRibbonDropDown();
            this.dropDown_separador = this.Factory.CreateRibbonDropDown();
            this.group_tabela = this.Factory.CreateRibbonGroup();
            this.group_conferencia = this.Factory.CreateRibbonGroup();
            this.group_entrega = this.Factory.CreateRibbonGroup();
            this.checkBox_assinar = this.Factory.CreateRibbonCheckBox();
            this.checkBox_abrir = this.Factory.CreateRibbonCheckBox();
            this.group_sobre = this.Factory.CreateRibbonGroup();
            this.label_nome = this.Factory.CreateRibbonLabel();
            this.label_criado = this.Factory.CreateRibbonLabel();
            this.label_email = this.Factory.CreateRibbonLabel();
            this.button_moeda = this.Factory.CreateRibbonButton();
            this.button_inteiro = this.Factory.CreateRibbonButton();
            this.button_alinha_legenda = this.Factory.CreateRibbonButton();
            this.button_numera_paragrafos = this.Factory.CreateRibbonButton();
            this.toggleButton_painel_de_estilos_velho = this.Factory.CreateRibbonToggleButton();
            this.toggleButton_painel_de_estilos = this.Factory.CreateRibbonToggleButton();
            this.button_importa_estilos = this.Factory.CreateRibbonButton();
            this.button_limpa_estilos = this.Factory.CreateRibbonButton();
            this.menu2 = this.Factory.CreateRibbonMenu();
            this.button_inserir_sumario = this.Factory.CreateRibbonButton();
            this.button_inserir_pagina = this.Factory.CreateRibbonButton();
            this.button_inserir_pagina_extenso = this.Factory.CreateRibbonButton();
            this.button_inserir_paginas = this.Factory.CreateRibbonButton();
            this.button_inserir_paginas_extenso = this.Factory.CreateRibbonButton();
            this.button_atualiza_campos = this.Factory.CreateRibbonButton();
            this.button1_separador = this.Factory.CreateRibbonButton();
            this.button2_separador = this.Factory.CreateRibbonButton();
            this.menu1 = this.Factory.CreateRibbonMenu();
            this.checkBox_destaca_campos = this.Factory.CreateRibbonCheckBox();
            this.checkBox_vercodigo_campos = this.Factory.CreateRibbonCheckBox();
            this.checkBox_atualizar_antes_de_imprimir_campos = this.Factory.CreateRibbonCheckBox();
            this.menu_inserir_imagem = this.Factory.CreateRibbonMenu();
            this.button_borda_preta = this.Factory.CreateRibbonButton();
            this.button_borda_vermelha = this.Factory.CreateRibbonButton();
            this.button_borda_amarela = this.Factory.CreateRibbonButton();
            this.menu_remover_imagem = this.Factory.CreateRibbonMenu();
            this.button_remove_borda = this.Factory.CreateRibbonButton();
            this.button_remove_formatacao = this.Factory.CreateRibbonButton();
            this.button_remove_forma = this.Factory.CreateRibbonButton();
            this.button_remove_texto_alt = this.Factory.CreateRibbonButton();
            this.button_remove_imagem = this.Factory.CreateRibbonButton();
            this.button_redimensiona_imagem = this.Factory.CreateRibbonButton();
            this.button_cola_imagem = this.Factory.CreateRibbonButton();
            this.menu_inserir_tabela = this.Factory.CreateRibbonMenu();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.menu_remover_tabela = this.Factory.CreateRibbonMenu();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.button8 = this.Factory.CreateRibbonButton();
            this.button_confere_formatacao = this.Factory.CreateRibbonButton();
            this.button_confere_preambulo = this.Factory.CreateRibbonButton();
            this.button_confere_num_legenda = this.Factory.CreateRibbonButton();
            this.button_renomeia_documento = this.Factory.CreateRibbonButton();
            this.button_gera_pdf = this.Factory.CreateRibbonButton();
            this.button_Subir_SISCRIM = this.Factory.CreateRibbonButton();
            this.tab_default.SuspendLayout();
            this.tab.SuspendLayout();
            this.group_porextenso.SuspendLayout();
            this.group_formatacao.SuspendLayout();
            this.group_estilos.SuspendLayout();
            this.group_campos.SuspendLayout();
            this.group_imagem.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
            this.group_tabela.SuspendLayout();
            this.group_conferencia.SuspendLayout();
            this.group_entrega.SuspendLayout();
            this.group_sobre.SuspendLayout();
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
            this.tab.Groups.Add(this.group_porextenso);
            this.tab.Groups.Add(this.group_formatacao);
            this.tab.Groups.Add(this.group_estilos);
            this.tab.Groups.Add(this.group_campos);
            this.tab.Groups.Add(this.group_imagem);
            this.tab.Groups.Add(this.group_tabela);
            this.tab.Groups.Add(this.group_conferencia);
            this.tab.Groups.Add(this.group_entrega);
            this.tab.Groups.Add(this.group_sobre);
            this.tab.Label = "PeriTAB";
            this.tab.Name = "tab";
            // 
            // group_porextenso
            // 
            this.group_porextenso.Items.Add(this.button_moeda);
            this.group_porextenso.Items.Add(this.button_inteiro);
            this.group_porextenso.Label = "Por Extenso";
            this.group_porextenso.Name = "group_porextenso";
            // 
            // group_formatacao
            // 
            this.group_formatacao.Items.Add(this.button_alinha_legenda);
            this.group_formatacao.Items.Add(this.button_numera_paragrafos);
            this.group_formatacao.Items.Add(this.toggleButton_painel_de_estilos_velho);
            this.group_formatacao.Items.Add(this.toggleButton_painel_de_estilos);
            this.group_formatacao.Label = "Formatação";
            this.group_formatacao.Name = "group_formatacao";
            // 
            // group_estilos
            // 
            this.group_estilos.Items.Add(this.button_importa_estilos);
            this.group_estilos.Items.Add(this.button_limpa_estilos);
            this.group_estilos.Label = "Estilos";
            this.group_estilos.Name = "group_estilos";
            // 
            // group_campos
            // 
            this.group_campos.Items.Add(this.menu2);
            this.group_campos.Items.Add(this.button_atualiza_campos);
            this.group_campos.Items.Add(this.button1_separador);
            this.group_campos.Items.Add(this.button2_separador);
            this.group_campos.Items.Add(this.menu1);
            this.group_campos.Label = "Campos";
            this.group_campos.Name = "group_campos";
            // 
            // group_imagem
            // 
            this.group_imagem.Items.Add(this.menu_inserir_imagem);
            this.group_imagem.Items.Add(this.menu_remover_imagem);
            this.group_imagem.Items.Add(this.separator2);
            this.group_imagem.Items.Add(this.button_redimensiona_imagem);
            this.group_imagem.Items.Add(this.button_cola_imagem);
            this.group_imagem.Items.Add(this.checkBox_referencia);
            this.group_imagem.Items.Add(this.box1);
            this.group_imagem.Items.Add(this.box2);
            this.group_imagem.Items.Add(this.separator1);
            this.group_imagem.Items.Add(this.label_multiplas_imagens);
            this.group_imagem.Items.Add(this.dropDown_ordem);
            this.group_imagem.Items.Add(this.dropDown_separador);
            this.group_imagem.Label = "Assistente de imagem";
            this.group_imagem.Name = "group_imagem";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // checkBox_referencia
            // 
            this.checkBox_referencia.Label = "Referência";
            this.checkBox_referencia.Name = "checkBox_referencia";
            this.checkBox_referencia.SuperTip = "Cola imagem como mera referência ao arquivo original.";
            this.checkBox_referencia.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_referencia_Click);
            // 
            // box1
            // 
            this.box1.Items.Add(this.checkBox_largura);
            this.box1.Items.Add(this.editBox_largura);
            this.box1.Name = "box1";
            // 
            // checkBox_largura
            // 
            this.checkBox_largura.Label = "";
            this.checkBox_largura.Name = "checkBox_largura";
            this.checkBox_largura.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_largura_Click);
            // 
            // editBox_largura
            // 
            this.editBox_largura.Label = "Largura (cm)";
            this.editBox_largura.Name = "editBox_largura";
            this.editBox_largura.SizeString = "00,00";
            this.editBox_largura.Tag = "";
            this.editBox_largura.Text = null;
            this.editBox_largura.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBox_largura_TextChanged);
            // 
            // box2
            // 
            this.box2.Items.Add(this.checkBox_altura);
            this.box2.Items.Add(this.editBox_altura);
            this.box2.Name = "box2";
            // 
            // checkBox_altura
            // 
            this.checkBox_altura.Label = "";
            this.checkBox_altura.Name = "checkBox_altura";
            this.checkBox_altura.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_altura_Click);
            // 
            // editBox_altura
            // 
            this.editBox_altura.Label = "Altura (cm)  ";
            this.editBox_altura.Name = "editBox_altura";
            this.editBox_altura.SizeString = "00,00";
            this.editBox_altura.Text = null;
            this.editBox_altura.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBox_altura_TextChanged);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // label_multiplas_imagens
            // 
            this.label_multiplas_imagens.Label = "Múltiplas Imagens";
            this.label_multiplas_imagens.Name = "label_multiplas_imagens";
            // 
            // dropDown_ordem
            // 
            ribbonDropDownItemImpl1.Label = "Alfabética";
            ribbonDropDownItemImpl2.Label = "Seleção";
            this.dropDown_ordem.Items.Add(ribbonDropDownItemImpl1);
            this.dropDown_ordem.Items.Add(ribbonDropDownItemImpl2);
            this.dropDown_ordem.Label = "Ordem";
            this.dropDown_ordem.Name = "dropDown_ordem";
            this.dropDown_ordem.SizeString = "000000000";
            // 
            // dropDown_separador
            // 
            ribbonDropDownItemImpl3.Label = "Nenhum";
            ribbonDropDownItemImpl4.Label = "Espaço";
            ribbonDropDownItemImpl5.Label = "Parágrafo";
            this.dropDown_separador.Items.Add(ribbonDropDownItemImpl3);
            this.dropDown_separador.Items.Add(ribbonDropDownItemImpl4);
            this.dropDown_separador.Items.Add(ribbonDropDownItemImpl5);
            this.dropDown_separador.Label = "Separador";
            this.dropDown_separador.Name = "dropDown_separador";
            this.dropDown_separador.SizeString = "000000000";
            // 
            // group_tabela
            // 
            this.group_tabela.Items.Add(this.menu_inserir_tabela);
            this.group_tabela.Items.Add(this.menu_remover_tabela);
            this.group_tabela.Label = "Assistente de tabela";
            this.group_tabela.Name = "group_tabela";
            // 
            // group_conferencia
            // 
            this.group_conferencia.Items.Add(this.button_confere_formatacao);
            this.group_conferencia.Items.Add(this.button_confere_preambulo);
            this.group_conferencia.Items.Add(this.button_confere_num_legenda);
            this.group_conferencia.Label = "Conferência";
            this.group_conferencia.Name = "group_conferencia";
            // 
            // group_entrega
            // 
            this.group_entrega.Items.Add(this.button_renomeia_documento);
            this.group_entrega.Items.Add(this.button_gera_pdf);
            this.group_entrega.Items.Add(this.checkBox_assinar);
            this.group_entrega.Items.Add(this.checkBox_abrir);
            this.group_entrega.Items.Add(this.button_Subir_SISCRIM);
            this.group_entrega.Label = "Entrega";
            this.group_entrega.Name = "group_entrega";
            // 
            // checkBox_assinar
            // 
            this.checkBox_assinar.Label = "Assinar PDF";
            this.checkBox_assinar.Name = "checkBox_assinar";
            // 
            // checkBox_abrir
            // 
            this.checkBox_abrir.Label = "Abrir PDF";
            this.checkBox_abrir.Name = "checkBox_abrir";
            // 
            // group_sobre
            // 
            this.group_sobre.Items.Add(this.label_nome);
            this.group_sobre.Items.Add(this.label_criado);
            this.group_sobre.Items.Add(this.label_email);
            this.group_sobre.Label = "Sobre";
            this.group_sobre.Name = "group_sobre";
            // 
            // label_nome
            // 
            this.label_nome.Label = "PeriTAB";
            this.label_nome.Name = "label_nome";
            // 
            // label_criado
            // 
            this.label_criado.Label = "Criado por PCF Gustavo";
            this.label_criado.Name = "label_criado";
            // 
            // label_email
            // 
            this.label_email.Label = "gustavo.gvs@pf.gov.br";
            this.label_email.Name = "label_email";
            // 
            // button_moeda
            // 
            this.button_moeda.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_moeda.Image = global::PeriTAB.Properties.Resources.dinheiro;
            this.button_moeda.Label = "Moeda";
            this.button_moeda.Name = "button_moeda";
            this.button_moeda.ShowImage = true;
            this.button_moeda.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_moeda_Click);
            // 
            // button_inteiro
            // 
            this.button_inteiro.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_inteiro.Image = global::PeriTAB.Properties.Resources.numero;
            this.button_inteiro.Label = "Número inteiro";
            this.button_inteiro.Name = "button_inteiro";
            this.button_inteiro.ShowImage = true;
            this.button_inteiro.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_inteiro_Click);
            // 
            // button_alinha_legenda
            // 
            this.button_alinha_legenda.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_alinha_legenda.Image = global::PeriTAB.Properties.Resources.seta3;
            this.button_alinha_legenda.Label = "Alinha legenda";
            this.button_alinha_legenda.Name = "button_alinha_legenda";
            this.button_alinha_legenda.ShowImage = true;
            this.button_alinha_legenda.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_alinha_legenda_Click);
            // 
            // button_numera_paragrafos
            // 
            this.button_numera_paragrafos.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_numera_paragrafos.Image = global::PeriTAB.Properties.Resources.numerar;
            this.button_numera_paragrafos.Label = "Numera parágrafos";
            this.button_numera_paragrafos.Name = "button_numera_paragrafos";
            this.button_numera_paragrafos.ShowImage = true;
            this.button_numera_paragrafos.Visible = false;
            this.button_numera_paragrafos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_numera_paragrafos_Click);
            // 
            // toggleButton_painel_de_estilos_velho
            // 
            this.toggleButton_painel_de_estilos_velho.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleButton_painel_de_estilos_velho.Image = global::PeriTAB.Properties.Resources.download3;
            this.toggleButton_painel_de_estilos_velho.Label = "Painel de Estilos velho";
            this.toggleButton_painel_de_estilos_velho.Name = "toggleButton_painel_de_estilos_velho";
            this.toggleButton_painel_de_estilos_velho.ShowImage = true;
            this.toggleButton_painel_de_estilos_velho.SuperTip = "Abre Painel de Estilos";
            this.toggleButton_painel_de_estilos_velho.Visible = false;
            this.toggleButton_painel_de_estilos_velho.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButton_painel_de_estilos_Click);
            // 
            // toggleButton_painel_de_estilos
            // 
            this.toggleButton_painel_de_estilos.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleButton_painel_de_estilos.Image = global::PeriTAB.Properties.Resources.download3;
            this.toggleButton_painel_de_estilos.Label = "Painel de Estilos";
            this.toggleButton_painel_de_estilos.Name = "toggleButton_painel_de_estilos";
            this.toggleButton_painel_de_estilos.ShowImage = true;
            this.toggleButton_painel_de_estilos.SuperTip = "Abre Painel de Estilos";
            this.toggleButton_painel_de_estilos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButton_painel_de_estilos_Click);
            // 
            // button_importa_estilos
            // 
            this.button_importa_estilos.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_importa_estilos.Image = global::PeriTAB.Properties.Resources.download2;
            this.button_importa_estilos.Label = "Importa estilos";
            this.button_importa_estilos.Name = "button_importa_estilos";
            this.button_importa_estilos.ShowImage = true;
            this.button_importa_estilos.Visible = false;
            this.button_importa_estilos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_importa_estilos_Click);
            // 
            // button_limpa_estilos
            // 
            this.button_limpa_estilos.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_limpa_estilos.Image = global::PeriTAB.Properties.Resources.vassoura;
            this.button_limpa_estilos.Label = "Limpa estilos";
            this.button_limpa_estilos.Name = "button_limpa_estilos";
            this.button_limpa_estilos.ShowImage = true;
            this.button_limpa_estilos.Visible = false;
            this.button_limpa_estilos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_limpa_estilos_Click);
            // 
            // menu2
            // 
            this.menu2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu2.Items.Add(this.button_inserir_sumario);
            this.menu2.Items.Add(this.button_inserir_pagina);
            this.menu2.Items.Add(this.button_inserir_pagina_extenso);
            this.menu2.Items.Add(this.button_inserir_paginas);
            this.menu2.Items.Add(this.button_inserir_paginas_extenso);
            this.menu2.Label = "Inserir";
            this.menu2.Name = "menu2";
            this.menu2.OfficeImageId = "FieldCodes";
            this.menu2.ShowImage = true;
            // 
            // button_inserir_sumario
            // 
            this.button_inserir_sumario.Label = "Sumário";
            this.button_inserir_sumario.Name = "button_inserir_sumario";
            this.button_inserir_sumario.ShowImage = true;
            this.button_inserir_sumario.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_inserir_sumario_Click);
            // 
            // button_inserir_pagina
            // 
            this.button_inserir_pagina.Label = "Página";
            this.button_inserir_pagina.Name = "button_inserir_pagina";
            this.button_inserir_pagina.ShowImage = true;
            this.button_inserir_pagina.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_inserir_pagina_Click);
            // 
            // button_inserir_pagina_extenso
            // 
            this.button_inserir_pagina_extenso.Label = "Página (extenso)";
            this.button_inserir_pagina_extenso.Name = "button_inserir_pagina_extenso";
            this.button_inserir_pagina_extenso.ShowImage = true;
            this.button_inserir_pagina_extenso.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_inserir_pagina_extenso_Click);
            // 
            // button_inserir_paginas
            // 
            this.button_inserir_paginas.Label = "Páginas";
            this.button_inserir_paginas.Name = "button_inserir_paginas";
            this.button_inserir_paginas.ShowImage = true;
            this.button_inserir_paginas.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_inserir_paginas_Click);
            // 
            // button_inserir_paginas_extenso
            // 
            this.button_inserir_paginas_extenso.Label = "Páginas (extenso)";
            this.button_inserir_paginas_extenso.Name = "button_inserir_paginas_extenso";
            this.button_inserir_paginas_extenso.ShowImage = true;
            this.button_inserir_paginas_extenso.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_inserir_paginas_extenso_Click);
            // 
            // button_atualiza_campos
            // 
            this.button_atualiza_campos.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_atualiza_campos.Image = global::PeriTAB.Properties.Resources.atualizar;
            this.button_atualiza_campos.Label = "Atualiza todos os Campos";
            this.button_atualiza_campos.Name = "button_atualiza_campos";
            this.button_atualiza_campos.ShowImage = true;
            this.button_atualiza_campos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_atualiza_campos_Click);
            // 
            // button1_separador
            // 
            this.button1_separador.Enabled = false;
            this.button1_separador.Label = "button1";
            this.button1_separador.Name = "button1_separador";
            this.button1_separador.ShowLabel = false;
            // 
            // button2_separador
            // 
            this.button2_separador.Enabled = false;
            this.button2_separador.Label = "button2";
            this.button2_separador.Name = "button2_separador";
            this.button2_separador.ShowLabel = false;
            // 
            // menu1
            // 
            this.menu1.Image = global::PeriTAB.Properties.Resources.engrenagem;
            this.menu1.Items.Add(this.checkBox_destaca_campos);
            this.menu1.Items.Add(this.checkBox_vercodigo_campos);
            this.menu1.Items.Add(this.checkBox_atualizar_antes_de_imprimir_campos);
            this.menu1.Label = " ";
            this.menu1.Name = "menu1";
            this.menu1.ShowImage = true;
            // 
            // checkBox_destaca_campos
            // 
            this.checkBox_destaca_campos.Label = "Destacar";
            this.checkBox_destaca_campos.Name = "checkBox_destaca_campos";
            this.checkBox_destaca_campos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_destaca_campos_Click);
            // 
            // checkBox_vercodigo_campos
            // 
            this.checkBox_vercodigo_campos.Label = "Ver código";
            this.checkBox_vercodigo_campos.Name = "checkBox_vercodigo_campos";
            this.checkBox_vercodigo_campos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_vercodigo_campos_Click);
            // 
            // checkBox_atualizar_antes_de_imprimir_campos
            // 
            this.checkBox_atualizar_antes_de_imprimir_campos.Label = "Atualizar antes de imprimir";
            this.checkBox_atualizar_antes_de_imprimir_campos.Name = "checkBox_atualizar_antes_de_imprimir_campos";
            this.checkBox_atualizar_antes_de_imprimir_campos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_atualizar_antes_de_imprimir_campos_Click);
            // 
            // menu_inserir_imagem
            // 
            this.menu_inserir_imagem.Image = global::PeriTAB.Properties.Resources._;
            this.menu_inserir_imagem.Items.Add(this.button_borda_preta);
            this.menu_inserir_imagem.Items.Add(this.button_borda_vermelha);
            this.menu_inserir_imagem.Items.Add(this.button_borda_amarela);
            this.menu_inserir_imagem.Label = "Inserir";
            this.menu_inserir_imagem.Name = "menu_inserir_imagem";
            this.menu_inserir_imagem.ShowImage = true;
            // 
            // button_borda_preta
            // 
            this.button_borda_preta.Image = global::PeriTAB.Properties.Resources.preto;
            this.button_borda_preta.Label = "Borda preta 0,5 pt";
            this.button_borda_preta.Name = "button_borda_preta";
            this.button_borda_preta.ShowImage = true;
            this.button_borda_preta.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_borda_preta_Click);
            // 
            // button_borda_vermelha
            // 
            this.button_borda_vermelha.Image = global::PeriTAB.Properties.Resources.vermelho;
            this.button_borda_vermelha.Label = "Borda vermelha 2 pt";
            this.button_borda_vermelha.Name = "button_borda_vermelha";
            this.button_borda_vermelha.ShowImage = true;
            this.button_borda_vermelha.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_borda_vermelha_Click);
            // 
            // button_borda_amarela
            // 
            this.button_borda_amarela.Image = global::PeriTAB.Properties.Resources.amarelo;
            this.button_borda_amarela.Label = "Borda amarela 3 pt";
            this.button_borda_amarela.Name = "button_borda_amarela";
            this.button_borda_amarela.ShowImage = true;
            this.button_borda_amarela.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_borda_amarela_Click);
            // 
            // menu_remover_imagem
            // 
            this.menu_remover_imagem.Image = global::PeriTAB.Properties.Resources.x;
            this.menu_remover_imagem.Items.Add(this.button_remove_borda);
            this.menu_remover_imagem.Items.Add(this.button_remove_formatacao);
            this.menu_remover_imagem.Items.Add(this.button_remove_forma);
            this.menu_remover_imagem.Items.Add(this.button_remove_texto_alt);
            this.menu_remover_imagem.Items.Add(this.button_remove_imagem);
            this.menu_remover_imagem.Label = "Remover";
            this.menu_remover_imagem.Name = "menu_remover_imagem";
            this.menu_remover_imagem.ShowImage = true;
            // 
            // button_remove_borda
            // 
            this.button_remove_borda.Image = global::PeriTAB.Properties.Resources.quadrado;
            this.button_remove_borda.Label = "Borda";
            this.button_remove_borda.Name = "button_remove_borda";
            this.button_remove_borda.ShowImage = true;
            this.button_remove_borda.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_remove_borda_Click);
            // 
            // button_remove_formatacao
            // 
            this.button_remove_formatacao.Label = "Formatação";
            this.button_remove_formatacao.Name = "button_remove_formatacao";
            this.button_remove_formatacao.OfficeImageId = "RestoreImageSize";
            this.button_remove_formatacao.ShowImage = true;
            this.button_remove_formatacao.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_remove_formatacao_Click_1);
            // 
            // button_remove_forma
            // 
            this.button_remove_forma.Label = "Forma";
            this.button_remove_forma.Name = "button_remove_forma";
            this.button_remove_forma.OfficeImageId = "GalleryAllShapesAndTextboxes";
            this.button_remove_forma.ShowImage = true;
            this.button_remove_forma.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_remove_forma_Click);
            // 
            // button_remove_texto_alt
            // 
            this.button_remove_texto_alt.Image = global::PeriTAB.Properties.Resources.cego;
            this.button_remove_texto_alt.Label = "Texto Alt";
            this.button_remove_texto_alt.Name = "button_remove_texto_alt";
            this.button_remove_texto_alt.ShowImage = true;
            this.button_remove_texto_alt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_remove_texto_alt_Click);
            // 
            // button_remove_imagem
            // 
            this.button_remove_imagem.Label = "Imagem";
            this.button_remove_imagem.Name = "button_remove_imagem";
            this.button_remove_imagem.OfficeImageId = "OmsDelete";
            this.button_remove_imagem.ShowImage = true;
            this.button_remove_imagem.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_remove_imagem_Click);
            // 
            // button_redimensiona_imagem
            // 
            this.button_redimensiona_imagem.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_redimensiona_imagem.Image = global::PeriTAB.Properties.Resources.redimensionar;
            this.button_redimensiona_imagem.Label = "Redimensiona Imagem";
            this.button_redimensiona_imagem.Name = "button_redimensiona_imagem";
            this.button_redimensiona_imagem.ShowImage = true;
            this.button_redimensiona_imagem.SuperTip = "Redimensiona as imagens selecionadas.";
            this.button_redimensiona_imagem.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_redimensiona_imagem_Click);
            // 
            // button_cola_imagem
            // 
            this.button_cola_imagem.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_cola_imagem.Image = global::PeriTAB.Properties.Resources.image_icon;
            this.button_cola_imagem.Label = "Cola Imagem";
            this.button_cola_imagem.Name = "button_cola_imagem";
            this.button_cola_imagem.ShowImage = true;
            this.button_cola_imagem.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_cola_imagem_Click);
            // 
            // menu_inserir_tabela
            // 
            this.menu_inserir_tabela.Image = global::PeriTAB.Properties.Resources._;
            this.menu_inserir_tabela.Items.Add(this.button1);
            this.menu_inserir_tabela.Items.Add(this.button2);
            this.menu_inserir_tabela.Items.Add(this.button3);
            this.menu_inserir_tabela.Label = "Inserir";
            this.menu_inserir_tabela.Name = "menu_inserir_tabela";
            this.menu_inserir_tabela.ShowImage = true;
            // 
            // button1
            // 
            this.button1.Image = global::PeriTAB.Properties.Resources.preto;
            this.button1.Label = "Borda preta 0,5 pt";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_borda_preta_Click);
            // 
            // button2
            // 
            this.button2.Image = global::PeriTAB.Properties.Resources.vermelho;
            this.button2.Label = "Borda vermelha 2 pt";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_borda_vermelha_Click);
            // 
            // button3
            // 
            this.button3.Image = global::PeriTAB.Properties.Resources.amarelo;
            this.button3.Label = "Borda amarela 3 pt";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_borda_amarela_Click);
            // 
            // menu_remover_tabela
            // 
            this.menu_remover_tabela.Image = global::PeriTAB.Properties.Resources.x;
            this.menu_remover_tabela.Items.Add(this.button4);
            this.menu_remover_tabela.Items.Add(this.button5);
            this.menu_remover_tabela.Items.Add(this.button6);
            this.menu_remover_tabela.Items.Add(this.button7);
            this.menu_remover_tabela.Items.Add(this.button8);
            this.menu_remover_tabela.Label = "Remover";
            this.menu_remover_tabela.Name = "menu_remover_tabela";
            this.menu_remover_tabela.ShowImage = true;
            // 
            // button4
            // 
            this.button4.Image = global::PeriTAB.Properties.Resources.quadrado;
            this.button4.Label = "Borda";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_remove_borda_Click);
            // 
            // button5
            // 
            this.button5.Label = "Formatação";
            this.button5.Name = "button5";
            this.button5.OfficeImageId = "RestoreImageSize";
            this.button5.ShowImage = true;
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_remove_formatacao_Click_1);
            // 
            // button6
            // 
            this.button6.Label = "Forma";
            this.button6.Name = "button6";
            this.button6.OfficeImageId = "GalleryAllShapesAndTextboxes";
            this.button6.ShowImage = true;
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_remove_forma_Click);
            // 
            // button7
            // 
            this.button7.Image = global::PeriTAB.Properties.Resources.cego;
            this.button7.Label = "Texto Alt";
            this.button7.Name = "button7";
            this.button7.ShowImage = true;
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_remove_texto_alt_Click);
            // 
            // button8
            // 
            this.button8.Label = "Imagem";
            this.button8.Name = "button8";
            this.button8.OfficeImageId = "OmsDelete";
            this.button8.ShowImage = true;
            this.button8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_remove_imagem_Click);
            // 
            // button_confere_formatacao
            // 
            this.button_confere_formatacao.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_confere_formatacao.Image = global::PeriTAB.Properties.Resources.formatacao;
            this.button_confere_formatacao.Label = "Formatação";
            this.button_confere_formatacao.Name = "button_confere_formatacao";
            this.button_confere_formatacao.ShowImage = true;
            this.button_confere_formatacao.Visible = false;
            this.button_confere_formatacao.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_confere_formatacao_Click);
            // 
            // button_confere_preambulo
            // 
            this.button_confere_preambulo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_confere_preambulo.Image = global::PeriTAB.Properties.Resources.checklist2;
            this.button_confere_preambulo.Label = "Preâmbulo";
            this.button_confere_preambulo.Name = "button_confere_preambulo";
            this.button_confere_preambulo.ShowImage = true;
            this.button_confere_preambulo.SuperTip = "Checa as informações do preâmbulo.";
            this.button_confere_preambulo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_confere_preambulo_Click);
            // 
            // button_confere_num_legenda
            // 
            this.button_confere_num_legenda.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_confere_num_legenda.Image = global::PeriTAB.Properties.Resources.lupa;
            this.button_confere_num_legenda.Label = "Numeração das legendas";
            this.button_confere_num_legenda.Name = "button_confere_num_legenda";
            this.button_confere_num_legenda.ShowImage = true;
            this.button_confere_num_legenda.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_confere_num_legenda_Click);
            // 
            // button_renomeia_documento
            // 
            this.button_renomeia_documento.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_renomeia_documento.Image = global::PeriTAB.Properties.Resources.abc;
            this.button_renomeia_documento.Label = "Renomeia documento";
            this.button_renomeia_documento.Name = "button_renomeia_documento";
            this.button_renomeia_documento.ShowImage = true;
            this.button_renomeia_documento.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_renomeia_documento_Click);
            // 
            // button_gera_pdf
            // 
            this.button_gera_pdf.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_gera_pdf.Image = global::PeriTAB.Properties.Resources.icone_pdf2;
            this.button_gera_pdf.Label = "Gera PDF";
            this.button_gera_pdf.Name = "button_gera_pdf";
            this.button_gera_pdf.ShowImage = true;
            this.button_gera_pdf.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_gerar_pdf_Click);
            // 
            // button_Subir_SISCRIM
            // 
            this.button_Subir_SISCRIM.Image = global::PeriTAB.Properties.Resources.subir2;
            this.button_Subir_SISCRIM.Label = "Subir no SISCRIM";
            this.button_Subir_SISCRIM.Name = "button_Subir_SISCRIM";
            this.button_Subir_SISCRIM.ShowImage = true;
            this.button_Subir_SISCRIM.Visible = false;
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
            this.group_porextenso.ResumeLayout(false);
            this.group_porextenso.PerformLayout();
            this.group_formatacao.ResumeLayout(false);
            this.group_formatacao.PerformLayout();
            this.group_estilos.ResumeLayout(false);
            this.group_estilos.PerformLayout();
            this.group_campos.ResumeLayout(false);
            this.group_campos.PerformLayout();
            this.group_imagem.ResumeLayout(false);
            this.group_imagem.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.group_tabela.ResumeLayout(false);
            this.group_tabela.PerformLayout();
            this.group_conferencia.ResumeLayout(false);
            this.group_conferencia.PerformLayout();
            this.group_entrega.ResumeLayout(false);
            this.group_entrega.PerformLayout();
            this.group_sobre.ResumeLayout(false);
            this.group_sobre.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab_default;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_conferencia;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_porextenso;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_estilos;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_sobre;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label_nome;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label_criado;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label_email;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_confere_num_legenda;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_alinha_legenda;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_renomeia_documento;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_inserir_sumario;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_inteiro;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_moeda;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_limpa_estilos;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_importa_estilos;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_destaca_campos;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_vercodigo_campos;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton_painel_de_estilos_velho;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_atualiza_campos;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_campos;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_imagem;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_cola_imagem;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_largura;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown_ordem;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox_largura;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox_altura;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_altura;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_referencia;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label_multiplas_imagens;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown_separador;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_inserir_pagina;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_inserir_pagina_extenso;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_entrega;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_gera_pdf;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_atualizar_antes_de_imprimir_campos;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_assinar;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_abrir;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_redimensiona_imagem;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_confere_preambulo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_inserir_paginas;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_inserir_paginas_extenso;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_numera_paragrafos;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_confere_formatacao;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1_separador;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2_separador;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu_inserir_imagem;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_borda_preta;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_borda_vermelha;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_borda_amarela;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu_remover_imagem;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_remove_borda;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_remove_formatacao;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_remove_forma;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_remove_texto_alt;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_remove_imagem;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_tabela;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu_inserir_tabela;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu_remover_tabela;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_Subir_SISCRIM;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton_painel_de_estilos;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_formatacao;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
