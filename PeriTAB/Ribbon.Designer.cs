using Microsoft.Office.Interop.Word;
using System.Reflection;
using System;
using System.Runtime.Versioning;

namespace PeriTAB{

    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {

        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl10 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl11 = this.Factory.CreateRibbonDropDownItem();
            this.tab_default = this.Factory.CreateRibbonTab();
            this.tab = this.Factory.CreateRibbonTab();
            this.group_porextenso = this.Factory.CreateRibbonGroup();
            this.button_teste = this.Factory.CreateRibbonButton();
            this.button_moeda = this.Factory.CreateRibbonButton();
            this.button_inteiro = this.Factory.CreateRibbonButton();
            this.dropDown_unidade = this.Factory.CreateRibbonDropDown();
            this.dropDown_precisao = this.Factory.CreateRibbonDropDown();
            this.button_massa = this.Factory.CreateRibbonButton();
            this.group_formatacao = this.Factory.CreateRibbonGroup();
            this.button_alinha_legenda = this.Factory.CreateRibbonButton();
            this.button_autoformata_laudo = this.Factory.CreateRibbonButton();
            this.toggleButton_painel_de_estilos = this.Factory.CreateRibbonToggleButton();
            this.button_separador1 = this.Factory.CreateRibbonButton();
            this.button_separador2 = this.Factory.CreateRibbonButton();
            this.menu_formatacao = this.Factory.CreateRibbonMenu();
            this.button_habilita_edicao = this.Factory.CreateRibbonButton();
            this.group_campos = this.Factory.CreateRibbonGroup();
            this.button_adiciona_indicador = this.Factory.CreateRibbonButton();
            this.menu_inserir_campos = this.Factory.CreateRibbonMenu();
            this.button_inserir_sumario = this.Factory.CreateRibbonButton();
            this.button_inserir_pagina = this.Factory.CreateRibbonButton();
            this.button_inserir_pagina_extenso = this.Factory.CreateRibbonButton();
            this.button_inserir_paginas = this.Factory.CreateRibbonButton();
            this.button_inserir_paginas_extenso = this.Factory.CreateRibbonButton();
            this.button_inserir_ano = this.Factory.CreateRibbonButton();
            this.button_atualiza_campos = this.Factory.CreateRibbonButton();
            this.menu_formatacao_campos = this.Factory.CreateRibbonMenu();
            this.button_minuscula_campos = this.Factory.CreateRibbonButton();
            this.button_separador3 = this.Factory.CreateRibbonButton();
            this.button_separador4 = this.Factory.CreateRibbonButton();
            this.menu_campos = this.Factory.CreateRibbonMenu();
            this.checkBox_destaca_campos = this.Factory.CreateRibbonCheckBox();
            this.checkBox_mostra_indicadores = this.Factory.CreateRibbonCheckBox();
            this.checkBox_atualizar_antes_de_imprimir_campos = this.Factory.CreateRibbonCheckBox();
            this.checkBox_vercodigo_campos = this.Factory.CreateRibbonCheckBox();
            this.group_imagem = this.Factory.CreateRibbonGroup();
            this.menu_inserir_imagem = this.Factory.CreateRibbonMenu();
            this.button_borda_preta = this.Factory.CreateRibbonButton();
            this.button_borda_vermelha = this.Factory.CreateRibbonButton();
            this.button_borda_amarela = this.Factory.CreateRibbonButton();
            this.button9 = this.Factory.CreateRibbonButton();
            this.menu_remover_imagem = this.Factory.CreateRibbonMenu();
            this.button_remove_borda = this.Factory.CreateRibbonButton();
            this.button_remove_formatacao = this.Factory.CreateRibbonButton();
            this.button_remove_forma = this.Factory.CreateRibbonButton();
            this.button_remove_texto_alt = this.Factory.CreateRibbonButton();
            this.button_remove_imagem = this.Factory.CreateRibbonButton();
            this.menu_formatacao_imagem = this.Factory.CreateRibbonMenu();
            this.button_estilo_figura = this.Factory.CreateRibbonButton();
            this.button_alinha_legenda_figuras = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.button_cola_imagem = this.Factory.CreateRibbonButton();
            this.dropDown_separador = this.Factory.CreateRibbonDropDown();
            this.box1 = this.Factory.CreateRibbonBox();
            this.checkBox_largura = this.Factory.CreateRibbonCheckBox();
            this.editBox_largura = this.Factory.CreateRibbonEditBox();
            this.box2 = this.Factory.CreateRibbonBox();
            this.checkBox_altura = this.Factory.CreateRibbonCheckBox();
            this.editBox_altura = this.Factory.CreateRibbonEditBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.button_redimensiona_imagem = this.Factory.CreateRibbonButton();
            this.button_autodimensiona_imagem = this.Factory.CreateRibbonButton();
            this.menu_imagem = this.Factory.CreateRibbonMenu();
            this.checkBox_referencia = this.Factory.CreateRibbonCheckBox();
            this.group_conferencia = this.Factory.CreateRibbonGroup();
            this.button_confere_preambulo = this.Factory.CreateRibbonButton();
            this.button_confere_num_legenda = this.Factory.CreateRibbonButton();
            this.group_entrega = this.Factory.CreateRibbonGroup();
            this.button_abre_SISCRIM = this.Factory.CreateRibbonButton();
            this.button_renomeia_documento = this.Factory.CreateRibbonButton();
            this.button_gera_pdf = this.Factory.CreateRibbonButton();
            this.checkBox_assinar = this.Factory.CreateRibbonCheckBox();
            this.checkBox_abrir = this.Factory.CreateRibbonCheckBox();
            this.menu_entrega = this.Factory.CreateRibbonMenu();
            this.checkBox_senha = this.Factory.CreateRibbonCheckBox();
            this.group_sobre = this.Factory.CreateRibbonGroup();
            this.label_nome = this.Factory.CreateRibbonLabel();
            this.label_criado = this.Factory.CreateRibbonLabel();
            this.label_email = this.Factory.CreateRibbonLabel();
            this.tab_default.SuspendLayout();
            this.tab.SuspendLayout();
            this.group_porextenso.SuspendLayout();
            this.group_formatacao.SuspendLayout();
            this.group_campos.SuspendLayout();
            this.group_imagem.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
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
            this.tab.Groups.Add(this.group_campos);
            this.tab.Groups.Add(this.group_imagem);
            this.tab.Groups.Add(this.group_conferencia);
            this.tab.Groups.Add(this.group_entrega);
            this.tab.Groups.Add(this.group_sobre);
            this.tab.Label = "PeriTAB";
            this.tab.Name = "tab";
            // 
            // group_porextenso
            // 
            this.group_porextenso.Items.Add(this.button_teste);
            this.group_porextenso.Items.Add(this.button_moeda);
            this.group_porextenso.Items.Add(this.button_inteiro);
            this.group_porextenso.Items.Add(this.dropDown_unidade);
            this.group_porextenso.Items.Add(this.dropDown_precisao);
            this.group_porextenso.Items.Add(this.button_massa);
            this.group_porextenso.Label = "Por Extenso";
            this.group_porextenso.Name = "group_porextenso";
            // 
            // button_teste
            // 
            this.button_teste.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_teste.Label = "Teste";
            this.button_teste.Name = "button_teste";
            this.button_teste.ShowImage = true;
            this.button_teste.Visible = false;
            this.button_teste.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_teste_Click);
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
            // dropDown_unidade
            // 
            ribbonDropDownItemImpl1.Label = "grama (g)";
            ribbonDropDownItemImpl2.Label = "quilograma (kg)";
            ribbonDropDownItemImpl3.Label = "litro (L)";
            ribbonDropDownItemImpl4.Label = "mililitro (mL)";
            this.dropDown_unidade.Items.Add(ribbonDropDownItemImpl1);
            this.dropDown_unidade.Items.Add(ribbonDropDownItemImpl2);
            this.dropDown_unidade.Items.Add(ribbonDropDownItemImpl3);
            this.dropDown_unidade.Items.Add(ribbonDropDownItemImpl4);
            this.dropDown_unidade.Label = "Unidade";
            this.dropDown_unidade.Name = "dropDown_unidade";
            this.dropDown_unidade.SizeString = "000000000";
            this.dropDown_unidade.Visible = false;
            // 
            // dropDown_precisao
            // 
            ribbonDropDownItemImpl5.Label = "0,0";
            ribbonDropDownItemImpl6.Label = "0,00";
            ribbonDropDownItemImpl7.Label = "0,000";
            this.dropDown_precisao.Items.Add(ribbonDropDownItemImpl5);
            this.dropDown_precisao.Items.Add(ribbonDropDownItemImpl6);
            this.dropDown_precisao.Items.Add(ribbonDropDownItemImpl7);
            this.dropDown_precisao.Label = "Precisão";
            this.dropDown_precisao.Name = "dropDown_precisao";
            this.dropDown_precisao.SizeString = "000000000";
            this.dropDown_precisao.Visible = false;
            // 
            // button_massa
            // 
            this.button_massa.Image = global::PeriTAB.Properties.Resources.peso;
            this.button_massa.Label = "Massa / Volume";
            this.button_massa.Name = "button_massa";
            this.button_massa.ShowImage = true;
            this.button_massa.Visible = false;
            // 
            // group_formatacao
            // 
            this.group_formatacao.Items.Add(this.button_alinha_legenda);
            this.group_formatacao.Items.Add(this.button_autoformata_laudo);
            this.group_formatacao.Items.Add(this.toggleButton_painel_de_estilos);
            this.group_formatacao.Items.Add(this.button_separador1);
            this.group_formatacao.Items.Add(this.button_separador2);
            this.group_formatacao.Items.Add(this.menu_formatacao);
            this.group_formatacao.Label = "Formatação";
            this.group_formatacao.Name = "group_formatacao";
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
            // button_autoformata_laudo
            // 
            this.button_autoformata_laudo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_autoformata_laudo.Image = global::PeriTAB.Properties.Resources.checklist2;
            this.button_autoformata_laudo.Label = "Autoformata Laudo";
            this.button_autoformata_laudo.Name = "button_autoformata_laudo";
            this.button_autoformata_laudo.ShowImage = true;
            this.button_autoformata_laudo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_autoformata_laudo_Click);
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
            // button_separador1
            // 
            this.button_separador1.Label = " ";
            this.button_separador1.Name = "button_separador1";
            // 
            // button_separador2
            // 
            this.button_separador2.Enabled = false;
            this.button_separador2.Label = " ";
            this.button_separador2.Name = "button_separador2";
            this.button_separador2.ShowLabel = false;
            // 
            // menu_formatacao
            // 
            this.menu_formatacao.Image = global::PeriTAB.Properties.Resources.engrenagem;
            this.menu_formatacao.Items.Add(this.button_habilita_edicao);
            this.menu_formatacao.Label = " ";
            this.menu_formatacao.Name = "menu_formatacao";
            this.menu_formatacao.ShowImage = true;
            // 
            // button_habilita_edicao
            // 
            this.button_habilita_edicao.Image = global::PeriTAB.Properties.Resources.emergencia;
            this.button_habilita_edicao.Label = "Habilitar edição (na seleção)";
            this.button_habilita_edicao.Name = "button_habilita_edicao";
            this.button_habilita_edicao.ShowImage = true;
            this.button_habilita_edicao.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_habilita_edicao_Click);
            // 
            // group_campos
            // 
            this.group_campos.Items.Add(this.button_adiciona_indicador);
            this.group_campos.Items.Add(this.menu_inserir_campos);
            this.group_campos.Items.Add(this.button_atualiza_campos);
            this.group_campos.Items.Add(this.menu_formatacao_campos);
            this.group_campos.Items.Add(this.button_separador3);
            this.group_campos.Items.Add(this.button_separador4);
            this.group_campos.Items.Add(this.menu_campos);
            this.group_campos.Label = "Campos";
            this.group_campos.Name = "group_campos";
            // 
            // button_adiciona_indicador
            // 
            this.button_adiciona_indicador.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_adiciona_indicador.Label = "Adiciona indicador";
            this.button_adiciona_indicador.Name = "button_adiciona_indicador";
            this.button_adiciona_indicador.ShowImage = true;
            this.button_adiciona_indicador.Visible = false;
            this.button_adiciona_indicador.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_adiciona_indicador_Click);
            // 
            // menu_inserir_campos
            // 
            this.menu_inserir_campos.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu_inserir_campos.Items.Add(this.button_inserir_sumario);
            this.menu_inserir_campos.Items.Add(this.button_inserir_pagina);
            this.menu_inserir_campos.Items.Add(this.button_inserir_pagina_extenso);
            this.menu_inserir_campos.Items.Add(this.button_inserir_paginas);
            this.menu_inserir_campos.Items.Add(this.button_inserir_paginas_extenso);
            this.menu_inserir_campos.Items.Add(this.button_inserir_ano);
            this.menu_inserir_campos.Label = "Inserir";
            this.menu_inserir_campos.Name = "menu_inserir_campos";
            this.menu_inserir_campos.OfficeImageId = "FieldCodes";
            this.menu_inserir_campos.ShowImage = true;
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
            this.button_inserir_pagina.Label = "Página atual (número)";
            this.button_inserir_pagina.Name = "button_inserir_pagina";
            this.button_inserir_pagina.ShowImage = true;
            this.button_inserir_pagina.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_inserir_pagina_Click);
            // 
            // button_inserir_pagina_extenso
            // 
            this.button_inserir_pagina_extenso.Label = "Página atual (extenso)";
            this.button_inserir_pagina_extenso.Name = "button_inserir_pagina_extenso";
            this.button_inserir_pagina_extenso.ShowImage = true;
            this.button_inserir_pagina_extenso.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_inserir_pagina_extenso_Click);
            // 
            // button_inserir_paginas
            // 
            this.button_inserir_paginas.Label = "Número de páginas (número)";
            this.button_inserir_paginas.Name = "button_inserir_paginas";
            this.button_inserir_paginas.ShowImage = true;
            this.button_inserir_paginas.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_inserir_paginas_Click);
            // 
            // button_inserir_paginas_extenso
            // 
            this.button_inserir_paginas_extenso.Label = "Número de páginas (extenso)";
            this.button_inserir_paginas_extenso.Name = "button_inserir_paginas_extenso";
            this.button_inserir_paginas_extenso.ShowImage = true;
            this.button_inserir_paginas_extenso.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_inserir_paginas_extenso_Click);
            // 
            // button_inserir_ano
            // 
            this.button_inserir_ano.Label = "Ano corrente";
            this.button_inserir_ano.Name = "button_inserir_ano";
            this.button_inserir_ano.ShowImage = true;
            this.button_inserir_ano.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_inserir_ano_Click);
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
            // menu_formatacao_campos
            // 
            this.menu_formatacao_campos.Image = global::PeriTAB.Properties.Resources.formatacao2;
            this.menu_formatacao_campos.Items.Add(this.button_minuscula_campos);
            this.menu_formatacao_campos.Label = "Formatação";
            this.menu_formatacao_campos.Name = "menu_formatacao_campos";
            this.menu_formatacao_campos.ShowImage = true;
            this.menu_formatacao_campos.Visible = false;
            // 
            // button_minuscula_campos
            // 
            this.button_minuscula_campos.Label = "minúscula";
            this.button_minuscula_campos.Name = "button_minuscula_campos";
            this.button_minuscula_campos.ShowImage = true;
            this.button_minuscula_campos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_minuscula_campos_Click);
            // 
            // button_separador3
            // 
            this.button_separador3.Label = " ";
            this.button_separador3.Name = "button_separador3";
            // 
            // button_separador4
            // 
            this.button_separador4.Enabled = false;
            this.button_separador4.Label = " ";
            this.button_separador4.Name = "button_separador4";
            this.button_separador4.ShowLabel = false;
            // 
            // menu_campos
            // 
            this.menu_campos.Image = global::PeriTAB.Properties.Resources.engrenagem;
            this.menu_campos.Items.Add(this.checkBox_destaca_campos);
            this.menu_campos.Items.Add(this.checkBox_mostra_indicadores);
            this.menu_campos.Items.Add(this.checkBox_atualizar_antes_de_imprimir_campos);
            this.menu_campos.Items.Add(this.checkBox_vercodigo_campos);
            this.menu_campos.Label = " ";
            this.menu_campos.Name = "menu_campos";
            this.menu_campos.ShowImage = true;
            // 
            // checkBox_destaca_campos
            // 
            this.checkBox_destaca_campos.Label = "Destacar campos";
            this.checkBox_destaca_campos.Name = "checkBox_destaca_campos";
            this.checkBox_destaca_campos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_destaca_campos_Click);
            // 
            // checkBox_mostra_indicadores
            // 
            this.checkBox_mostra_indicadores.Label = "Mostrar indicadores";
            this.checkBox_mostra_indicadores.Name = "checkBox_mostra_indicadores";
            this.checkBox_mostra_indicadores.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_mostra_indicadores_Click);
            // 
            // checkBox_atualizar_antes_de_imprimir_campos
            // 
            this.checkBox_atualizar_antes_de_imprimir_campos.Label = "Atualizar antes de imprimir";
            this.checkBox_atualizar_antes_de_imprimir_campos.Name = "checkBox_atualizar_antes_de_imprimir_campos";
            this.checkBox_atualizar_antes_de_imprimir_campos.Visible = false;
            this.checkBox_atualizar_antes_de_imprimir_campos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_atualizar_antes_de_imprimir_campos_Click);
            // 
            // checkBox_vercodigo_campos
            // 
            this.checkBox_vercodigo_campos.Label = "Ver código";
            this.checkBox_vercodigo_campos.Name = "checkBox_vercodigo_campos";
            this.checkBox_vercodigo_campos.Visible = false;
            this.checkBox_vercodigo_campos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_vercodigo_campos_Click);
            // 
            // group_imagem
            // 
            this.group_imagem.Items.Add(this.menu_inserir_imagem);
            this.group_imagem.Items.Add(this.menu_remover_imagem);
            this.group_imagem.Items.Add(this.menu_formatacao_imagem);
            this.group_imagem.Items.Add(this.separator2);
            this.group_imagem.Items.Add(this.button_cola_imagem);
            this.group_imagem.Items.Add(this.dropDown_separador);
            this.group_imagem.Items.Add(this.box1);
            this.group_imagem.Items.Add(this.box2);
            this.group_imagem.Items.Add(this.separator1);
            this.group_imagem.Items.Add(this.button_redimensiona_imagem);
            this.group_imagem.Items.Add(this.button_autodimensiona_imagem);
            this.group_imagem.Items.Add(this.menu_imagem);
            this.group_imagem.Label = "Assistente de imagem";
            this.group_imagem.Name = "group_imagem";
            // 
            // menu_inserir_imagem
            // 
            this.menu_inserir_imagem.Image = global::PeriTAB.Properties.Resources._;
            this.menu_inserir_imagem.Items.Add(this.button_borda_preta);
            this.menu_inserir_imagem.Items.Add(this.button_borda_vermelha);
            this.menu_inserir_imagem.Items.Add(this.button_borda_amarela);
            this.menu_inserir_imagem.Items.Add(this.button9);
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
            // button9
            // 
            this.button9.Image = global::PeriTAB.Properties.Resources.abc;
            this.button9.Label = "Legenda";
            this.button9.Name = "button9";
            this.button9.ShowImage = true;
            this.button9.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_legenda_imagem_Click);
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
            // menu_formatacao_imagem
            // 
            this.menu_formatacao_imagem.Image = global::PeriTAB.Properties.Resources.formatacao2;
            this.menu_formatacao_imagem.Items.Add(this.button_estilo_figura);
            this.menu_formatacao_imagem.Items.Add(this.button_alinha_legenda_figuras);
            this.menu_formatacao_imagem.Label = "Formatação";
            this.menu_formatacao_imagem.Name = "menu_formatacao_imagem";
            this.menu_formatacao_imagem.ShowImage = true;
            this.menu_formatacao_imagem.Visible = false;
            // 
            // button_estilo_figura
            // 
            this.button_estilo_figura.Label = "Estilo Figura";
            this.button_estilo_figura.Name = "button_estilo_figura";
            this.button_estilo_figura.ShowImage = true;
            // 
            // button_alinha_legenda_figuras
            // 
            this.button_alinha_legenda_figuras.Label = "Alinha legendas de figuras";
            this.button_alinha_legenda_figuras.Name = "button_alinha_legenda_figuras";
            this.button_alinha_legenda_figuras.ShowImage = true;
            this.button_alinha_legenda_figuras.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_alinha_legenda_figuras_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
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
            // dropDown_separador
            // 
            ribbonDropDownItemImpl8.Label = "Nenhum";
            ribbonDropDownItemImpl9.Label = "Espaço";
            ribbonDropDownItemImpl10.Label = "Parágrafo";
            ribbonDropDownItemImpl11.Label = "Parágrafo + 3pt";
            this.dropDown_separador.Items.Add(ribbonDropDownItemImpl8);
            this.dropDown_separador.Items.Add(ribbonDropDownItemImpl9);
            this.dropDown_separador.Items.Add(ribbonDropDownItemImpl10);
            this.dropDown_separador.Items.Add(ribbonDropDownItemImpl11);
            this.dropDown_separador.Label = "Separador";
            this.dropDown_separador.Name = "dropDown_separador";
            this.dropDown_separador.SizeString = "Parágrafo + 3pt";
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
            this.editBox_altura.Label = "Altura (cm)";
            this.editBox_altura.Name = "editBox_altura";
            this.editBox_altura.SizeString = "00,00";
            this.editBox_altura.Text = null;
            this.editBox_altura.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBox_altura_TextChanged);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // button_redimensiona_imagem
            // 
            this.button_redimensiona_imagem.Image = global::PeriTAB.Properties.Resources.redimensionar2;
            this.button_redimensiona_imagem.Label = "Redimensiona";
            this.button_redimensiona_imagem.Name = "button_redimensiona_imagem";
            this.button_redimensiona_imagem.ShowImage = true;
            this.button_redimensiona_imagem.SuperTip = "Redimensiona as imagens selecionadas conforme o tamanho digitado.";
            this.button_redimensiona_imagem.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_redimensiona_imagem_Click);
            // 
            // button_autodimensiona_imagem
            // 
            this.button_autodimensiona_imagem.Image = global::PeriTAB.Properties.Resources.redimensionar3;
            this.button_autodimensiona_imagem.Label = "Autodimensiona";
            this.button_autodimensiona_imagem.Name = "button_autodimensiona_imagem";
            this.button_autodimensiona_imagem.ShowImage = true;
            this.button_autodimensiona_imagem.SuperTip = "Redimensiona as imagens selecionadas para o tamanho da linha.";
            this.button_autodimensiona_imagem.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_autodimensiona_imagem_Click);
            // 
            // menu_imagem
            // 
            this.menu_imagem.Image = global::PeriTAB.Properties.Resources.engrenagem;
            this.menu_imagem.Items.Add(this.checkBox_referencia);
            this.menu_imagem.Label = " ";
            this.menu_imagem.Name = "menu_imagem";
            this.menu_imagem.ShowImage = true;
            // 
            // checkBox_referencia
            // 
            this.checkBox_referencia.Label = "Referência";
            this.checkBox_referencia.Name = "checkBox_referencia";
            this.checkBox_referencia.SuperTip = "Cola imagem como mera referência ao arquivo original.";
            this.checkBox_referencia.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_referencia_Click);
            // 
            // group_conferencia
            // 
            this.group_conferencia.Items.Add(this.button_confere_preambulo);
            this.group_conferencia.Items.Add(this.button_confere_num_legenda);
            this.group_conferencia.Label = "Conferência";
            this.group_conferencia.Name = "group_conferencia";
            // 
            // button_confere_preambulo
            // 
            this.button_confere_preambulo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_confere_preambulo.Image = global::PeriTAB.Properties.Resources.checklist2;
            this.button_confere_preambulo.Label = "Preâmbulo";
            this.button_confere_preambulo.Name = "button_confere_preambulo";
            this.button_confere_preambulo.ShowImage = true;
            this.button_confere_preambulo.SuperTip = "Checa as informações do preâmbulo.";
            this.button_confere_preambulo.Visible = false;
            this.button_confere_preambulo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_confere_preambulo_Click);
            // 
            // button_confere_num_legenda
            // 
            this.button_confere_num_legenda.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_confere_num_legenda.Image = global::PeriTAB.Properties.Resources.lupa;
            this.button_confere_num_legenda.Label = "Numeração das legendas";
            this.button_confere_num_legenda.Name = "button_confere_num_legenda";
            this.button_confere_num_legenda.ShowImage = true;
            this.button_confere_num_legenda.Visible = false;
            this.button_confere_num_legenda.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_confere_num_legenda_Click);
            // 
            // group_entrega
            // 
            this.group_entrega.Items.Add(this.button_abre_SISCRIM);
            this.group_entrega.Items.Add(this.button_renomeia_documento);
            this.group_entrega.Items.Add(this.button_gera_pdf);
            this.group_entrega.Items.Add(this.checkBox_assinar);
            this.group_entrega.Items.Add(this.checkBox_abrir);
            this.group_entrega.Items.Add(this.menu_entrega);
            this.group_entrega.Label = "Entrega";
            this.group_entrega.Name = "group_entrega";
            // 
            // button_abre_SISCRIM
            // 
            this.button_abre_SISCRIM.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_abre_SISCRIM.Image = global::PeriTAB.Properties.Resources.subir2;
            this.button_abre_SISCRIM.Label = "Abre SISCRIM";
            this.button_abre_SISCRIM.Name = "button_abre_SISCRIM";
            this.button_abre_SISCRIM.ShowImage = true;
            this.button_abre_SISCRIM.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_abre_SISCRIM_Click);
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
            this.button_gera_pdf.Image = global::PeriTAB.Properties.Resources.icone_pdf;
            this.button_gera_pdf.Label = "Gera PDF";
            this.button_gera_pdf.Name = "button_gera_pdf";
            this.button_gera_pdf.ShowImage = true;
            this.button_gera_pdf.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_gerar_pdf_Click);
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
            // menu_entrega
            // 
            this.menu_entrega.Image = global::PeriTAB.Properties.Resources.engrenagem;
            this.menu_entrega.Items.Add(this.checkBox_senha);
            this.menu_entrega.Label = " ";
            this.menu_entrega.Name = "menu_entrega";
            this.menu_entrega.ShowImage = true;
            this.menu_entrega.Visible = false;
            // 
            // checkBox_senha
            // 
            this.checkBox_senha.Label = "Manter assinatura ativa";
            this.checkBox_senha.Name = "checkBox_senha";
            this.checkBox_senha.SuperTip = "Marque esta opção para digitar a sua senha apenas uma vez, enquanto o Word não fo" +
    "r fechado.";
            this.checkBox_senha.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox_destaca_campos_Click);
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
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab_default);
            this.Tabs.Add(this.tab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab_default.ResumeLayout(false);
            this.tab_default.PerformLayout();
            this.tab.ResumeLayout(false);
            this.tab.PerformLayout();
            this.group_porextenso.ResumeLayout(false);
            this.group_porextenso.PerformLayout();
            this.group_formatacao.ResumeLayout(false);
            this.group_formatacao.PerformLayout();
            this.group_campos.ResumeLayout(false);
            this.group_campos.PerformLayout();
            this.group_imagem.ResumeLayout(false);
            this.group_imagem.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu_campos;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_destaca_campos;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_vercodigo_campos;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu_inserir_campos;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_atualiza_campos;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_campos;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_imagem;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_cola_imagem;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_largura;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox_largura;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox_altura;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_altura;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_separador4;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_abre_SISCRIM;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton_painel_de_estilos;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_formatacao;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu_formatacao_imagem;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_estilo_figura;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_alinha_legenda_figuras;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_mostra_indicadores;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu_formatacao_campos;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_minuscula_campos;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_separador3;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu_entrega;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_senha;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_massa;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown_unidade;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown_precisao;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_autodimensiona_imagem;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu_imagem;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_referencia;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_teste;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_inserir_ano;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_autoformata_laudo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_habilita_edicao;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu_formatacao;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_separador1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_separador2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_adiciona_indicador;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
