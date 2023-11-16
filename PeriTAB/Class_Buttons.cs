using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace PeriTAB
{
    internal class Class_Buttons
    {        
        public void DefaultAll()
        {
            button_confere_num_legenda_Default();
            button_alinha_legenda_Default();            
            button_atualiza_campos_Default();
            button_moeda_Default();
            button_inteiro_Default();
            button_importa_estilos_Default();
            button_limpa_estilos_Default();
            toggleButton_estilos_Default();
            button_cola_imagem_Default();
            checkBox_altura_Default();
            checkBox_largura_Default();
            editBox_largura_Default();
            editBox_altura_Default();
            dropDown_ordem_Default();
            dropDown_separador_Default();
            button_renomeia_documento_Default();
            button_gera_pdf_Default();
            button_abre_SISCRIM_Default();
        }
        public void button_confere_num_legenda_Default()
        {
            Globals.Ribbons.Ribbon1.button_confere_num_legenda.Enabled = true;
            Globals.Ribbons.Ribbon1.button_confere_num_legenda.ScreenTip = "Macro confere_numeracao_legendas";
            Globals.Ribbons.Ribbon1.button_confere_num_legenda.SuperTip = "Corrige erros de numeração em Figuras, Tabelas etc.";
        }
        public void button_alinha_legenda_Default()
        {
            Globals.Ribbons.Ribbon1.button_alinha_legenda.Enabled = true;
            Globals.Ribbons.Ribbon1.button_alinha_legenda.ScreenTip = "Macro alinha_legenda";
            Globals.Ribbons.Ribbon1.button_alinha_legenda.SuperTip = "Alinha legenda de Figuras, Tabelas etc.";
        }

        public void button_atualiza_campos_Default()
        {
            Globals.Ribbons.Ribbon1.button_atualiza_campos.Enabled = true;
            Globals.Ribbons.Ribbon1.button_atualiza_campos.ScreenTip = "Macro atualiza_todos_campos";
            Globals.Ribbons.Ribbon1.button_atualiza_campos.SuperTip = "Atualiza todos os campos (MS Word Fields) do documento.";
        }
        public void button_moeda_Default()
        {
            Globals.Ribbons.Ribbon1.button_moeda.Enabled = true;
            Globals.Ribbons.Ribbon1.button_moeda.ScreenTip = "Macro moeda_por_extenso";
            Globals.Ribbons.Ribbon1.button_moeda.SuperTip = "Escreve por extenso o valor em Reais. Posicione o cursor ao final do número.";
        }
        public void button_inteiro_Default()
        {
            Globals.Ribbons.Ribbon1.button_inteiro.Enabled = true;
            Globals.Ribbons.Ribbon1.button_inteiro.ScreenTip = "Macro moeda_por_extenso";
            Globals.Ribbons.Ribbon1.button_inteiro.SuperTip = "Escreve por extenso o número inteiro. Posicione o cursor ao final do número.";
        }
        public void button_importa_estilos_Default()
        {
            Globals.Ribbons.Ribbon1.button_importa_estilos.Enabled = true;
            Globals.Ribbons.Ribbon1.button_importa_estilos.SuperTip = "Importa Estilos de parágrafos.";
        }
        public void button_limpa_estilos_Default()
        {
            Globals.Ribbons.Ribbon1.button_limpa_estilos.Enabled = true;
            Globals.Ribbons.Ribbon1.button_limpa_estilos.ScreenTip = "Macro limpa_estilos";
            Globals.Ribbons.Ribbon1.button_limpa_estilos.SuperTip = "Remove Estilos de parágrafos não utilizados.";
        }

        public void toggleButton_estilos_Default()
        {            
            Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked = bool.Parse(preferences.painel_de_estilos);
            //Globals.ThisAddIn.TaskPane1.Visible = bool.Parse(preferences.painel_de_estilos);
        }

        public void button_cola_imagem_Default()
        {
            Globals.Ribbons.Ribbon1.button_cola_imagem.Enabled = true;
            Globals.Ribbons.Ribbon1.button_cola_imagem.ScreenTip = "";
            Globals.Ribbons.Ribbon1.button_cola_imagem.SuperTip = "Cola imagem do Clipboard.";
        }

        public void checkBox_largura_Default()
        {
            Globals.Ribbons.Ribbon1.checkBox_largura.Checked = bool.Parse(preferences.largura_checked);
        }
        public void editBox_largura_Default()
        {
            Globals.Ribbons.Ribbon1.editBox_largura.Enabled = bool.Parse(preferences.largura_checked);
            if (Globals.Ribbons.Ribbon1.editBox_largura.Enabled) { Globals.Ribbons.Ribbon1.editBox_largura.Text = preferences.largura; }            
        }

        public void checkBox_altura_Default()
        {
            Globals.Ribbons.Ribbon1.checkBox_altura.Checked = !bool.Parse(preferences.largura_checked);
        }

        public void editBox_altura_Default()
        {
            Globals.Ribbons.Ribbon1.editBox_altura.Enabled = !bool.Parse(preferences.largura_checked);
            if (Globals.Ribbons.Ribbon1.checkBox_altura.Checked) { Globals.Ribbons.Ribbon1.editBox_altura.Text = preferences.altura; }
        }

        public void dropDown_ordem_Default()
        {
            int index = -1;
            if (preferences.ordem == "Alfabética") { index = 0; }
            if (preferences.ordem == "Seleção") { index = 1; }
            Globals.Ribbons.Ribbon1.dropDown_ordem.SelectedItemIndex = index;
        }

        public void dropDown_separador_Default()
        {
            int index = -1;
            if (preferences.separador == "Nenhum") { index = 0; }
            if (preferences.separador == "Espaço") { index = 1; }
            if (preferences.separador == "Parágrafo") { index = 2; }
            Globals.Ribbons.Ribbon1.dropDown_separador.SelectedItemIndex = index;
        }
        public void button_renomeia_documento_Default()
        {
            Globals.Ribbons.Ribbon1.button_renomeia_documento.Enabled = true;
            Globals.Ribbons.Ribbon1.button_renomeia_documento.ScreenTip = "";
            Globals.Ribbons.Ribbon1.button_renomeia_documento.SuperTip = "Renomeia o documento atual.";
        }
        public void button_gera_pdf_Default()
        {
            Globals.Ribbons.Ribbon1.button_gera_pdf.Enabled = true;
            Globals.Ribbons.Ribbon1.button_gera_pdf.ScreenTip = "";
            Globals.Ribbons.Ribbon1.button_gera_pdf.SuperTip = "Gera o PDF do documento na pasta onde está salvo.";
        }
        public void button_abre_SISCRIM_Default()
        {
            Globals.Ribbons.Ribbon1.button_abre_SISCRIM.Enabled = true;
            Globals.Ribbons.Ribbon1.button_abre_SISCRIM.ScreenTip = "";
            Globals.Ribbons.Ribbon1.button_abre_SISCRIM.SuperTip = "Abre SISCRIM na página do Laudo ou da Requisição.";
            //Globals.Ribbons.Ribbon1.button_abre_SISCRIM.Enabled = false; 
            //Globals.Ribbons.Ribbon1.button_abre_SISCRIM.ScreenTip = "Desabilitado"; 
            //Globals.Ribbons.Ribbon1.button_abre_SISCRIM.SuperTip = "O PDF do laudo ainda não foi gerado.";
        }

        //public void button_gera_pdf_image(bool load)
        //{
        //    if (load) Globals.Ribbons.Ribbon1.button_gera_pdf.Image = Properties.Resources.load_icon_png_7969;
        //    else Globals.Ribbons.Ribbon1.button_gera_pdf.Image = Properties.Resources.icone_pdf2;
        //}
        public void muda_imagem(string botao, System.Drawing.Bitmap imagem)
        {
            switch (botao)
            {
                case "button_atualiza_campos":
                    Globals.Ribbons.Ribbon1.button_atualiza_campos.Image = imagem;
                    break;
                case "button_redimensiona_imagem":
                    Globals.Ribbons.Ribbon1.button_redimensiona_imagem.Image = imagem;
                    break;
                case "button_cola_imagem":
                    Globals.Ribbons.Ribbon1.button_cola_imagem.Image = imagem;
                    break;
                case "button_confere_preambulo":
                    Globals.Ribbons.Ribbon1.button_confere_preambulo.Image = imagem;
                    break;
                case "button_confere_num_legenda":
                    Globals.Ribbons.Ribbon1.button_confere_num_legenda.Image = imagem;
                    break;
                case "button_gera_pdf":
                    Globals.Ribbons.Ribbon1.button_gera_pdf.Image = imagem;
                    break;
                case "menu_inserir_imagem":
                    Globals.Ribbons.Ribbon1.menu_inserir_imagem.Image = imagem;
                    break;
                case "menu_remover_imagem":
                    Globals.Ribbons.Ribbon1.menu_remover_imagem.Image = imagem;
                    break;
                case "menu_formatacao_imagem":
                    Globals.Ribbons.Ribbon1.menu_formatacao_imagem.Image = imagem;
                    break;
                case "menu_inserir_tabela":
                    Globals.Ribbons.Ribbon1.menu_inserir_tabela.Image = imagem;
                    break;
                case "menu_remover_tabela":
                    Globals.Ribbons.Ribbon1.menu_remover_tabela.Image = imagem;
                    break;
                case "menu_formatacao_tabela":
                    Globals.Ribbons.Ribbon1.menu_formatacao_tabela.Image = imagem;
                    break;
                case "menu_formatacao_campos":
                    Globals.Ribbons.Ribbon1.menu_formatacao_campos.Image = imagem;
                    break;
                case "button_Subir_SISCRIM":
                    Globals.Ribbons.Ribbon1.button_abre_SISCRIM.Image = imagem;
                    break;
                    //default:
                    //    break;
            }
        }


        public class preferences
        {
            private static string var1, var2, var3, var4, var5, var6;
            public static string largura { get { return var1; } set { var1 = value; } }
            public static string altura { get { return var2; } set { var2 = value; } }
            public static string largura_checked { get { return var3; } set { var3 = value; } }
            public static string ordem { get { return var4; } set { var4 = value; } }
            public static string separador { get { return var5; } set { var5 = value; } }
            public static string painel_de_estilos { get { return var6; } set { var6 = value; } }
        }


    }
}
