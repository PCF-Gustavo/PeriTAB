using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
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
            button_renomeia_documento_Default();
            button_atualiza_campos_Default();
            button_moeda_Default();
            button_inteiro_Default();
            button_importa_estilos_Default();
            button_limpa_estilos_Default();
            button_cola_imagem_Default();
            checkBox_altura_Default();
            checkBox_largura_Default();
            editBox_largura_Default();
            editBox_altura_Default();
            dropDown_ordem_Default();
            dropDown_separador_Default();
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
        public void button_renomeia_documento_Default()
        {
            Globals.Ribbons.Ribbon1.button_renomeia_documento.Enabled = true;
            Globals.Ribbons.Ribbon1.button_renomeia_documento.ScreenTip = "Macro renomeia_documento";
            Globals.Ribbons.Ribbon1.button_renomeia_documento.SuperTip = "Renomeia o documento atual.";
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

        public void button_cola_imagem_Default()
        {
            Globals.Ribbons.Ribbon1.button_cola_imagem.Enabled = true;
            Globals.Ribbons.Ribbon1.button_cola_imagem.ScreenTip = "";
            Globals.Ribbons.Ribbon1.button_cola_imagem.SuperTip = "Cola imagem do Clipboard.";
        }

        public void checkBox_largura_Default()
        {
            Globals.Ribbons.Ribbon1.checkBox_largura.Checked = preferences.largura_checked;
        }
        public void editBox_largura_Default()
        {
            Globals.Ribbons.Ribbon1.editBox_largura.Enabled = preferences.largura_checked;
            if (Globals.Ribbons.Ribbon1.editBox_largura.Enabled) { Globals.Ribbons.Ribbon1.editBox_largura.Text = preferences.largura; }            
        }

        public void checkBox_altura_Default()
        {
            Globals.Ribbons.Ribbon1.checkBox_altura.Checked = !preferences.largura_checked;
        }

        public void editBox_altura_Default()
        {
            Globals.Ribbons.Ribbon1.editBox_altura.Enabled = !preferences.largura_checked;
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

        public class preferences
        {
            private static string var1, var2, var4, var5;
            private static bool var3;

            public static string largura { get { return var1; } set { var1 = value; } }
            public static string altura { get { return var2; } set { var2 = value; } }
            public static bool largura_checked { get { return var3; } set { var3 = value; } }
            public static string ordem { get { return var4; } set { var4 = value; } }
            public static string separador { get { return var5; } set { var5 = value; } }
        }


    }
}
