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
      

        }
}
