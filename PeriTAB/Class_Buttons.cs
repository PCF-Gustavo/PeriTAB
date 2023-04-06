using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PeriTAB
{
    internal class Class_Buttons
    {        
        public void Default()
        {
            Globals.Ribbons.Ribbon1.button1.Enabled = true;
            Globals.Ribbons.Ribbon1.button1.ScreenTip = "Macro confere_numeracao_legendas";
            Globals.Ribbons.Ribbon1.button1.SuperTip = "Corrige erros de numeração em Figuras, Tabelas etc.";

            Globals.Ribbons.Ribbon1.button2.Enabled = true;
            Globals.Ribbons.Ribbon1.button2.ScreenTip = "Macro alinha_legenda";
            Globals.Ribbons.Ribbon1.button2.SuperTip = "Alinha legenda de Figuras, Tabelas etc.";

            Globals.Ribbons.Ribbon1.button3.Enabled = true;
            Globals.Ribbons.Ribbon1.button3.ScreenTip = "Macro renomeia_documento";
            Globals.Ribbons.Ribbon1.button3.SuperTip = "Renomeia o documento atual";





            Globals.Ribbons.Ribbon1.button6.Enabled = true;
            Globals.Ribbons.Ribbon1.button6.ScreenTip = "Macro atualiza_todos_campos";
            Globals.Ribbons.Ribbon1.button6.SuperTip = "Atualiza os campos (MS Word Fields) do documento";

            Globals.Ribbons.Ribbon1.button7.Enabled = true;
            Globals.Ribbons.Ribbon1.button7.ScreenTip = "Macro moeda_por_extenso";
            Globals.Ribbons.Ribbon1.button7.SuperTip = "Escreve por extenso o valor em Reais. Posicione o cursor ao final do número.";

            Globals.Ribbons.Ribbon1.button8.Enabled = true;
            Globals.Ribbons.Ribbon1.button8.ScreenTip = "Macro moeda_por_extenso";
            Globals.Ribbons.Ribbon1.button8.SuperTip = "Escreve por extenso o número inteiro. Posicione o cursor ao final do número.";

            Globals.Ribbons.Ribbon1.button9.Enabled = true;            
            Globals.Ribbons.Ribbon1.button9.SuperTip = "Importa Estilos de parágrafos";

            Globals.Ribbons.Ribbon1.button10.Enabled = true;
            Globals.Ribbons.Ribbon1.button10.ScreenTip = "Macro limpa_estilos";
            Globals.Ribbons.Ribbon1.button10.SuperTip = "Remove Estilos de parágrafos não utilizados";


        }            
        
    }
}
