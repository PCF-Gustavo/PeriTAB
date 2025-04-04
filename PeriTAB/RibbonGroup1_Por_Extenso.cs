using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System.Globalization;
using System;
using Tarefa = System.Threading.Tasks.Task;
using System.Windows.Forms;
using System.Linq;
using Org.BouncyCastle.Crypto.Generators;

namespace PeriTAB
{
    public partial class Ribbon
    {
        private static readonly string[] unidade = { "", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove" };
        private static readonly string[] dezena1 = { "dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove" };
        private static readonly string[] dezena2 = { "", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa" };
        private static readonly string[] centena = { "", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos" };
        private static readonly string[] grandezaSingular = { "", "mil", "milhão", "bilhão", "trilhão", "quadrilhão" };
        private static readonly string[] grandezaPlural = { "", "mil", "milhões", "bilhões", "trilhões", "quadrilhões" };

        private async void Por_Extenso_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonButton RibbonButton = (RibbonButton)sender;
            String Botao_Label = RibbonButton.Label;
            System.Drawing.Image RibbonButton_Imagem_inicial = RibbonButton.Image;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;
            bool success = true;
            string msg_Error = string.Empty;


            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                Selection Selecao = Globals.ThisAddIn.Application.Selection;
                Range Selecao_inicial = Selecao.Range.Duplicate;
                Range Selecao_anterior = null;
                string unidade = dropDown_unidade.SelectedItem.Label.Split('(', ')')[1];
                bool unidade_encontrada = false;

                if (Selecao.Text.Count() == 1)
                {
                    while (true)
                    {
                        Selecao_anterior = Selecao.Previous(WdUnits.wdCharacter, 1);
                        if (Selecao_anterior != null)
                        {
                            char caractereAnterior = Selecao_anterior.Text[0];
                            if (!(caractereAnterior == (char)32 || caractereAnterior == (char)10 || caractereAnterior == (char)13 || caractereAnterior == (char)160 || caractereAnterior == (char)9))
                            {
                                Selecao.MoveLeft(WdUnits.wdCharacter, 1, WdMovementType.wdExtend);
                            }
                            else
                            {
                                if (!unidade_encontrada && Botao_Label.Equals("Massa / Volume") && Selecao.Range.Text.EndsWith(dropDown_unidade.SelectedItem.Label.Split('(', ')')[1]))
                                {
                                    while (true)
                                    {
                                        Selecao_anterior = Selecao.Previous(WdUnits.wdCharacter, 1);
                                        if (Selecao_anterior != null)
                                        {
                                            caractereAnterior = Selecao_anterior.Text[0];
                                            if (caractereAnterior == (char)32 || caractereAnterior == (char)160 || caractereAnterior == (char)9)
                                            {
                                                Selecao.MoveLeft(WdUnits.wdCharacter, 1, WdMovementType.wdExtend);
                                            }
                                            else break;
                                        }
                                        else break;
                                    }
                                    unidade_encontrada = true;
                                }
                                else break;
                            }
                        }
                        else break;
                    }
                    String numero_selecionado = Selecao.Text;
                    if (unidade_encontrada) numero_selecionado = numero_selecionado.Replace(unidade,"").Trim();

                    Selecao_anterior = Selecao.Previous(WdUnits.wdCharacter, 1);
                    if (Selecao_anterior != null)
                    {
                        if (Botao_Label.Equals("Moeda") && Selecao.Previous(WdUnits.wdWord, 1).Text.Trim() == "$" && Selecao.Previous(WdUnits.wdWord, 2).Text.Trim() == "R")
                        {
                            Selecao.MoveLeft(WdUnits.wdWord, 2, WdMovementType.wdExtend);
                        }
                    }

                    if (decimal.TryParse(numero_selecionado, out decimal numero))
                    {
                        if (numero >= 1000000000000000000) { success = false; msg_Error = "Número máximo permitido excedido."; }
                        else if (Botao_Label.Equals("Número inteiro") && numero % 1 != 0) { success = false; msg_Error = "Números inteiros não podem ter casas decimais."; }
                        else
                        {
                            string resultado = string.Empty;
                            string formatoNumero = string.Empty;
                            if (Botao_Label.Equals("Moeda"))
                            {
                                numero = Math.Round(numero, 2); // Arredondamento para duas casas decimais
                                resultado = ConverterParaMoeda(numero);
                                formatoNumero = $"R$ {numero:N2}";
                            }
                            if (Botao_Label.Equals("Número inteiro"))
                            {
                                resultado = ConverterParaExtenso((long)numero);
                                formatoNumero = $"{numero:N0}";
                            }
                            if (Botao_Label.Equals("Massa / Volume"))
                            {
                                int precisao = dropDown_precisao.SelectedItem.Label.Split(',')[1].Length;
                                numero = Math.Round(numero, precisao);
                                resultado = ConverterParaMassaVolume(numero,unidade);
                                formatoNumero = $"{numero.ToString($"N{precisao}")} {unidade}";
                            }
                            Selecao.TypeText($"{formatoNumero} ({resultado})");
                        }
                    }
                    else { success = false; msg_Error = "Posicione o cursor ao final de um número válido."; }
                }
                else { success = false; msg_Error = "Posicione o cursor ao final de um número válido."; }

                if (!success)
                {
                    MessageBox.Show(msg_Error, Botao_Label, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Selecao_inicial.Select();
                }

                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                Globals.ThisAddIn.Application.ScreenUpdating = true;
            });

            RibbonButton.Image = RibbonButton_Imagem_inicial;
            RibbonButton.Enabled = true;
        }

        private static string ConverterParaMoeda(decimal numero)
        {
            if (numero == 0) return "zero reais";

            string strNumero = numero.ToString("F2");
            string[] partes = strNumero.Split(',');

            long parteInteira = long.Parse(partes[0]);
            int centavos = int.Parse(partes[1]);

            // **Se for menor que R$ 1,00, retorna apenas os centavos**
            if (parteInteira == 0 && centavos > 0)
            {
                return centavos == 1 ? "um centavo" : $"{ConverterParaExtenso(centavos)} centavos";
            }

            string resultado = ConverterParaExtenso(parteInteira);

            // **Corrige "um reais" para "um real"**
            if (parteInteira == 1)
            {
                resultado = "um real";
            }
            else
            {
                // **Verifica se precisa de "de reais"**
                bool precisaDe = parteInteira % 1_000_000 == 0;
                resultado += precisaDe ? " de reais" : " reais";
            }

            // **Adiciona centavos apenas se houver**
            if (centavos > 0)
            {
                resultado += centavos == 1 ? " e um centavo" : $" e {ConverterParaExtenso(centavos)} centavos";
            }

            return resultado;
        }

        private static string ConverterParaMassaVolume(decimal numero, string unidade)
        {
            if (numero == 0)
            {
                switch (unidade)
                {
                    case "g": return "zero gramas";
                    case "kg": return "zero quilogramas";
                    case "mL": return "zero mililitros";
                    case "L": return "zero litros";
                }
            }

            string resultado = "";

            if (unidade == "g" || unidade == "mL")
            {
                long inteiro = (long)numero;  // Parte inteira (gramas ou mililitros)
                int fracao = (int)Math.Round((numero - inteiro) * 1000);  // Parte decimal (miligramas ou microlitros)

                string unidadeInteira = unidade == "g" ? "grama" : "mililitro";
                string unidadeFracionaria = unidade == "g" ? "miligrama" : "microlitro";

                if (inteiro > 0)
                {
                    bool precisaDe = (inteiro >= 1_000_000) && (inteiro % 1_000_000 == 0);
                    resultado = ConverterParaExtenso(inteiro) + (precisaDe ? " de" : "") + (inteiro == 1 ? $" {unidadeInteira}" : $" {unidadeInteira}s");
                }

                if (fracao > 0)
                {
                    string fracaoTexto = ConverterParaExtenso(fracao) + (fracao == 1 ? $" {unidadeFracionaria}" : $" {unidadeFracionaria}s");
                    resultado += inteiro > 0 ? " e " + fracaoTexto : fracaoTexto;
                }
            }
            else if (unidade == "kg" || unidade == "L")
            {
                long inteiro = (long)numero;  // Parte inteira (quilogramas ou litros)
                int fracao = (int)Math.Round((numero - inteiro) * 1000);  // Parte decimal (gramas ou mililitros)

                string unidadeInteira = unidade == "kg" ? "quilograma" : "litro";
                string unidadeFracionaria = unidade == "kg" ? "grama" : "mililitro";

                if (inteiro > 0)
                {
                    bool precisaDe = (inteiro >= 1_000_000) && (inteiro % 1_000_000 == 0);
                    resultado = ConverterParaExtenso(inteiro) + (precisaDe ? " de" : "") + (inteiro == 1 ? $" {unidadeInteira}" : $" {unidadeInteira}s");
                }

                if (fracao > 0)
                {
                    string fracaoTexto = ConverterParaExtenso(fracao) + (fracao == 1 ? $" {unidadeFracionaria}" : $" {unidadeFracionaria}s");
                    resultado += inteiro > 0 ? " e " + fracaoTexto : fracaoTexto;
                }
            }

            return resultado;
        }




        private static string ConverterParaExtenso(decimal numero)
        {
            if (numero == 0) return "zero";

            string resultado = "";
            long parteInteira = (long)numero;
            int grupo = 0;

            while (parteInteira > 0)
            {
                int parte = (int)(parteInteira % 1000);
                if (parte > 0)
                {
                    string textoGrupo = ConverterGrupo(parte);
                    string grandeza = (parte == 1 && grupo > 1) ? grandezaSingular[grupo] : grandezaPlural[grupo];

                    // ✅ Corrige "mil" sem "um mil"
                    if (grupo == 1 && parte == 1)
                    {
                        textoGrupo = "mil";
                        grandeza = "";
                    }

                    string trecho = $"{textoGrupo} {grandeza}".Trim();

                    if (string.IsNullOrEmpty(resultado))
                    {
                        resultado = trecho;
                    }
                    else
                    {
                        // 🚀 Correção: "e" deve aparecer antes do último grupo se ele for menor que 1000
                        bool precisaDeE = (parte < 1000) && !resultado.Contains(" e ");

                        if (precisaDeE)
                        {
                            resultado = $"{trecho} e {resultado}";
                        }
                        else
                        {
                            resultado = $"{trecho} {resultado}";
                        }
                    }
                }

                parteInteira /= 1000;
                grupo++;
            }

            // ✅ Correção final para evitar "um reais" quando o número for "1 real"
            if (resultado == "um reais")
            {
                resultado = "um real";
            }

            return resultado.Trim();
        }
        private static string ConverterGrupo(int numero)
        {
            if (numero == 100) return "cem";

            string resultado = "";

            if (numero >= 100)
            {
                resultado += centena[numero / 100];
                numero %= 100;
                if (numero > 0) resultado += " e ";
            }

            if (numero >= 10 && numero < 20)
            {
                resultado += dezena1[numero - 10];
            }
            else
            {
                if (numero >= 20)
                {
                    resultado += dezena2[numero / 10];
                    numero %= 10;
                    if (numero > 0) resultado += " e ";
                }

                if (numero > 0)
                {
                    resultado += unidade[numero];
                }
            }

            return resultado;
        }





        private async void button_moeda_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;

            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                Globals.ThisAddIn.Application.Run("moeda_por_extenso");
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            RibbonButton.Image = Properties.Resources.dinheiro;
            RibbonButton.Enabled = true;
        }

        private async void button_inteiro_Click(object sender, RibbonControlEventArgs e)
        {
            // Atualiza a UI na Thread principal
            RibbonButton RibbonButton = (RibbonButton)sender;
            RibbonButton.Image = Properties.Resources.load_icon_png_7969;
            RibbonButton.Enabled = false;

            // Executa as tarefas em segundo plano
            await Tarefa.Run(() =>
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("");
                Globals.ThisAddIn.Application.Run("inteiro_por_extenso");
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            });

            // Após a execução das tarefas, atualiza a UI na Thread principal
            RibbonButton.Image = Properties.Resources.numero;
            RibbonButton.Enabled = true;
        }

        // Função para obter a próxima seção
        
    }
}