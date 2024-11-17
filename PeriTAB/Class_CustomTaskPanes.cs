using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using Tarefa = System.Threading.Tasks.Task;
using System.Windows.Forms;

namespace PeriTAB
{
    public class Class_CustomTaskPanes
    {
        public void Redimensionar(MyUserControl MyUserControl, Microsoft.Office.Tools.CustomTaskPane CustomTaskPane)
        {
            CustomTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
            CustomTaskPane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;

            var screenArea = Screen.PrimaryScreen.WorkingArea;
            float dpiFactor = 96f / Graphics.FromHwnd(IntPtr.Zero).DpiX;

            int taskPaneWidth = screenArea.Width;
            taskPaneWidth = (int)(taskPaneWidth * 0.99); // Reduzir a largura do painel em 1% para evitar a barra de rolagem horizontal

            List<Button> list_botoes = MyUserControl.Controls.OfType<Button>().ToList(); // Obter todos os botões de userControl
            int buttonCount = list_botoes.Count(); // Contar todos os botões de userControl

            // Calcular a largura total disponível para os botões
            int totalSpacingWidth; // Inicializado depois, com base em buttonWidth
            int totalButtonWidth = taskPaneWidth;

            // Definir o espaçamento entre os botões como 1/10 da largura do botão
            int buttonWidth = totalButtonWidth / buttonCount;
            int spacingWidth = buttonWidth / 20; // O espaçamento será 1/10 da largura do botão

            // Calcular novamente o total disponível para os botões após definir spacingWidth
            totalSpacingWidth = spacingWidth * (buttonCount + 1); // Total de espaço entre os botões e as bordas
            totalButtonWidth = taskPaneWidth - totalSpacingWidth; // Largura total disponível para os botões

            // Calcular o tamanho dos botões
            buttonWidth = totalButtonWidth / buttonCount; // Atualizar a largura do botão com base no espaçamento calculado
            int buttonHeight = (int)(buttonWidth / 2); // A altura do botão será metade da largura
            int spacingHeight = (int)(buttonHeight / 10); // Espaçamento vertical entre os botões

            int headerHeight = (int)(50 / dpiFactor); // Ajuste do DPI para altura do cabeçalho
            int taskPaneHeight = buttonHeight + headerHeight + 2 * spacingHeight;

            CustomTaskPane.Height = taskPaneHeight;

            list_botoes = list_botoes.OrderBy(b => b.Location.X).ToList(); // Ordenar os botões pela posição X

            int currentX = spacingWidth; // Começar o primeiro botão com o espaço inicial

            // Definir o tamanho máximo de fonte com base nos botões
            float maxFontSize = CalcularTamanhoFonteMaximo(list_botoes, buttonWidth, buttonHeight);

            foreach (Button botao in list_botoes)
            {
                // Definir a largura e altura do botão
                botao.Width = buttonWidth;
                botao.Height = buttonHeight;

                // Manter a coordenada Y fixa em 10, conforme o código original
                botao.Location = new System.Drawing.Point(currentX, spacingHeight);

                // Aplicar o tamanho de fonte ajustado para todos os botões
                botao.Font = new System.Drawing.Font(botao.Font.FontFamily, maxFontSize);

                // Atualizar a coordenada X para o próximo botão
                currentX += buttonWidth + spacingWidth;  // Atualizar a posição X para o próximo botão
            }

            
        }

        private float CalcularTamanhoFonteMaximo(List<Button> botoes, int largura, int altura)
        {
            float tamanhoMaximo = 14f; // Fonte máxima por padrão
            float tamanhoFonte = tamanhoMaximo + 0.1f; // Tamanho inicial de fonte
            float tamanhoMinimo = 5f; // Fonte máxima por padrão
            float margemDeFolga = 0.9f; // Margem de folga de 10%
            largura = (int)(largura * margemDeFolga); // Ajustar a largura com base na margem de folga

            Button botao_iteracao = new Button();


            foreach (var botao in botoes)
            {

                botao_iteracao.Font = new System.Drawing.Font(botao.Font.FontFamily, tamanhoFonte);
                // Medir o texto do botão e determinar o tamanho de fonte necessário para caber
                SizeF tamanhoTexto = botao.CreateGraphics().MeasureString(botao.Text, new System.Drawing.Font(botao.Font.FontFamily, tamanhoFonte));

                int num_linhas;

                if (!botao.Text.Contains(" "))
                {
                    num_linhas = 1;
                }
                else
                {
                    num_linhas = 2;
                }

                tamanhoFonte = tamanhoMaximo + 0.1f; // Restaurar o tamanho da fonte para o valor máximo

                // Calcular o tamanho máximo de fonte para o botão
                while (tamanhoTexto.Width > largura * num_linhas || botao_iteracao.Font.GetHeight() * num_linhas > altura)
                {
                    tamanhoFonte -= 0.1f; // Reduzir a fonte para caber

                    if (tamanhoFonte <= tamanhoMinimo) // Definir um limite para o tamanho da fonte
                    {
                        tamanhoFonte = 5;
                        break;
                    }

                    tamanhoTexto = botao.CreateGraphics().MeasureString(botao.Text, new System.Drawing.Font(botao.Font.FontFamily, tamanhoFonte));

                    // Definir o maior tamanho de fonte encontrado
                    if (tamanhoFonte < tamanhoMaximo)
                    {
                        tamanhoMaximo = tamanhoFonte;
                    }
                }
            }
                return tamanhoMaximo; // Retorna o maior tamanho de fonte que cabe em todos os botões
            
        }

        //private float CalcularTamanhoFonteMaximo(List<Button> botoes, int largura, int altura)
        //{
        //    // Definir o tamanho da fonte inicial para o cálculo
        //    float tamanhoFonteInicial = 12f; // Tamanho inicial de fonte
        //    float margemDeFolga = 1f; // Margem de folga de 5% (0.95 significa 95% da largura disponível)
        //    float tamanhoFonte = tamanhoFonteInicial;

        //    // Calcula a largura máxima que o texto pode ocupar
        //    int larguraDisponivel = (int)(largura * margemDeFolga);

        //    // Verificar se o texto pode caber em uma única linha
        //    foreach (var botao in botoes)
        //    {
        //        // Medir o tamanho do texto no botão
        //        SizeF tamanhoTexto = MeasuredTextSize(botao, botao.Text, botao.Font);

        //        // Ajustar o tamanho da fonte até que o texto caiba no botão (dentro da margem de 5%)
        //        while (tamanhoTexto.Width > larguraDisponivel || tamanhoTexto.Height > altura)
        //        {
        //            tamanhoFonte -= 0.5f; // Ajuste mais suave, diminuir de 0.5 em 0.5

        //            if (tamanhoFonte <= 5) // Definir um limite para o tamanho da fonte
        //            {
        //                tamanhoFonte = 5;
        //                break;
        //            }

        //            botao.Font = new System.Drawing.Font(botao.Font.FontFamily, tamanhoFonte);
        //            tamanhoTexto = MeasuredTextSize(botao, botao.Text, botao.Font);
        //        }

        //        // Definir o maior tamanho de fonte encontrado
        //        if (tamanhoFonte > 5)
        //        {
        //            tamanhoFonte = Math.Min(tamanhoFonte, 12f); // Não deixar ultrapassar o tamanho máximo de 12f
        //        }
        //    }

        //    return tamanhoFonte; // Retorna o tamanho da fonte ajustado
        //}

        //// Função para medir o tamanho do texto sem quebra de linha
        //private SizeF MeasuredTextSize(Button botao, string text, System.Drawing.Font font)
        //{
        //    using (Graphics g = botao.CreateGraphics())
        //    {
        //        // Usar a propriedade "NoWrap" para evitar a quebra de linha
        //        StringFormat sf = new StringFormat();
        //        sf.Trimming = StringTrimming.None;  // Garante que o texto não será cortado por causa de uma quebra de linha
        //        sf.FormatFlags = StringFormatFlags.NoWrap;  // Garantir que o texto não será quebrado em várias linhas

        //        return g.MeasureString(text, font, new PointF(0, 0), sf);
        //    }
        //}
        ////******* redimensionar texto de acordo com o tamanho dos botoes

        //Graphics g = botao.CreateGraphics();
        //System.Drawing.Font currentFont = botao.Font;
        //SizeF textSize = g.MeasureString(botao.Text, currentFont);

        //if (textSize.Width > buttonWidth && textSize.Height > buttonHeight)
        //{
        //    float scaleFactor = Math.Min(buttonWidth / textSize.Width, buttonHeight / textSize.Height);
        //    System.Windows.Forms.MessageBox.Show(scaleFactor.ToString());
        //    int newFontSize = (int)(currentFont.Size * scaleFactor * 0.8);
        //    botao.Font = new System.Drawing.Font(currentFont.FontFamily, newFontSize);
        //}

        //if (textSize.Width > buttonWidth || textSize.Height > buttonHeight)
        //{
        //    float scaleFactor = Math.Min(buttonWidth / textSize.Width, buttonHeight / textSize.Height);
        //    int newFontSize = (int)(currentFont.Size * scaleFactor);
        //    botao.Font = new System.Drawing.Font(currentFont.FontFamily, newFontSize);
        //}

        //if (textSize.Width > buttonWidth)
        //{
        //    float scaleFactor = buttonWidth / textSize.Width;
        //    System.Windows.Forms.MessageBox.Show(scaleFactor.ToString());
        //    int newFontSize = (int)(currentFont.Size * scaleFactor);
        //    botao.Font = new System.Drawing.Font(currentFont.FontFamily, newFontSize);
        //}
        //if (textSize.Height > buttonHeight)
        //{
        //    float scaleFactor = buttonHeight / textSize.Height;
        //    System.Windows.Forms.MessageBox.Show(scaleFactor.ToString());
        //    int newFontSize = (int)(currentFont.Size * scaleFactor);
        //    botao.Font = new System.Drawing.Font(currentFont.FontFamily, newFontSize);
        //}

        public void Visible(bool b)
        {
            foreach (Microsoft.Office.Tools.CustomTaskPane CTP in Class_New_or_Open_Event.Dicionario_Doc_e_TaskPane.Values) CTP.Visible = b;
        }

        public void MyCustomTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            Tarefa.Run(() =>
            {
                if (Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked) // Se o botão do Ribbon estiver marcado
                {
                    if (((Microsoft.Office.Tools.CustomTaskPane)sender).Visible == false) // Se o painel de estilos foi fechado
                    {
                        ((Microsoft.Office.Tools.CustomTaskPane)sender).Visible = true;
                    }
                }
            });
        }
    }
}
