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

            List<Button> list_botoes = MyUserControl.Controls.OfType<Button>().ToList(); // Obter todos os botões de userControl
            int buttonCount = list_botoes.Count(); // Contar todos os botões de userControl

            // Calcular a largura total disponível para os botões
            int spacingWidth = 5;
            int totalSpacingWidth = spacingWidth * (buttonCount + 1);  // Total de espaço entre os botões e as bordas
            int totalButtonWidth = taskPaneWidth - totalSpacingWidth;  // Largura total disponível para os botões

            // A largura de cada botão será a largura total disponível dividida pelo número de botões
            int buttonWidth = totalButtonWidth / buttonCount;

            int buttonHeight = (int)(buttonWidth / 3);
            int spacingHeight = (int)(buttonHeight / 10);
            int headerHeight = (int)(80 / dpiFactor);
            int taskPaneHeight = buttonHeight + headerHeight + 2 * spacingHeight;

            CustomTaskPane.Height = taskPaneHeight;

            list_botoes = list_botoes.OrderBy(b => b.Location.X).ToList(); // Ordenar os botões pela posição X

            int currentX = spacingWidth; // Começar o primeiro botão com o espaço inicial

            foreach (Button botao in list_botoes)
            {
                // Definir a largura e altura do botão
                botao.Width = buttonWidth;
                botao.Height = buttonHeight;
                botao.Height = (int)(buttonWidth / 2);

                // Manter a coordenada Y fixa em 10, conforme o código original
                botao.Location = new System.Drawing.Point(currentX, spacingHeight);

                // Atualizar a coordenada X para o próximo botão
                currentX += buttonWidth + spacingWidth;  // Atualizar a posição X para o próximo botão
            }
        }

        //******* redimensionar texto de acordo com o tamanho dos botoes

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
