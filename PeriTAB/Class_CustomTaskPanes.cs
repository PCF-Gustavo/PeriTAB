using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Windows.Controls;
using System.Windows.Forms;
using Button = System.Windows.Forms.Button;
using CustomTaskPane = Microsoft.Office.Tools.CustomTaskPane;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Task = System.Threading.Tasks.Task;

internal static class NativeMethods
{
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
}

namespace PeriTAB
{
    public class Class_CustomTaskPanes
    {
        //private readonly Class_AnyButtonClick_Event iClass_AnyButtonClick_Event = new Class_AnyButtonClick_Event();

        //private MyUserControl UserControl;
        //private readonly CustomTaskPane TaskPane;

        //public static Microsoft.Office.Tools.CustomTaskPane iTaskPane;

        public void Redimensionar(MyUserControl MyUserControl)
        {
            MyUserControl.TaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
            MyUserControl.TaskPane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;

            IntPtr hwndWord = NativeMethods.GetForegroundWindow();
            Screen screenDoWord = Screen.FromHandle(hwndWord);
            var screenArea = screenDoWord.WorkingArea;

            float dpiFactor;
            using (Graphics g = Graphics.FromHwnd(hwndWord))
            {
                dpiFactor = 96f / g.DpiX;
            }

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

            MyUserControl.TaskPane.Height = taskPaneHeight;

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
                //System.Drawing.Font fonteAntiga = botao.Font;
                botao.Font = new System.Drawing.Font(botao.Font.FontFamily, maxFontSize);
                //fonteAntiga.Dispose();

                // Atualizar a coordenada X para o próximo botão
                currentX += buttonWidth + spacingWidth;  // Atualizar a posição X para o próximo botão
            }
        }

        private float CalcularTamanhoFonteMaximo(List<Button> botoes, int largura, int altura)
        {
            float tamanhoMaximo = 15f; // Fonte máxima por padrão
            float tamanhoMinimo = 5f; // Fonte máxima por padrão
            float margemDeFolga_altura = 0.85f; // Margem de folga da altura de 15%
            altura = (int)(altura * margemDeFolga_altura); // Ajustar a largura com base na margem de folga
            int NumeroDeLinhasMaximo = 2;
            float margemDeFolga_largura = 0.9f; // Margem de folga da largura de 10%
            largura = (int)(largura * margemDeFolga_largura); // Ajustar a largura com base na margem de folga

            // Medir a altura do texto do botão e determinar o maior tamanho de fonte que cabe em relação a altura
            float tamanhoFonte = tamanhoMaximo + 0.1f; // Tamanho inicial de fonte
            Button button_teste = new Button();
            button_teste.Font = new System.Drawing.Font(button_teste.Font.FontFamily, tamanhoFonte);

            while (button_teste.Height + (NumeroDeLinhasMaximo - 1) * button_teste.Font.GetHeight() > altura)
            {
                tamanhoFonte -= 0.1f; // Reduzir a fonte para caber

                if (tamanhoFonte <= tamanhoMinimo) // Definir um limite para o tamanho da fonte
                {
                    return tamanhoMinimo;
                }

                button_teste.Font = new System.Drawing.Font(button_teste.Font.FontFamily, tamanhoFonte);

            }
            tamanhoMaximo = tamanhoFonte;

            // Medir a largura de cada palavr' do texto dos botões e determinar o maior tamanho de fonte que cabe em relação a largura
            tamanhoFonte = tamanhoMaximo + 0.1f; // Tamanho inicial de fonte
            foreach (var botao in botoes)
            {
                using (Graphics Graphics = botao.CreateGraphics())
                {
                    // Loop em cada palavra do texto do botao
                    foreach (string palavra in botao.Text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        SizeF tamanhoTexto = Graphics.MeasureString(palavra, new System.Drawing.Font(botao.Font.FontFamily, tamanhoFonte));

                        tamanhoFonte = tamanhoMaximo + 0.1f; // Restaurar o tamanho da fonte para o valor máximo

                        while (tamanhoTexto.Width > largura)
                        {
                            tamanhoFonte -= 0.1f; // Reduzir a fonte para caber

                            if (tamanhoFonte <= tamanhoMinimo) // Definir um limite para o tamanho da fonte
                            {
                                return tamanhoMinimo;
                            }

                            tamanhoTexto = Graphics.MeasureString(palavra, new System.Drawing.Font(botao.Font.FontFamily, tamanhoFonte));

                            // Definir o maior tamanho de fonte encontrado
                            if (tamanhoFonte < tamanhoMaximo)
                            {
                                tamanhoMaximo = tamanhoFonte;
                            }
                        }
                    }
                }
            }
            return tamanhoMaximo; // Retorna o maior tamanho de fonte que cabe em todos os botões
        }

        //public void Visible(bool b)
        //{
        //    foreach (Microsoft.Office.Tools.CustomTaskPane CTP in Class_New_or_Open_Event.Dicionario_Doc_e_TaskPane.Values) CTP.Visible = b;
        //}
        public void Visible(bool b)
        {
            foreach (MyUserControl uc in Globals.ThisAddIn.Dicionario_Window_e_UserControl.Values) { uc.TaskPane.Visible = b; }

            //foreach (Microsoft.Office.Tools.CustomTaskPane CTP in Globals.ThisAddIn.Dicionario_Window_e_TaskPane.Values) CTP.Visible = b;
        }

        public void MyCustomTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            Task.Run(() =>
            {
                if (Globals.Ribbons.Ribbon.ToggleButton_painel_de_estilos.Checked) // Se o botão do Ribbon estiver marcado
                {
                    if (((Microsoft.Office.Tools.CustomTaskPane)sender).Visible == false) // Se o painel de estilos foi fechado
                    {
                        ((Microsoft.Office.Tools.CustomTaskPane)sender).Visible = true;
                    }
                }
            });
        }

        //public (MyUserControl,CustomTaskPane) AddUserControl_and_TaskPane(Microsoft.Office.Interop.Word.Document Doc)
        //{
        //    UserControl = new MyUserControl();
        //    UserControl.AutoScroll = true;

        //    //Globals.ThisAddIn.iMyUserControl = new MyUserControl();
        //    //Globals.ThisAddIn.iMyUserControl.AutoScroll = true;

        //    iClass_AnyButtonClick_Event.Evento_AnyButtonClick(UserControl);

        //    iTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(UserControl, "Painel de Estilos (PeriTAB)");
        //    iTaskPane.VisibleChanged += MyCustomTaskPane_VisibleChanged;
        //    Redimensionar(UserControl, iTaskPane);
        //    if (Globals.Ribbons.Ribbon.toggleButton_painel_de_estilos.Checked) iTaskPane.Visible = true; //Checa se deve mostrar o "Painel de Estilos" do Ribbon

        //    return (UserControl, iTaskPane);
        //}

        //public (MyUserControl, CustomTaskPane) AddUserControl_and_TaskPane(Window Wn)
        //{
        //    UserControl = new MyUserControl { AutoScroll = true };

        //    //Globals.ThisAddIn.iMyUserControl = new MyUserControl();
        //    //Globals.ThisAddIn.iMyUserControl.AutoScroll = true;

        //    iClass_AnyButtonClick_Event.Evento_AnyButtonClick(UserControl);

        //    iTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(UserControl, "Painel de Estilos (PeriTAB)", Wn);
        //    iTaskPane.VisibleChanged += MyCustomTaskPane_VisibleChanged;
        //    Redimensionar(UserControl, iTaskPane);
        //    if (Globals.Ribbons.Ribbon.toggleButton_painel_de_estilos.Checked) iTaskPane.Visible = true; //Checa se deve mostrar o "Painel de Estilos" do Ribbon

        //    return (UserControl, iTaskPane);
        //}

        //public static void Atualiza_Destaques(MyUserControl UserControl)
        //{
        //    UserControl.Remove_Destaque_Botoes();

        //    if (Globals.ThisAddIn.Application.Selection.Tables.Count == 0) // Inseri pq selecionar paragrafos com tabela causa problemas de seleção.
        //    {
        //        List<Paragraph> paragrafosSelecionados = Globals.ThisAddIn.Application.Selection.Paragraphs.Cast<Paragraph>().ToList();

        //        foreach (Paragraph p in paragrafosSelecionados)
        //        {
        //            Style estilo = null;
        //            if (p.Range.StoryType == WdStoryType.wdMainTextStory)
        //            {
        //                try { estilo = p.Range.get_Style(); } catch (System.Runtime.InteropServices.COMException) { }

        //                if (estilo != null && UserControl.Dicionario_Estilo_e_Botao.ContainsKey(estilo.NameLocal))
        //                {
        //                    System.Windows.Forms.Button botao = UserControl.Dicionario_Estilo_e_Botao[estilo.NameLocal];
        //                    UserControl.Habilita_Destaca(botao, true, true);
        //                }
        //            }
        //            if (p.Range.StoryType == WdStoryType.wdFootnotesStory)
        //            {
        //                Range Selecao_inicial = Globals.ThisAddIn.Application.Selection.Range; //Salva a seleção inicial (Inseri pq estilo = p.Range.ParagraphFormat.get_Style(); estava modificando implicitamente a selação)
        //                try { estilo = p.Range.ParagraphFormat.get_Style(); } catch (System.Runtime.InteropServices.COMException) { }
        //                Selecao_inicial.Select(); // Restaura a seleção inicial
        //                if (estilo != null && UserControl.Dicionario_Estilo_e_Botao.ContainsKey(estilo.NameLocal))
        //                {
        //                    System.Windows.Forms.Button botao = UserControl.Dicionario_Estilo_e_Botao[estilo.NameLocal];
        //                    UserControl.Habilita_Destaca(botao, true, true);
        //                }
        //            }
        //        }
        //    }
        //}
    }
}
