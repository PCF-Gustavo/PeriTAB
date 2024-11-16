using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using Tarefa = System.Threading.Tasks.Task;
using System.Windows.Forms;

namespace PeriTAB
{
    public class Class_New_or_Open_Event
    {
        Class_DocumentClose_Event iClass_DocumentClose_Event = new Class_DocumentClose_Event();

        public static Microsoft.Office.Tools.CustomTaskPane iTaskPane;
        public static Dictionary<Microsoft.Office.Interop.Word.Document, Microsoft.Office.Tools.CustomTaskPane> Dicionario_Doc_e_TaskPane = new Dictionary<Microsoft.Office.Interop.Word.Document, Microsoft.Office.Tools.CustomTaskPane>();

        public void Evento_New_or_Open()
        {
            ((Microsoft.Office.Interop.Word.ApplicationEvents4_Event)Globals.ThisAddIn.Application).NewDocument += new ApplicationEvents4_NewDocumentEventHandler(Metodo_New_or_Open);
            Globals.ThisAddIn.Application.DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(Metodo_New_or_Open);
        }
        public void Metodo_New_or_Open(Microsoft.Office.Interop.Word.Document Doc) 
        {
            //System.Windows.Forms.MessageBox.Show("new or open");
            iClass_DocumentClose_Event.Tracking_OpenDocumentNumber();
            
            if (!Globals.ThisAddIn.Dicionario_Doc_e_UserControl.ContainsKey(Doc))
            {

                Globals.ThisAddIn.iMyUserControl = new MyUserControl();
                Globals.ThisAddIn.iMyUserControl.AutoScroll = true;

                Class_AnyButtonClick_Event iClass_AnyButtonClick_Event = new Class_AnyButtonClick_Event();
                iClass_AnyButtonClick_Event.Evento_AnyButtonClick(Globals.ThisAddIn.iMyUserControl);

                iTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(Globals.ThisAddIn.iMyUserControl, "Painel de Estilos (PeriTAB)");
                Globals.ThisAddIn.Dicionario_Doc_e_UserControl.Add(Doc, Globals.ThisAddIn.iMyUserControl);
                Dicionario_Doc_e_TaskPane.Add(Doc, iTaskPane);
                iTaskPane.VisibleChanged += MyCustomTaskPane_VisibleChanged;
                RedimensionarTaskPane();
            }

            //Checa se deve mostrar o "Painel de Estilos" do Ribbon
            if (Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked)
            {
                Class_New_or_Open_Event.Metodo_TaskPanes_Visible(true);
            }

        }
        public void RedimensionarTaskPane()
        {
            iTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
            iTaskPane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;

            var screenArea = Screen.PrimaryScreen.WorkingArea;
            float dpiFactor = 96f / Graphics.FromHwnd(IntPtr.Zero).DpiX;

            int taskPaneWidth = screenArea.Width;

            var userControl = Globals.ThisAddIn.iMyUserControl;

            int buttonCount = userControl.Controls.OfType<Button>().Count(); // Contar todos os botões de userControl

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

            iTaskPane.Height = taskPaneHeight;

            List<Button> list_botoes = userControl.Controls.OfType<Button>().ToList(); // Obter todos os botões de userControl

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



        public static void Metodo_TaskPanes_Visible(bool b)
        {
            /*new Thread(() =>*/Tarefa.Run(() =>
            {
                foreach (Microsoft.Office.Tools.CustomTaskPane CTP in Class_New_or_Open_Event.Dicionario_Doc_e_TaskPane.Values) CTP.Visible = b;
            /*}).Start();*/});
        }


        private void MyCustomTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon1.Variables.debugging) { stopwatch.Start(); }

            bool Visib = ((Microsoft.Office.Tools.CustomTaskPane)sender).Visible;
            bool TB_checked = Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked;

            if (Visib != TB_checked)
            {
                Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked = Visib;
                Metodo_TaskPanes_Visible(Visib);
                if (Ribbon1.Variables.debugging)
                {
                    string msg_StatusBar = "";
                    switch (Visib)
                    {
                        case true:
                            msg_StatusBar = "Painel de Estilos: Aberto ???";
                            break;
                        case false:
                            msg_StatusBar = "Painel de Estilos: Fechado pelo X";
                            break;
                    }
                    stopwatch.Stop();
                    msg_StatusBar += $" (Tempo de execução: {stopwatch.Elapsed.TotalSeconds:F2} segundos)";
                    Globals.ThisAddIn.Application.StatusBar = msg_StatusBar;
                }
            }
        }
    }
}
