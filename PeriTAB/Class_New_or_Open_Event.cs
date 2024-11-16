using iTextSharp.text.pdf.parser;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using Tarefa = System.Threading.Tasks.Task;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media.TextFormatting;
using static System.Net.WebRequestMethods;

namespace PeriTAB
{
    public class Class_New_or_Open_Event
    {
        Class_DocumentClose_Event iClass_DocumentClose_Event = new Class_DocumentClose_Event();

        //Class_AnyButtonClick_Event iClass_AnyButtonClick_Event;
        //public static MyUserControl iUserControl_1;
        public static Microsoft.Office.Tools.CustomTaskPane iTaskPane;
        //public static List<MyUserControl> list_UserControl = new List<MyUserControl>();

        //public static List<Microsoft.Office.Tools.CustomTaskPane> list_TaskPane = new List<Microsoft.Office.Tools.CustomTaskPane>();
        //public static List<Microsoft.Office.Interop.Word.Document> list_Doc = new List<Microsoft.Office.Interop.Word.Document>();



        public static Dictionary<Microsoft.Office.Interop.Word.Document, Microsoft.Office.Tools.CustomTaskPane> Dicionario_Doc_e_TaskPane = new Dictionary<Microsoft.Office.Interop.Word.Document, Microsoft.Office.Tools.CustomTaskPane>();


        //public class Variables
        //{
        //    private static List<UserControl1> var1 = new List<UserControl1>();
        //    private static List<Microsoft.Office.Tools.CustomTaskPane> var2 = new List<Microsoft.Office.Tools.CustomTaskPane>();
        //    public static List<UserControl1> list_UserControl1 { get { return var1; } set { } }
        //    public static List<Microsoft.Office.Tools.CustomTaskPane> list_TaskPane { get { return var2; } set { } }
        //}

        public void Evento_New_or_Open()
        {
            ((Microsoft.Office.Interop.Word.ApplicationEvents4_Event)Globals.ThisAddIn.Application).NewDocument += new ApplicationEvents4_NewDocumentEventHandler(Metodo_New_or_Open);
            Globals.ThisAddIn.Application.DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(Metodo_New_or_Open);
        }
        public void Metodo_New_or_Open(Microsoft.Office.Interop.Word.Document Doc) 
        {
            //System.Windows.Forms.MessageBox.Show("new or open");
            //Class_Buttons iClass_Buttons = new Class_Buttons();

            //Class_ChangeActiveDocumentNumber_Event iClass_ChangeActiveDocumentNumber_Event = new Class_ChangeActiveDocumentNumber_Event();
            iClass_DocumentClose_Event.Tracking_OpenDocumentNumber();

                //Class_AnyButtonClick_Event iClass_AnyButtonClick_Event = new Class_AnyButtonClick_Event();
                //iClass_AnyButtonClick_Event.Evento_AnyButtonClick(Globals.ThisAddIn.iMyUserControl);

                //if (Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked) Metodo_TaskPanes_Visible(true);

                //if (Globals.ThisAddIn.Dicionario_Doc_e_UserControl.ContainsKey(Doc)) return; //Se o documento já tem Taskpane, retorna.
                Class_DocChange_Event iClass_DocChange_Event = new Class_DocChange_Event(); iClass_DocChange_Event.Evento_DocChange();

            //Configura o Task Pane
            //List<UserControl1> list_UserControl1 = new List<UserControl1>();
            //MessageBox.Show(Globals.ThisAddIn.CustomTaskPanes.Count.ToString());
            //MessageBox.Show(Globals.ThisAddIn.Application.Documents.Count.ToString());

            //if (Globals.ThisAddIn.CustomTaskPanes.Count == 0 | Globals.ThisAddIn.Application.Documents.Count > Globals.ThisAddIn.CustomTaskPanes.Count)
            //{
            //MessageBox.Show("fefewfw fwefw");
            //if (Doc != null)
            //{
            //MessageBox.Show("fefewfw fwefw2222");

            if (!Globals.ThisAddIn.Dicionario_Doc_e_UserControl.ContainsKey(Doc))
            {

                Globals.ThisAddIn.iMyUserControl = new MyUserControl();
                Globals.ThisAddIn.iMyUserControl.AutoScroll = true;
                //System.Windows.Forms.MessageBox.Show(Control.DefaultFont.ToString());
                //Globals.ThisAddIn.iMyUserControl.Font = Control.DefaultFont;
                //Globals.ThisAddIn.iMyUserControl.Font = Control.;
                //Globals.ThisAddIn.iMyUserControl.AutoScroll = false;

                Class_AnyButtonClick_Event iClass_AnyButtonClick_Event = new Class_AnyButtonClick_Event();
                iClass_AnyButtonClick_Event.Evento_AnyButtonClick(Globals.ThisAddIn.iMyUserControl);

                //list_UserControl.Add(Globals.ThisAddIn.iMyUserControl);
                iTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(Globals.ThisAddIn.iMyUserControl, "Painel de Estilos (PeriTAB)");
                Globals.ThisAddIn.Dicionario_Doc_e_UserControl.Add(Doc, Globals.ThisAddIn.iMyUserControl);
                Dicionario_Doc_e_TaskPane.Add(Doc, iTaskPane);
                //MessageBox.Show("Taskpane adicionado");
                //MessageBox.Show(Globals.ThisAddIn.CustomTaskPanes.Count.ToString());
                //MessageBox.Show(Globals.ThisAddIn.Application.Documents.Count.ToString());
                //iTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
                //iTaskPane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
                //iTaskPane.Height = 90;
                iTaskPane.VisibleChanged += MyCustomTaskPane_VisibleChanged;
                RedimensionarTaskPane();
            }
            //iTaskPane.VisibleChanged += MyCustomTaskPane_VisibleChanged;

            //if (Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked) iTaskPane.Visible = true;

            //Revisa a habilitação do ToggleButton "Painel de Estilos" do Ribbon
            /*new Thread(() =>*/Tarefa.Run(() =>
            {
                int count = 0;
                while (!Globals.ThisAddIn.CustomTaskPanes[0].Visible)
                {
                    count++;
                    //Thread.Sleep(1);
                    if (Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked)
                    {
                        Class_New_or_Open_Event.Metodo_TaskPanes_Visible(true);
                    }
                    else { break; }
                }

                if (Ribbon1.Variables.debugging && count > 5)
                {
                    System.Windows.Forms.MessageBox.Show("While do Revisa a habilitação do ToggleButton \"Painel de Estilos\" do Ribbon rodou " + count.ToString() + " vezes");
                    return;
                }

            /*}).Start();*/});

            //list_TaskPane.Add(iTaskPane);
            //Dicionario_Doc_e_TaskPane.Add(Globals.ThisAddIn.Application.ActiveDocument, iTaskPane);

            //if (Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked) iTaskPane.Visible = true;
            //MessageBox.Show(Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked.ToString());
            //if (Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked) Class_New_or_Open_Event.Metodo_TaskPanes_Visible(true);
            //MessageBox.Show("taskpane added");
            //}

            //}

        }
        public void RedimensionarTaskPane()
        {
            iTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
            iTaskPane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;

            var screenArea = Screen.PrimaryScreen.WorkingArea;
            float dpiFactor = 96f / Graphics.FromHwnd(IntPtr.Zero).DpiX;

            //int spacingHeight = 5;
            //int buttonHeight = (int)(40 / dpiFactor);
            //int headerHeight = (int)(54 / dpiFactor);
            ////int headerHeight = (int)(38 / dpiFactor);
            //int taskPaneHeight = buttonHeight + headerHeight + 2*spacingHeight;

            //iTaskPane.Height = taskPaneHeight;
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
            //int headerHeight = (int)(38 / dpiFactor);
            int taskPaneHeight = buttonHeight + headerHeight + 2 * spacingHeight;

            iTaskPane.Height = taskPaneHeight;

            List<Button> list_botoes = userControl.Controls.OfType<Button>().ToList(); // Obter todos os botões de userControl

            list_botoes = list_botoes.OrderBy(b => b.Location.X).ToList(); // Ordenar os botões pela posição X

            int currentX = spacingWidth; // Começar o primeiro botão com o espaço inicial

            foreach (Button botao in list_botoes)
            {
                // Definir a largura e altura do botão
                botao.Width = buttonWidth;
                //System.Windows.Forms.MessageBox.Show(buttonWidth.ToString());
                botao.Height = buttonHeight;
                botao.Height = (int)(buttonWidth / 2);

                // Manter a coordenada Y fixa em 10, conforme o código original
                botao.Location = new System.Drawing.Point(currentX, spacingHeight);

                // Atualizar a coordenada X para o próximo botão
                currentX += buttonWidth + spacingWidth;  // Atualizar a posição X para o próximo botão
            }
        }
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

                //while (true)
                //{
                //    SizeF textSize = g.MeasureString(botao.Text, currentFont);
                //    if (textSize.Width > buttonWidth)
                //    {
                //        float scaleFactor = (float)0.9;
                //        int newFontSize = (int)(currentFont.Size * scaleFactor);
                //        botao.Font = new System.Drawing.Font(currentFont.FontFamily, newFontSize);
                //    }
                //    else
                //    {
                //        break;
                //    }
                //    currentFont = botao.Font;
                //}

                //    // Se o texto for maior que a largura do botão, ajusta o tamanho da fonte
                //    if (textSize.Width > buttonWidth)
                //{

                //    // Calcula o fator de ajuste baseado na largura
                //    float scaleFactor = buttonWidth / textSize.Width;

                //    // Ajusta o tamanho da fonte de acordo com a largura
                //    float newFontSize = currentFont.Size * scaleFactor;

                //    // Ajusta a fonte até que o texto caiba também na altura do botão
                //    // Verifique se o texto em várias linhas se ajusta ao botão verticalmente
                //    while (true)
                //    {
                //        // Cria uma nova fonte com o tamanho ajustado
                //        System.Drawing.Font adjustedFont = new System.Drawing.Font(currentFont.FontFamily, newFontSize);

                //        // Medimos o tamanho do texto com a fonte ajustada
                //        SizeF adjustedTextSize = g.MeasureString(botao.Text, adjustedFont);

                //        // Verifica se o texto ajustado cabe na largura e na altura do botão
                //        if (adjustedTextSize.Width <= buttonWidth && adjustedTextSize.Height <= buttonHeight)
                //        {
                //            // Se couber, definimos a nova fonte
                //            botao.Font = adjustedFont;
                //            break;
                //        }
                //        else
                //        {
                //            // Caso não caiba, diminui o tamanho da fonte e tenta novamente
                //            newFontSize -= 1;
                //        }


                        // Clean up the Graphics object.
                        //g.Dispose();

                //    }
                //}
            //}

            //    foreach (Control control in userControl.Controls)
            //{
            //    if (control is Button button)
            //    {
            //        // Definir a largura e altura do botão
            //        button.Width = buttonWidth;
            //        button.Height = buttonHeight;

            //        // Manter a coordenada Y fixa em 10, conforme o código original
            //        button.Location = new System.Drawing.Point(currentX, spacingHeight);

            //        // Atualizar a coordenada X para o próximo botão
            //        currentX += buttonWidth + spacingWidth;  // Atualizar a posição X para o próximo botão
            //    }
            //}

            //// Garantir que o último botão ocupe o espaço total restante sem espaçamento extra
            //Control lastButton = userControl.Controls.OfType<Button>().LastOrDefault();
            //if (lastButton != null)
            //{
            //    // Ajustar a posição do último botão para não ter espaçamento à direita
            //    lastButton.Location = new System.Drawing.Point(currentX - (buttonWidth + spacingWidth), spacingHeight);
            //}
        
        

        //public static void Metodo_TaskPanes_Visible(bool b)
        //{
        //    foreach (Microsoft.Office.Tools.CustomTaskPane CTP in list_TaskPane)
        //    {
        //        try { CTP.Visible = b; } catch (System.Runtime.InteropServices.COMException ex) { }
        //    }
        //    //TaskPane2.Visible = b;
        //}

        public static void Metodo_TaskPanes_Visible(bool b)
        {
            /*new Thread(() =>*/Tarefa.Run(() =>
            {
                foreach (Microsoft.Office.Tools.CustomTaskPane CTP in Class_New_or_Open_Event.Dicionario_Doc_e_TaskPane.Values) CTP.Visible = b;
            /*}).Start();*/});


            //foreach (Microsoft.Office.Tools.CustomTaskPane CTP in Class_New_or_Open_Event.Dicionario_Doc_e_TaskPane.Values)
            //{
            //    try { CTP.Visible = b; } catch (System.Runtime.InteropServices.COMException ex) { }
            //}

            //TaskPane2.Visible = b;
        }


        private void MyCustomTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch(); if (Ribbon1.Variables.debugging) { stopwatch.Start(); }
            //bool b = false;
            //bool checked1 = Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos2.Checked;
            //var botao_painel_de_estilos2 = (Microsoft.Office.Tools.CustomTaskPane)sender;
            //bool a = botao_painel_de_estilos2.Visible;

            bool Visib = ((Microsoft.Office.Tools.CustomTaskPane)sender).Visible;
            bool TB_checked = Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked;

            //MessageBox.Show("Visib = " + Visib.ToString() + " e TB_checked = " + Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked.ToString());
            //if (Visib == false & Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked == true) 
            //{
            //    Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked = false;
            //    //foreach (Microsoft.Office.Tools.CustomTaskPane CTP in list_TaskPane)
            //    foreach (Microsoft.Office.Tools.CustomTaskPane CTP in Class_New_or_Open_Event.Dicionario_Doc_e_TaskPane.Values)
            //    {
            //        try { CTP.Visible = false; } catch (System.Runtime.InteropServices.COMException ex) { }
            //    }
            //}
            if (Visib != TB_checked)
            {
                //new Thread(() =>
                //{
                //    Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked = Visib;
                //    Metodo_TaskPanes_Visible(Visib);
                //}).Start();
                Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked = Visib;
                Metodo_TaskPanes_Visible(Visib);
                if (Ribbon1.Variables.debugging)
                {
                    string msg_StatusBar = "";
                    //if (!Visib) msg_StatusBar = "Painel de Estilos: Fechado pelo X";
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
                    //foreach (Microsoft.Office.Tools.CustomTaskPane CTP in list_TaskPane)

                    //foreach (Microsoft.Office.Tools.CustomTaskPane CTP in Class_New_or_Open_Event.Dicionario_Doc_e_TaskPane.Values)
                    //{
                    //    try { CTP.Visible = TB_checked; } catch (System.Runtime.InteropServices.COMException ex) { }
                    //}
                }


            //if (botao_painel_de_estilos2.Visible)
            ////{
            //MessageBox.Show(Visib.ToString());
            //foreach (Microsoft.Office.Tools.CustomTaskPane CTP in list_TaskPane)
            //    {
            //        //CTP.Visible = a;
            //    try { CTP.Visible = Visib; Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos2.Checked = Visib; } catch (System.Runtime.InteropServices.COMException ex) { }
            //}
            //}

            //if (checked1)
            //{
            //    foreach (Microsoft.Office.Tools.CustomTaskPane CTP in list_TaskPane)
            //    {
            //        CTP.Visible = true;
            //        //try { CTP.Visible = true; } catch { }
            //        //try { CTP.Visible = checked1; Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos2.Checked = checked1; } catch (System.Runtime.InteropServices.COMException ex) { }
            //    }
            //}
            //else
            //{
            //    foreach (Microsoft.Office.Tools.CustomTaskPane CTP in list_TaskPane)
            //    {
            //        CTP.Visible = false;
            //        //try { CTP.Visible = false; } catch { }
            //    }
            //}


            //foreach (Microsoft.Office.Tools.CustomTaskPane CTP in list_TaskPane)
            //{
            //if (botao_painel_de_estilos2.Visible != Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos2.Checked)
            //{
            //    b = true;
            //}
            //}

            //foreach (Microsoft.Office.Tools.CustomTaskPane CTP in list_TaskPane)
            //{
            //    //try {
            //        if (CTP.Visible == false & Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos2.Checked == true)
            //        { 
            //            b = true;
            //        }                    
            //    //} catch (System.Runtime.InteropServices.COMException ex) { }
            //}
            //if (b)
            //{
            //    foreach (Microsoft.Office.Tools.CustomTaskPane CTP in list_TaskPane)
            //    {
            //        try { CTP.Visible = checked1; Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos2.Checked = checked1; } catch (System.Runtime.InteropServices.COMException ex) { }
            //    }
            //}
            //if (Globals.ThisAddIn.TaskPane2.Visible == false & Globals.Ribbons.Ribbon1.toggleButton_estilos.Checked == true) { Globals.Ribbons.Ribbon1.toggleButton_estilos.Checked = false; }
        }
    }
}
