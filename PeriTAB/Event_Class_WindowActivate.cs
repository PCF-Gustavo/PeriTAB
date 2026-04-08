using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;

using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace PeriTAB
{
    internal class Class_WindowActivate_Event
    {
        private readonly Class_ContentControlOnExit_Event iClass_ContentControlOnExit_Event = new Class_ContentControlOnExit_Event();
        private readonly Class_CustomTaskPanes iClass_CustomTaskPanes = new Class_CustomTaskPanes();
        private readonly Class_AnyButtonClick_Event iClass_AnyButtonClick_Event = new Class_AnyButtonClick_Event();

        //private MyUserControl UserControl;
        //private CustomTaskPane TaskPane;

        private readonly Class_RibbonControls iClass_RibbonControls = new Class_RibbonControls();

        public void Evento_WindowActivate()
        {
            Globals.ThisAddIn.Application.WindowActivate += new ApplicationEvents4_WindowActivateEventHandler(Metodo_WindowActivate);
        }
        public void Metodo_WindowActivate(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn)
        {
            //MessageBox.Show("Winact");

            //if (!Globals.ThisAddIn.Dicionario_Window_e_UserControl.ContainsKey(Wn))
            if (!Window_Possui_TaskPane(Wn))
            {
                //if (!Globals.ThisAddIn.Dicionario_Window_e_UserControl.Values.Any(uc => uc.Document == Doc)) { 
                //    iClass_ContentControlOnExit_Event.Metodo_ContentControlOnExit();
                //    MessageBox.Show("vinculou os ContentControl");
                //}

                //Globals.ThisAddIn.Dicionario_Window_e_Doc.Add(Wn, Doc);

                MyUserControl UserControl = Add_UserControl(Wn);

                Globals.ThisAddIn.Dicionario_Window_e_UserControl.Add(Wn, UserControl);
                //Globals.ThisAddIn.Dicionario_Window_e_TaskPane.Add(Wn, iTaskPane);
            }

            iClass_RibbonControls.Atualiza_Habilitacao(Globals.Ribbons.Ribbon.CheckBox_destaca_campos);
            iClass_RibbonControls.Atualiza_Habilitacao(Globals.Ribbons.Ribbon.CheckBox_mostra_indicadores);
            iClass_RibbonControls.Atualiza_Habilitacao(Globals.Ribbons.Ribbon.Button_renomeia_documento);
            iClass_RibbonControls.Atualiza_Habilitacao(Globals.Ribbons.Ribbon.Button_gera_pdf);

            //try
            //{
            //    //Revisa a habilitação do CheckBox "Destacar campos" do Ribbon
            //    if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)1) { Globals.Ribbons.Ribbon.CheckBox_destaca_campos.Checked = true; }
            //    if (Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)0 | Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading == (WdFieldShading)2) { Globals.Ribbons.Ribbon.CheckBox_destaca_campos.Checked = false; }

            //    //Revisa a habilitação do CheckBox "Mostrar indicadores" do Ribbon
            //    if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowBookmarks == true) { Globals.Ribbons.Ribbon.CheckBox_mostra_indicadores.Checked = true; }
            //    if (Globals.ThisAddIn.Application.ActiveWindow.View.ShowBookmarks == false) { Globals.Ribbons.Ribbon.CheckBox_mostra_indicadores.Checked = false; }

            //    //Revisa a habilitação do botao "Renomeia Documento" do Ribbon
            //    iClass_RibbonControls.Button_renomeia_documento_valorinicial();
            //    //iClass_Buttons.button_renomeia_documento_Default();
            //    if (Globals.ThisAddIn.Application.ActiveDocument.Path == "") { Globals.Ribbons.Ribbon.Button_renomeia_documento.Enabled = false; Globals.Ribbons.Ribbon.Button_renomeia_documento.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon.Button_renomeia_documento.SuperTip = "Este documento ainda não foi salvo."; }

            //    //Revisa a habilitação do botao "Gera PDF" do Ribbon
            //    iClass_RibbonControls.Button_gera_pdf_valorinicial();
            //    //iClass_Buttons.button_gera_pdf_Default();
            //    if (Globals.ThisAddIn.Application.ActiveDocument.Path == "") { Globals.Ribbons.Ribbon.Button_gera_pdf.Enabled = false; Globals.Ribbons.Ribbon.Button_gera_pdf.ScreenTip = "Desabilitado"; Globals.Ribbons.Ribbon.Button_gera_pdf.SuperTip = "Este documento ainda não foi salvo."; }

            //}
            //catch (System.Runtime.InteropServices.COMException) { }


        }
        private MyUserControl Add_UserControl(Window Wn)
        {
            MyUserControl UserControl = new MyUserControl { AutoScroll = true };

            iClass_AnyButtonClick_Event.Evento_AnyButtonClick(UserControl);

            UserControl.TaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(UserControl, "Painel de Estilos (PeriTAB)", Wn);
            UserControl.TaskPane.VisibleChanged += iClass_CustomTaskPanes.MyCustomTaskPane_VisibleChanged;
            iClass_CustomTaskPanes.Redimensionar(UserControl);
            if (Globals.Ribbons.Ribbon.ToggleButton_painel_de_estilos.Checked) UserControl.TaskPane.Visible = true; //Checa se deve mostrar o "Painel de Estilos" do Ribbon

            return UserControl;
        }

        private bool Window_Possui_TaskPane(Window window)
        {
            foreach (CustomTaskPane pane in Globals.ThisAddIn.CustomTaskPanes)
            {
                if (pane.Window == window && pane.Title == "Painel de Estilos (PeriTAB)")
                {
                    return true;
                }
            }
            return false;
        }

    }
}
