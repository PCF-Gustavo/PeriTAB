using iTextSharp.text.pdf.parser;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using static System.Net.WebRequestMethods;

namespace PeriTAB
{
    public class Class_New_or_Open_Event
    {

        //Class_AnyButtonClick_Event iClass_AnyButtonClick_Event;
        //public static MyUserControl iUserControl_1;
        public static Microsoft.Office.Tools.CustomTaskPane TaskPane_1;
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
            if (Globals.ThisAddIn.Dicionario_Doc_e_UserControl.ContainsKey(Doc)) return;
            //MessageBox.Show("new or open");
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
                
            Globals.ThisAddIn.iMyUserControl = new MyUserControl();
            
            //list_UserControl.Add(Globals.ThisAddIn.iMyUserControl);
            TaskPane_1 = Globals.ThisAddIn.CustomTaskPanes.Add(Globals.ThisAddIn.iMyUserControl, "Painel de Estilos (PeriTAB)");
            Globals.ThisAddIn.Dicionario_Doc_e_UserControl.Add(Doc, Globals.ThisAddIn.iMyUserControl);
            Dicionario_Doc_e_TaskPane.Add(Doc, TaskPane_1);
            //MessageBox.Show(Globals.ThisAddIn.CustomTaskPanes.Count.ToString());
            //MessageBox.Show(Globals.ThisAddIn.Application.Documents.Count.ToString());
            TaskPane_1.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
            TaskPane_1.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
            TaskPane_1.Height = 90;
            TaskPane_1.VisibleChanged += MyCustomTaskPane_VisibleChanged;

            Class_AnyButtonClick_Event iClass_AnyButtonClick_Event = new Class_AnyButtonClick_Event();
            iClass_AnyButtonClick_Event.Evento_AnyButtonClick(Globals.ThisAddIn.iMyUserControl);

            //list_TaskPane.Add(TaskPane_1);
            //Dicionario_Doc_e_TaskPane.Add(Globals.ThisAddIn.Application.ActiveDocument, TaskPane_1);

            //if (Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked) TaskPane_1.Visible = true;
            //MessageBox.Show(Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked.ToString());
            //if (Globals.Ribbons.Ribbon1.toggleButton_painel_de_estilos.Checked) Class_New_or_Open_Event.Metodo_TaskPanes_Visible(true);
            //MessageBox.Show("taskpane added");
            //}

            //}

        }

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
            new Thread(() =>
            {
                foreach (Microsoft.Office.Tools.CustomTaskPane CTP in Class_New_or_Open_Event.Dicionario_Doc_e_TaskPane.Values) CTP.Visible = b;
            }).Start();


            //foreach (Microsoft.Office.Tools.CustomTaskPane CTP in Class_New_or_Open_Event.Dicionario_Doc_e_TaskPane.Values)
            //{
            //    try { CTP.Visible = b; } catch (System.Runtime.InteropServices.COMException ex) { }
            //}

            //TaskPane2.Visible = b;
        }

        

        private void MyCustomTaskPane_VisibleChanged(object sender, EventArgs e)
        {
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
