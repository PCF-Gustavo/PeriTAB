using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace PeriTAB
{
    public class Class1
    {
        public void Metodo_DocumentOpen(Microsoft.Office.Interop.Word.Document Doc)
        {
            MessageBox.Show("DocumentOpen");
        }

        public void Metodo_NewDocument(Microsoft.Office.Interop.Word.Document Doc)
        {
            MessageBox.Show("NewDocument");
        }

        public void Metodo_DocumentBeforeClose(Microsoft.Office.Interop.Word.Document Doc, ref bool Cancel)
        {
            MessageBox.Show("DocumentBeforeClose");
        }

        public void Metodo_DocumentChange()
        {
            MessageBox.Show("DocumentChange");
        }

        public void Metodo_DocumentSync(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Core.MsoSyncEventType SyncEventType)
        {
            MessageBox.Show("DocumentSync");
        }

        public void Metodo_ProtectedViewWindowActivate(Microsoft.Office.Interop.Word.ProtectedViewWindow PvWindow)
        {
            MessageBox.Show("DocumentSync");
        }

        public void Metodo_WindowActivate(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn)
        {
            MessageBox.Show("WindowActivate");
        }

        public void Metodo_WindowBeforeDoubleClick(Microsoft.Office.Interop.Word.Selection Sel, ref bool Cancel)
        {
            MessageBox.Show("WindowBeforeDoubleClick");
        }
        public void Metodo_WindowBeforeRightClick(Microsoft.Office.Interop.Word.Selection Sel, ref bool Cancel)
        {
            MessageBox.Show("WindowBeforeRightClick");
        }
        public void Metodo_WindowDeactivate(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn)
        {
            MessageBox.Show("WindowDeactivate");
        }

        public void Metodo_WindowSelectionChange(Microsoft.Office.Interop.Word.Selection Sel)
        {
            MessageBox.Show("WindowSelectionChange");
        }

        public void Metodo_WindowSize(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn)
        {
            MessageBox.Show("WindowSize");

        }
        public void SettingChanging(object sender, System.Configuration.SettingChangingEventArgs e)
        {
            MessageBox.Show("SettingChanging");
        }
        public void Metodo_TabDisposed(object sender, System.EventArgs e)
        {
            MessageBox.Show("TabDisposed");
        }
        public void Metodo_RibbonClose(object sender, System.EventArgs e)
        {
            MessageBox.Show("RibbonClose");
        }
        public void Metodo_menu1ItemsLoading(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("menu1ItemsLoading");
        }
        private void Metodo_ContentControlOnEnter(Microsoft.Office.Interop.Word.ContentControl ContentControl)
        {
            MessageBox.Show("ContentControlOnEnter");
        }

        //internal static void Metodo_DocumentOpen()
        //{
        //    //throw new NotImplementedException();
        //    MessageBox.Show("DocumentOpen2");
        //}

        public void Evento_DocumentOpen()
        {
            Globals.ThisAddIn.Application.DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(Metodo_DocumentOpen);
        }

        public void Evento_NewDocument()
        {
            ((Microsoft.Office.Interop.Word.ApplicationEvents4_Event)Globals.ThisAddIn.Application).NewDocument += new ApplicationEvents4_NewDocumentEventHandler(Metodo_NewDocument);
        }

        public void Evento_DocumentBeforeClose()
        {
            Globals.ThisAddIn.Application.DocumentBeforeClose += new ApplicationEvents4_DocumentBeforeCloseEventHandler(Metodo_DocumentBeforeClose);
        }

        public void Evento_DocumentChange()
        {
            Globals.ThisAddIn.Application.DocumentChange += new ApplicationEvents4_DocumentChangeEventHandler(Metodo_DocumentChange);
        }

        public void Evento_DocumentSync()
        {
            Globals.ThisAddIn.Application.DocumentSync += new ApplicationEvents4_DocumentSyncEventHandler(Metodo_DocumentSync);
        }
        public void Evento_ProtectedViewWindowActivate()
        {
            Globals.ThisAddIn.Application.ProtectedViewWindowActivate += new ApplicationEvents4_ProtectedViewWindowActivateEventHandler(Metodo_ProtectedViewWindowActivate);
        }
        public void Evento_WindowActivate()
        {
            Globals.ThisAddIn.Application.WindowActivate += new ApplicationEvents4_WindowActivateEventHandler(Metodo_WindowActivate);
        }
        public void Evento_WindowBeforeDoubleClick()
        {
            Globals.ThisAddIn.Application.WindowBeforeDoubleClick += new ApplicationEvents4_WindowBeforeDoubleClickEventHandler(Metodo_WindowBeforeDoubleClick);
        }

        public void Evento_WindowBeforeRightClick()
        {
            Globals.ThisAddIn.Application.WindowBeforeRightClick += new ApplicationEvents4_WindowBeforeRightClickEventHandler(Metodo_WindowBeforeRightClick);
        }
        public void Evento_WindowDeactivate()
        {
            Globals.ThisAddIn.Application.WindowDeactivate += new ApplicationEvents4_WindowDeactivateEventHandler(Metodo_WindowDeactivate);
        }
        public void Evento_WindowSelectionChange()
        {
            Globals.ThisAddIn.Application.WindowSelectionChange += new ApplicationEvents4_WindowSelectionChangeEventHandler(Metodo_WindowSelectionChange);
        }
        public void Evento_WindowSize()
        {
            Globals.ThisAddIn.Application.WindowSize += new ApplicationEvents4_WindowSizeEventHandler(Metodo_WindowSize);
        }
        public void Evento_SettingChanging()
        {
            Properties.Settings.Default.SettingChanging += new System.Configuration.SettingChangingEventHandler(SettingChanging);
        }
        public void Evento_TabDisposed()
        {
            Globals.Ribbons.Ribbon.tab.Disposed += new System.EventHandler(Metodo_TabDisposed);
        }
        public void Evento_RibbonClose()
        {
            Globals.Ribbons.Ribbon.Close += new System.EventHandler(Metodo_RibbonClose);
        }
        public void Evento_menu1ItemsLoading()
        {
            Globals.Ribbons.Ribbon.menu_campos.ItemsLoading += new RibbonControlEventHandler(Metodo_menu1ItemsLoading);
        }
        public void Evento_ContentControlOnEnter()
        {
            Globals.ThisAddIn.Application.ActiveWindow.Document.ContentControlOnEnter += new DocumentEvents2_ContentControlOnEnterEventHandler(Metodo_ContentControlOnEnter);
        }
        public void Evento_()
        {

        }

    }
}
