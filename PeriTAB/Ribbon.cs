using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Task = System.Threading.Tasks.Task;

namespace PeriTAB
{
    public partial class Ribbon
    {
        const string Nome_do_app = "PeriTAB";
        const string PeriTAB_Version = "1.2.6";
        const string Arquivo_PeriTAB_Template = "PeriTAB_Template_tmp.dotm";
        const string Arquivo_preferencias = "preferences.xml";
        const string Arquivo_lista_de_arquivos_para_excluir = "arquivos_para_excluir.txt";
        const string quote = "\"";
        const string slash = @"\";

        // Cria instância das classes
        private readonly Class_ContentControlOnExit_Event iClass_ContentControlOnExit_Event = new Class_ContentControlOnExit_Event();
        private readonly Class_CustomTaskPanes iClass_CustomTaskPanes = new Class_CustomTaskPanes();

        private IRibbonUI oRibbonUI;

        // Gerencia variáveis "globais"
        public class Variables
        {
            // Declara variáveis privadas
            private static readonly string private_caminho_template, private_caminho_AppData_Roaming_PeriTAB, private_caminho_preferences, private_caminho_lista_de_arquivos_para_excluir;
            private static Template private_Template_PeriTAB;
            private static readonly List<string> private_lista_arquivos_para_excluir;
            static Variables() // Bloco estático para definir o valor inicial das variáveis
            {
                private_caminho_template = Path.GetTempPath() + Arquivo_PeriTAB_Template;
                private_caminho_AppData_Roaming_PeriTAB = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), Nome_do_app);
                private_caminho_preferences = Path.Combine(private_caminho_AppData_Roaming_PeriTAB, Arquivo_preferencias);
                private_caminho_lista_de_arquivos_para_excluir = Path.Combine(private_caminho_AppData_Roaming_PeriTAB, Arquivo_lista_de_arquivos_para_excluir);
                if (File.Exists(private_caminho_lista_de_arquivos_para_excluir))
                {
                    private_lista_arquivos_para_excluir = File.ReadAllLines(private_caminho_lista_de_arquivos_para_excluir).ToList();
                }
                else
                {
                    private_lista_arquivos_para_excluir = new List<string>();
                }
            }

            // Declara variáveis públicas
            public static string Caminho_template { get { return private_caminho_template; } }
            public static Template Template_PeriTAB { get { return private_Template_PeriTAB; } set { private_Template_PeriTAB = value; } }
            public static string Caminho_AppData_Roaming_PeriTAB { get { return private_caminho_AppData_Roaming_PeriTAB; } }
            public static string Caminho_preferences { get { return private_caminho_preferences; } }
            public static List<string> Lista_arquivos_para_excluir { get { return private_lista_arquivos_para_excluir; } }
            public static string Caminho_lista_de_arquivos_para_excluir { get { return private_caminho_lista_de_arquivos_para_excluir; } }
        }



        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            oRibbonUI = e.RibbonUI;
            //MessageBox.Show("Ribbon_Load");

            //Escreve o Template na pasta tmp e adiciona ela como suplemento.
            //try { File.WriteAllBytes(Variables.Caminho_template, Properties.Resources.Normal); }
            //catch (IOException)
            //{
            //    if (!File.Exists(Variables.Caminho_template))
            //    {
            //        MessageBox.Show($"{Arquivo_PeriTAB_Template} não encontrado"); Globals.ThisAddIn.Application.Quit(); return;
            //    }
            //}
            File.WriteAllBytes(Variables.Caminho_template, Properties.Resources.Normal);
            Globals.ThisAddIn.Application.AddIns.Add(Variables.Caminho_template);
            Variables.Template_PeriTAB = Retonar_Template_do_Caminho(Variables.Caminho_template);

            Globals.Ribbons.Ribbon.label_nome.Label = Nome_do_app + " " + PeriTAB_Version;

            ThisAddIn.Excluir_arquivos_da_lista(Variables.Lista_arquivos_para_excluir);
        }

        private Template Retonar_Template_do_Caminho(string caminho_template)
        {
            foreach (Template template in Globals.ThisAddIn.Application.Templates)
            {
                if (template.Name == Path.GetFileName(caminho_template))
                {
                    return template;
                }
            }
            return null;
        }

        public BuildingBlock Inserir_autotexto(Range range, string autotextName)
        {
            if (range != null)
            {
                for (int i = 1; i <= Variables.Template_PeriTAB.BuildingBlockEntries.Count; i++)
                {
                    BuildingBlock bb = Variables.Template_PeriTAB.BuildingBlockEntries.Item(i);
                    if (bb.Name == autotextName)
                    {
                        bb.Insert(range);
                        Range Previous = Globals.ThisAddIn.Application.Selection.Range.Previous();
                        if (Previous != null) if (Previous.Fields.Count > 0) Previous.Words[1].Fields.Update();
                        return bb;
                    }
                }
            }
            return null;
        }

        private async void Button_teste_Click(object sender, RibbonControlEventArgs e)
        {
            await Executar_Ribbon_com_UI_responsiva(sender, e, async progress =>
            {
                int total = 100000;
                for (int i = 0; i < total; i++)
                {
                    double result = 0;

                    for (int j = 0; j < 1000; j++)
                    {
                        result += Math.Sqrt(j) * Math.Cos(j);
                    }

                    await progress.Tick_50ms((int)((i * 10.0) / total));
                }
            }, barra_de_progresso: true, desabilitar_ScreenUpdating: false, desabilitar_TrackRevisions: false);
        }

        
        private async Task Executar_Ribbon_com_UI_responsiva(
            object sender,
            RibbonControlEventArgs e,
            Func<IRibbonTick, Task> action,
            bool barra_de_progresso = false,
            bool desabilitar_ScreenUpdating = false,
            bool desabilitar_TrackRevisions = false
            )
        {
            #if DEBUG
                Stopwatch Stopwatch = Stopwatch.StartNew();
            #endif

            RibbonButton ribbonButton = (RibbonButton)sender;
            RibbonMenu ribbonMenu = ribbonButton.Parent as RibbonMenu;

            Image imagemInicial;
            string mensagemStatusBar = "";
            bool success = true;
            string mensagemFalha = "";

            // ================= UI inicial =================
            if (ribbonMenu != null)
            {
                imagemInicial = ribbonMenu.Image;
                ribbonMenu.Image = Properties.Resources.loading;
                ribbonMenu.Enabled = false;
                mensagemStatusBar += ribbonMenu.Label + "/";
            }
            else
            {
                imagemInicial = ribbonButton.Image;
                ribbonButton.Image = Properties.Resources.loading;
                ribbonButton.Enabled = false;
            }

            mensagemStatusBar += ribbonButton.Label + ": ";

            oRibbonUI.InvalidateControl(e.Control.Id);
            await Task.Yield();

            // ================= Estado inicial Word =================
            bool screenUpdatingInicial = Globals.ThisAddIn.Application.ScreenUpdating;
            bool trackRevisionsInicial = Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions;
            WdCursorType CursorInicial = Globals.ThisAddIn.Application.System.Cursor;

            if (desabilitar_ScreenUpdating)
                Globals.ThisAddIn.Application.ScreenUpdating = false;

            if (desabilitar_TrackRevisions)
                Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions = false;

            Range selecaoInicial = Globals.ThisAddIn.Application.Selection.Range.Duplicate;

            // ================= Progress =================
            IProgress<int> progress = null;
            IRibbonTick tick;

            if (barra_de_progresso)
            {
                Globals.ThisAddIn.Application.StatusBar =
                    mensagemStatusBar + Barra_de_progresso(0);

                progress = new Progress<int>(p =>
                {
                    Globals.ThisAddIn.Application.StatusBar =
                        mensagemStatusBar + Barra_de_progresso(p);
                });

                tick = new RibbonTickComProgresso(progress);
            }
            else
            {
                tick = new RibbonTickNenhum();
            }

            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord();

            try
            {
                await action(tick);
            }
            catch (Exception ex)
            {
                success = false;
                mensagemFalha = ex.Message;
                selecaoInicial.Select();
            }
            finally
            {
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();

                if (desabilitar_ScreenUpdating)
                    Globals.ThisAddIn.Application.ScreenUpdating = screenUpdatingInicial;

                if (desabilitar_TrackRevisions)
                    Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions = trackRevisionsInicial;

                #if DEBUG
                Stopwatch.Stop();
                    double tempo = Stopwatch.Elapsed.TotalSeconds;
                #endif

                if (barra_de_progresso)
                {
                    string StatusBar = mensagemStatusBar + Barra_de_progresso(success ? 10 : 0) + (success ? " Sucesso" : " Falha");
                        #if DEBUG
                            StatusBar += $" (Tempo: {tempo:F2}s)";
                        #endif
                    Globals.ThisAddIn.Application.StatusBar = StatusBar;
                }

                if (!string.IsNullOrEmpty(mensagemFalha))
                {
                    MessageBox.Show(
                        new WindowWrapper(
                            new IntPtr(Globals.ThisAddIn.Application.ActiveWindow.Hwnd)),
                        mensagemFalha,
                        ribbonButton.Label,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }

                if (ribbonMenu != null)
                {
                    ribbonMenu.Image = imagemInicial;
                    ribbonMenu.Enabled = true;
                }
                else
                {
                    ribbonButton.Image = imagemInicial;
                    ribbonButton.Enabled = true;
                }
            }
        }

        private string Barra_de_progresso(int progress)
        {
            char filledSquare = (char)0x2588;  // Caractere '█' (quadrado preenchido).
            char emptySquare = (char)0x2591;   // Caractere '░' (quadrado não preenchido).
            string progressBar = new string(filledSquare, progress) + new string(emptySquare, 10 - progress); // Cria a "barra de progresso".
            return progressBar;
        }

    }
}