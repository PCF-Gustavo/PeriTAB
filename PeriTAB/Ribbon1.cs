using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Drawing;
using System.IO;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Controls;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;


namespace PeriTAB
{    
    public partial class Ribbon1
    {
        public class Variables
        {
            private static string var = Path.GetTempPath() + "PeriTAB_Template_tmp.dotm";
            public static string caminho_template { get { return var; } set { var = value; } }
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //Escreve o Template na pasta tmp e adiciona ela como suplemento.
            File.WriteAllBytes(Variables.caminho_template, Properties.Resources.Normal);
            Globals.ThisAddIn.Application.AddIns.Add(Variables.caminho_template);

            // Escreve o número da versão
            System.Version publish_version = Assembly.GetExecutingAssembly().GetName().Version;
            Globals.Ribbons.Ribbon1.label1.Label = "PeriTAB " + publish_version.Major + "." + publish_version.Minor + "." + publish_version.Build;
                        
        }

        private void button_confere_num_legenda_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("confere_numeracao_legendas");
        }

        private void button_alinha_legenda_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("alinha_legenda");
        }

        private void button_renomeia_documento_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("renomeia_documento");
        }

        private void button_atualiza_campos_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("atualiza_todos_campos");
        }
        private void checkBox_destaca_campos_Click(object sender, RibbonControlEventArgs e)
        {
            var Botao_checkBox = (Microsoft.Office.Tools.Ribbon.RibbonCheckBox)sender;
            if (Botao_checkBox.Checked == true) Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading = (WdFieldShading)1;
            if (Botao_checkBox.Checked == false) Globals.ThisAddIn.Application.ActiveWindow.View.FieldShading = (WdFieldShading)2;
        }

        private void checkBox_vercodigo_campos_Click(object sender, RibbonControlEventArgs e)
        {
            var Botao_checkBox = (Microsoft.Office.Tools.Ribbon.RibbonCheckBox)sender;
            if (Botao_checkBox.Checked == true) Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes = true;
            if (Botao_checkBox.Checked == false) Globals.ThisAddIn.Application.ActiveWindow.View.ShowFieldCodes = false;
        }

        private void button_moeda_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("moeda_por_extenso");
        }

        private void button_inteiro_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("inteiro_por_extenso");
        }

        private void button_importa_estilos_Click(object sender, RibbonControlEventArgs e)
        {
            string[] aStyles = { "01 - Sem Formatação (PeriTAB)", "02 - Corpo do Texto (PeriTAB)", "03 - Citações (PeriTAB)", "04 - Seções (PeriTAB)", "05 - Enumerações (PeriTAB)", "06 - Figuras (PeriTAB)", "07 - Legendas de Figuras (PeriTAB)", "08 - Legendas de Tabelas (PeriTAB)", "09 - Quesitos (PeriTAB)", "Normal", "Texto de nota de rodapé", "Legenda" };
            for (int i = 0; i <= aStyles.Length - 1; i++)
            {
                Globals.ThisAddIn.Application.OrganizerCopy(Variables.caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, aStyles[i], WdOrganizerObject.wdOrganizerObjectStyles);
            }
        }

        private void button_limpa_estilos_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.Run("limpa_estilos");
        }

        private void toggleButton_estilos_Click(object sender, RibbonControlEventArgs e)
        {
            var botao_toggle = (Microsoft.Office.Tools.Ribbon.RibbonToggleButton)sender;
            if (botao_toggle.Checked == true) Globals.ThisAddIn.TaskPane1.Visible = true;
            if (botao_toggle.Checked == false) Globals.ThisAddIn.TaskPane1.Visible = false;
        }

        private void button_inserir_sumario_Click(object sender, RibbonControlEventArgs e)
        {           

        }
        public static byte[] ObjectToByteArray(Object obj)
        {
            BinaryFormatter bf = new BinaryFormatter();
            using (var ms = new MemoryStream())
            {
                bf.Serialize(ms, obj);
                return ms.ToArray();
            }
        }

        private void dropDown1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }



        private void button_cola_imagem_Click(object sender, RibbonControlEventArgs e)
        {
            if (System.Windows.Clipboard.ContainsData("FileDrop"))
            {
                object obj = System.Windows.Clipboard.GetData("FileDrop");
                string[] pathfile = (string[])obj;
                if (pathfile.Length == 1) //Se tem um arquivo no clipboard
                {
                    string extensao = pathfile[0].Substring(pathfile[0].Length - 4);
                    if (extensao == ".jpg" | extensao == "jpeg" | extensao == ".png" | extensao == ".bmp" | extensao == ".gif") //Se tem extensao de imagem
                    {
                        Globals.ThisAddIn.Application.ScreenUpdating = false;



                        //Globals.ThisAddIn.Application.Options.save

                        //Globals.ThisAddIn.Application.Selection.Paste();
                        ////Clipboard.Clear(); 
                        ////Globals.ThisAddIn.Application.Selection.Paste();


                        //System.Drawing.Image img = System.Drawing.Image.FromFile(pathfile[0]);
                        //MessageBox.Show(img.HorizontalResolution.ToString());
                        //MessageBox.Show(img.VerticalResolution.ToString());
                        //Bitmap img_bitmap = new Bitmap(img);
                        //img_bitmap.SetResolution(20.0F, 20.0F);


                        //Clipboard.SetImage(img_bitmap.GetThumbnailImage(20,20, gethum as, callbackData));


                        //Globals.ThisAddIn.Application.Selection.Co

                        //Globals.ThisAddIn.Application.Selection.Paste();

                        if (checkBox_largura.Checked)
                        {
                            InlineShape imagem = Globals.ThisAddIn.Application.Selection.InlineShapes.AddPicture(pathfile[0]);


                            MsoTriState LockAspectRatio_i = imagem.LockAspectRatio;
                            imagem.LockAspectRatio = (MsoTriState)1;

                            string larg_string = Globals.Ribbons.Ribbon1.editBox_largura.Text;
                            float.TryParse(larg_string, out float larg);
                            imagem.Width = Globals.ThisAddIn.Application.CentimetersToPoints(larg);
                            imagem.LockAspectRatio = LockAspectRatio_i;
                        }

                        if (checkBox_altura.Checked)
                        {
                            InlineShape imagem = Globals.ThisAddIn.Application.Selection.InlineShapes.AddPicture(pathfile[0]);

                            MsoTriState LockAspectRatio_i = imagem.LockAspectRatio;
                            imagem.LockAspectRatio = (MsoTriState)1;
                            string alt_string = Globals.Ribbons.Ribbon1.editBox_altura.Text;
                            float.TryParse(alt_string, out float alt);
                            imagem.Height = Globals.ThisAddIn.Application.CentimetersToPoints(alt);

                            imagem.LockAspectRatio = LockAspectRatio_i;
                        }



                        Globals.ThisAddIn.Application.ScreenUpdating = true;
                    }
                }
            }



            //Globals.ThisAddIn.Application.Options.UpdateFieldsAtPrint



            //if (System.Windows.Clipboard.ContainsImage()) Globals.ThisAddIn.Application.Selection.Paste();
            //try
            //{
            //    System.Windows.IDataObject Clipboard_content = System.Windows.Clipboard.GetDataObject();
            //    string[] Clipboard_content_formats = Clipboard_content.GetFormats();

            //    //foreach (string str in Clipboard_content_formats)
            //    //{
            //    //    MessageBox.Show(str);
            //    //}

            //    for (int i = 0; i <= Clipboard_content_formats.Length - 1; i++)
            //    {
            //        if (System.Windows.Clipboard.ContainsData(Clipboard_content_formats[i])) MessageBox.Show("Contém dados do tipo '" + Clipboard_content_formats[i] + "' = " + System.Windows.Clipboard.ContainsData(Clipboard_content_formats[i]).ToString() + Environment.NewLine + Environment.NewLine + System.Windows.Clipboard.GetData(Clipboard_content_formats[i]).ToString()); 
            //        else MessageBox.Show("Contém dados do tipo '" + Clipboard_content_formats[i] + "' = " + System.Windows.Clipboard.ContainsData(Clipboard_content_formats[i]).ToString());

            //        if (System.Windows.Clipboard.ContainsData(Clipboard_content_formats[i]) & System.Windows.Clipboard.GetData(Clipboard_content_formats[i]).ToString() == "System.String[]")
            //        {
            //            object obj = System.Windows.Clipboard.GetData(Clipboard_content_formats[i]);
            //            string[] stringArray = (string[])obj;
            //            foreach (string str in stringArray)
            //            {
            //                MessageBox.Show(str);
            //            }
            //        }

            //    }

            //if (System.Windows.Clipboard.ContainsData("FileDrop")) 
            //{
            //    object obj = System.Windows.Clipboard.GetData("FileDrop");
            //    string[] stringArray = (string[])obj;



            //} 


            //    //MessageBox.Show(System.Windows.Clipboard.GetData("Bitmap").GetType().ToString());
            //    string[] a = System.Windows.Clipboard.GetDataObject().GetFormats();




            ////MessageBox.Show(w.ToString());
            //string[] b = w.GetFormats();
            //foreach (string oo in b)
            //{
            //    MessageBox.Show(oo);
            //}



            //MessageBox.Show(w.GetFormats().ToString());

            //for (int i = 0; i <= a.Length - 1; i++)
            //{
            //    MessageBox.Show("Contém dados do tipo '" + a[i] + "' = " + System.Windows.Clipboard.ContainsData(a[i]).ToString());
            //    if (System.Windows.Clipboard.ContainsData(a[i])) MessageBox.Show(System.Windows.Clipboard.GetData(a[i]).ToString());




            //if (System.Windows.Clipboard.GetData(a[i]).ToString() == "System.String[]") 
            //{
            //    System.String dd = System.Windows.Clipboard.GetData(a[i]).ToString();
            //foreach (string value in System.Windows.Clipboard.GetData(a[i]).ToString())
            //{
            //    MessageBox.Show(value.ToString());
            //}

            //MessageBox.Show(System.Windows.Clipboard.GetData(a[i]));
            //object rr = System.Windows.Clipboard.GetData(a[i]);
            //byte[] bb = ObjectToByteArray(rr);
            //MessageBox.Show(bb.ToString());
            //System.Runtime.Serialization.Formatters.Binary.BinaryFormatter edd = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();
            //System.String[] a = System.String[];
            //}
            //MessageBox.Show(System.Windows.Clipboard.ContainsData(a[i]).ToString());
            //MessageBox.Show(a[i] + " = " + System.Windows.Clipboard.GetDataObject().GetData(a[i]));



            //}


            //}
            //catch 
            //{
            //    MessageBox.Show("null");
            //}



        }      

        private void checkBox_largura_Click(object sender, RibbonControlEventArgs e)
        {
            if (checkBox_largura.Checked)
            {
                checkBox_altura.Checked = false;
                editBox_altura.Enabled = false;
                editBox_altura.Text = "";
                editBox_largura.Enabled = true;
                editBox_largura.Text = "10";
            }
            else
            {
                checkBox_largura.Checked = true;
            }
        }

        private void checkBox_altura_Click(object sender, RibbonControlEventArgs e)
        {
            if (checkBox_altura.Checked)
            {
                checkBox_largura.Checked = false;
                editBox_largura.Enabled= false;
                editBox_largura.Text = "";
                editBox_altura.Enabled = true;
                editBox_altura.Text = "10";
            }
        else
        {
                checkBox_altura.Checked = true;
            }
}
    }
}
