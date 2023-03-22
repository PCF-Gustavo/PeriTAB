using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using PeriTAB.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Deployment;
using static System.Net.Mime.MediaTypeNames;

    namespace PeriTAB
{
    
    public partial class Ribbon1
    {
        private bool first_time = true;
        String caminho_template = Path.GetTempPath() + "PeriTAB_Template_tmp.dotm";

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {            
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if (first_time)
            {
                File.WriteAllBytes(caminho_template, Properties.Resources.Normal_copy);
                Globals.ThisAddIn.Application.AddIns.Add(caminho_template);
                first_time = false;
            }
                        
            Globals.ThisAddIn.Application.Run("confere_numeracao_legendas");
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            if (first_time)
            {
                File.WriteAllBytes(caminho_template, Properties.Resources.Normal_copy);
                Globals.ThisAddIn.Application.AddIns.Add(caminho_template);
                first_time = false;
            }

            Globals.ThisAddIn.Application.Run("alinha_legenda");
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            if (first_time)
            {
                File.WriteAllBytes(caminho_template, Properties.Resources.Normal_copy);
                Globals.ThisAddIn.Application.AddIns.Add(caminho_template);
                first_time = false;
            }

            Globals.ThisAddIn.Application.Run("renomeia_documento");
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            if (first_time)
            {
                File.WriteAllBytes(caminho_template, Properties.Resources.Normal_copy);
                Globals.ThisAddIn.Application.AddIns.Add(caminho_template);
                first_time = false;
            }

            Globals.ThisAddIn.Application.Run("alterna_visualizacao_campos");
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            if (first_time)
            {
                File.WriteAllBytes(caminho_template, Properties.Resources.Normal_copy);
                Globals.ThisAddIn.Application.AddIns.Add(caminho_template);
                first_time = false;
            }

            Globals.ThisAddIn.Application.Run("alterna_destaque_campos");
        }

        private void button6_Click_1(object sender, RibbonControlEventArgs e)
        {
            if (first_time)
            {
                File.WriteAllBytes(caminho_template, Properties.Resources.Normal_copy);
                Globals.ThisAddIn.Application.AddIns.Add(caminho_template);
                first_time = false;
            }

            Globals.ThisAddIn.Application.Run("atualizar_todos_campos");
        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            if (first_time)
            {
                File.WriteAllBytes(caminho_template, Properties.Resources.Normal_copy);
                Globals.ThisAddIn.Application.AddIns.Add(caminho_template);
                first_time = false;
            }

            Globals.ThisAddIn.Application.Run("moeda_por_extenso");
        }

        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            if (first_time)
            {
                File.WriteAllBytes(caminho_template, Properties.Resources.Normal_copy);
                Globals.ThisAddIn.Application.AddIns.Add(caminho_template);
                first_time = false;
            }

            Globals.ThisAddIn.Application.Run("inteiro_por_extenso");
        }

        private void button9_Click(object sender, RibbonControlEventArgs e)
        {
            //Globals.ThisAddIn.Application.ActiveDocument.CopyStylesFromTemplate("C:\\Users\\gustavo.gvs.PF\\source\\repos\\PeriTAB\\PeriTAB\\Resources\\Normal_copy.dotm"); -- Comando para importar todos os estilos

            if (first_time)
            {
                File.WriteAllBytes(caminho_template, Properties.Resources.Normal_copy);
                Globals.ThisAddIn.Application.AddIns.Add(caminho_template);
                first_time = false;                
            }

            string[] aStyles = {"01 - corpo de texto", "02 - seções e subseções", "03 - citações", "04 - enumerações", "05 - figuras", "06 - legendas de figuras", "07 - notas de rodapé", "08 - legendas de tabelas", "09 - quesitos", "10 - anexo" };
            for (int i = 0; i <= aStyles.Length - 1; i++)
            {                
                Globals.ThisAddIn.Application.OrganizerCopy(caminho_template, Globals.ThisAddIn.Application.ActiveDocument.FullName, aStyles[i], WdOrganizerObject.wdOrganizerObjectStyles);
            }
            
        }

        private void button10_Click(object sender, RibbonControlEventArgs e)
        {
            if (first_time)
            {
                File.WriteAllBytes(caminho_template, Properties.Resources.Normal_copy);
                Globals.ThisAddIn.Application.AddIns.Add(caminho_template);
                first_time = false;
            }

            Globals.ThisAddIn.Application.Run("DelUnusedStyles");
        }
    }
}
