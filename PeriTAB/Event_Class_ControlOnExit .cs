using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PeriTAB
{
    internal class Event_Class_ControlOnExit
    {
        private static Dictionary<string, string> dict_Unidade_e_Unidade_da_PF = new Dictionary<string, string>()
        {
            { "INC/DITEC/PF", "DITEC - INSTITUTO NACIONAL DE CRIMINALÍSTICA" },
        };
        private static Dictionary<string, string> dict_Unidade_da_PF_e_Unidade = dict_Unidade_e_Unidade_da_PF.ToDictionary(par => par.Value, par => par.Key);
        public void Metodo_ControlOnExit()
        {
            // Configura o evento global para monitorar quando o controle é alterado
            Globals.ThisAddIn.Application.ActiveDocument.ContentControlOnExit += (ContentControl contentControl, ref bool cancel) =>
            {
                VincularLista(contentControl, "Unidade", "Unidade da PF", dict_Unidade_e_Unidade_da_PF);
                VincularLista(contentControl, "Unidade da PF", "Unidade", dict_Unidade_da_PF_e_Unidade);
            };

        }

        private void VincularLista(ContentControl contentControl_lista1, string titulo_lista1, string titulo_lista2, Dictionary<string,string> dicionario)
        {
            if (contentControl_lista1.Title == titulo_lista1)
            {
                ContentControl contentControl_lista2 = GetContentControl(titulo_lista2);

                if (contentControl_lista2 != null)
                {
                    foreach (KeyValuePair<string, string> item in dicionario)
                    {
                        if (contentControl_lista1.Range.Text == item.Key)
                        {
                            ChangeEntry(contentControl_lista2, item.Value);
                            break;
                        }
                    }
                }
            }
        }

        private ContentControl GetContentControl(string titulo_do_controle)
        {
            foreach (Section section in Globals.ThisAddIn.Application.ActiveDocument.Sections)
            {
                foreach (ContentControl control in section.Range.ContentControls)
                {
                    if (control.Title == titulo_do_controle)
                    {
                        return control;
                    }
                }
                foreach (HeaderFooter Header in section.Headers)
                {
                    foreach (ContentControl control in Header.Range.ContentControls)
                    {
                        if (control.Title == titulo_do_controle)
                        {
                            return control;
                        }
                    }
                }
                foreach (HeaderFooter Footer in section.Footers)
                {
                    foreach (ContentControl control in Footer.Range.ContentControls)
                    {
                        if (control.Title == titulo_do_controle)
                        {
                            return control;
                        }
                    }
                }
            }
            return null;
        }

        private void ChangeEntry(ContentControl ContentControl, string valor_da_lista)
        {
            foreach (ContentControlListEntry entry in ContentControl.DropdownListEntries)
            {
                if (entry.Text == valor_da_lista)
                {
                    ContentControl.DropdownListEntries[entry.Index].Select();
                    break;
                }
            }
        }

    }
}
