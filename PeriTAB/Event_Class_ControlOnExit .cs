using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PeriTAB
{
    internal class Class_ContentControlOnExit_Event
    {
        private static Dictionary<string, string> dict_Unidade_e_Unidade_da_PF = new Dictionary<string, string>()
        {
             { "INC/DITEC/PF", "DITEC - INSTITUTO NACIONAL DE CRIMINALÍSTICA" }
            ,{ "SETEC/SR/PF/AC" , "SUPERINTENDÊNCIA REGIONAL NO ACRE" }
            ,{ "SETEC/SR/PF/AL" , "SUPERINTENDÊNCIA REGIONAL NO ALAGOAS" }
            ,{ "SETEC/SR/PF/AP" , "SUPERINTENDÊNCIA REGIONAL NO AMAPÁ" }
            ,{ "SETEC/SR/PF/AM" , "SUPERINTENDÊNCIA REGIONAL NO AMAZONAS" }
            ,{ "SETEC/SR/PF/BA" , "SUPERINTENDÊNCIA REGIONAL NO BAHIA" }
            ,{ "SETEC/SR/PF/CE" , "SUPERINTENDÊNCIA REGIONAL NO CEARÁ" }
            ,{ "SETEC/SR/PF/DF" , "SUPERINTENDÊNCIA REGIONAL NO DISTRITO FEDERAL" }
            ,{ "SETEC/SR/PF/ES" , "SUPERINTENDÊNCIA REGIONAL NO ESPÍRITO SANTO" }
            ,{ "SETEC/SR/PF/GO" , "SUPERINTENDÊNCIA REGIONAL NO GOIÁS" }
            ,{ "SETEC/SR/PF/MA" , "SUPERINTENDÊNCIA REGIONAL NO MARANHÃO" }
            ,{ "SETEC/SR/PF/MT" , "SUPERINTENDÊNCIA REGIONAL NO MATO GROSSO" }
            ,{ "SETEC/SR/PF/MS" , "SUPERINTENDÊNCIA REGIONAL NO MATO GROSSO DO SUL" }
            ,{ "SETEC/SR/PF/MG" , "SUPERINTENDÊNCIA REGIONAL NO MINAS GERAIS" }
            ,{ "SETEC/SR/PF/PA" , "SUPERINTENDÊNCIA REGIONAL NO PARÁ" }
            ,{ "SETEC/SR/PF/PB" , "SUPERINTENDÊNCIA REGIONAL NO PARAÍBA" }
            ,{ "SETEC/SR/PF/PR" , "SUPERINTENDÊNCIA REGIONAL NO PARANÁ" }
            ,{ "SETEC/SR/PF/PE" , "SUPERINTENDÊNCIA REGIONAL NO PERNAMBUCO" }
            ,{ "SETEC/SR/PF/PI" , "SUPERINTENDÊNCIA REGIONAL NO PIAUÍ" }
            ,{ "SETEC/SR/PF/RJ" , "SUPERINTENDÊNCIA REGIONAL NO RIO DE JANEIRO" }
            ,{ "SETEC/SR/PF/RN" , "SUPERINTENDÊNCIA REGIONAL NO RIO GRANDE DO NORTE" }
            ,{ "SETEC/SR/PF/RS" , "SUPERINTENDÊNCIA REGIONAL NO RIO GRANDE DO SUL" }
            ,{ "SETEC/SR/PF/RO" , "SUPERINTENDÊNCIA REGIONAL NO RONDÔNIA" }
            ,{ "SETEC/SR/PF/RR" , "SUPERINTENDÊNCIA REGIONAL NO RORAIMA" }
            ,{ "SETEC/SR/PF/SC" , "SUPERINTENDÊNCIA REGIONAL NO SANTA CATARINA" }
            ,{ "SETEC/SR/PF/SP" , "SUPERINTENDÊNCIA REGIONAL NO SÃO PAULO" }
            ,{ "SETEC/SR/PF/SE" , "SUPERINTENDÊNCIA REGIONAL NO SERGIPE" }
            ,{ "SETEC/SR/PF/TO" , "SUPERINTENDÊNCIA REGIONAL NO TOCANTINS" }
            ,{ "NUTEC/DPF/ARU/SP" , "DELEGACIA DE POLÍCIA FEDERAL EM ARAÇATUBA" }
            ,{ "NUTEC/DPF/CAS/SP" , "DELEGACIA DE POLÍCIA FEDERAL EM CAMPINAS" }
            ,{ "NUTEC/DPF/DRS/MS" , "DELEGACIA DE POLÍCIA FEDERAL EM DOURADOS" }
            ,{ "NUTEC/DPF/FIG/PR" , "DELEGACIA DE POLÍCIA FEDERAL EM FOZ DO IGUAÇU" }
            ,{ "NUTEC/DPF/GRA/PR" , "DELEGACIA DE POLÍCIA FEDERAL EM GUAÍRA" }
            ,{ "NUTEC/DPF/JFA/MG" , "DELEGACIA DE POLÍCIA FEDERAL EM JUIZ DE FORA" }
            ,{ "NUTEC/DPF/JZO/BA" , "DELEGACIA DE POLÍCIA FEDERAL EM JUAZEIRO" }
            ,{ "NUTEC/DPF/LDA/PR" , "DELEGACIA DE POLÍCIA FEDERAL EM LONDRINA" }
            ,{ "NUTEC/DPF/MII/SP" , "DELEGACIA DE POLÍCIA FEDERAL EM MARÍLIA" }
            ,{ "NUTEC/DPF/PDE/SP" , "DELEGACIA DE POLÍCIA FEDERAL EM PRESIDENTE PRUDENTE" }
            ,{ "NUTEC/DPF/PFO/RS" , "DELEGACIA DE POLÍCIA FEDERAL EM PASSO FUNDO" }
            ,{ "NUTEC/DPF/PTS/RS" , "DELEGACIA DE POLÍCIA FEDERAL EM PELOTAS" }
            ,{ "NUTEC/DPF/RPO/SP" , "DELEGACIA DE POLÍCIA FEDERAL EM RIBEIRÃO PRETO" }
            ,{ "NUTEC/DPF/SIC/MT" , "DELEGACIA DE POLÍCIA FEDERAL EM SINOP" }
            ,{ "NUTEC/DPF/SJK/SP" , "DELEGACIA DE POLÍCIA FEDERAL EM SÃO JOSÉ DOS CAMPOS" }
            ,{ "NUTEC/DPF/SMA/RS" , "DELEGACIA DE POLÍCIA FEDERAL EM SANTA MARIA" }
            ,{ "NUTEC/DPF/SNM/PA" , "DELEGACIA DE POLÍCIA FEDERAL EM SANTARÉM" }
            ,{ "NUTEC/DPF/SOD/SP" , "DELEGACIA DE POLÍCIA FEDERAL EM SOROCABA" }
            ,{ "NUTEC/DPF/STS/SP" , "DELEGACIA DE POLÍCIA FEDERAL EM SANTOS" }
            ,{ "NUTEC/DPF/UDI/MG" , "DELEGACIA DE POLÍCIA FEDERAL EM UBERLÂNDIA" }
            ,{ "NUTEC/DPF/VLA/RO" , "DELEGACIA DE POLÍCIA FEDERAL EM VILHENA" }

        };
        private static Dictionary<string, string> dict_Unidade_da_PF_e_Unidade = dict_Unidade_e_Unidade_da_PF.ToDictionary(par => par.Value, par => par.Key); // Dicionario invertido

        private static Dictionary<string, string> dict_Unidade_e_Tipo_de_Unidade_de_criminalistica = new Dictionary<string, string>()
        {
             { "SETEC/SR/PF/AC" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/AL" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/AP" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/AM" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/BA" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/CE" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/DF" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/ES" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/GO" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/MA" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/MT" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/MS" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/MG" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/PA" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/PB" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/PR" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/PE" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/PI" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/RJ" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/RN" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/RS" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/RO" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/RR" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/SC" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/SP" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/SE" , "SETOR TECNICO CIENTIFICO" }
            ,{ "SETEC/SR/PF/TO" , "SETOR TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/ARU/SP" , "NUCLEO TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/CAS/SP" , "NUCLEO TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/DRS/MS" , "NUCLEO TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/FIG/PR" , "NUCLEO TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/GRA/PR" , "NUCLEO TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/JFA/MG" , "NUCLEO TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/JZO/BA" , "NUCLEO TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/LDA/PR" , "NUCLEO TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/MII/SP" , "NUCLEO TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/PDE/SP" , "NUCLEO TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/PFO/RS" , "NUCLEO TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/PTS/RS" , "NUCLEO TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/RPO/SP" , "NUCLEO TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/SIC/MT" , "NUCLEO TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/SJK/SP" , "NUCLEO TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/SMA/RS" , "NUCLEO TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/SNM/PA" , "NUCLEO TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/SOD/SP" , "NUCLEO TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/STS/SP" , "NUCLEO TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/UDI/MG" , "NUCLEO TECNICO CIENTIFICO" }
            ,{ "NUTEC/DPF/VLA/RO" , "NUCLEO TECNICO CIENTIFICO" }
        };
        
        public void Metodo_ContentControlOnExit()
        {
            // Configura o evento global para monitorar quando o controle é alterado
            Globals.ThisAddIn.Application.ActiveDocument.ContentControlOnExit += (ContentControl contentControl, ref bool cancel) =>
            {
                VincularLista(contentControl, "Unidade", "Unidade da PF", dict_Unidade_e_Unidade_da_PF);
                VincularLista(contentControl, "Unidade da PF", "Unidade", dict_Unidade_da_PF_e_Unidade);
                VincularLista(contentControl, "Unidade", "Tipo de unidade de criminalistica", dict_Unidade_e_Tipo_de_Unidade_de_criminalistica);
                Add_or_remove_ultima_linha_cabecalho(contentControl);
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

        private void Add_or_remove_ultima_linha_cabecalho(ContentControl ContentControl)
        {
            if (ContentControl.Title == "Unidade" || ContentControl.Title == "Unidade da PF")
            {
                ContentControl controle_Unidade_da_PF = GetContentControl("Unidade da PF");

                Paragraph paragraph = controle_Unidade_da_PF.Range.Paragraphs[1].Next();
                if (controle_Unidade_da_PF.Range.Text == "DITEC - INSTITUTO NACIONAL DE CRIMINALÍSTICA")
                {
                    if (paragraph != null)
                    {
                        paragraph.Range.Delete();
                    }
                }
                else 
                {
                    // Procurar pelo autotexto Numero_de_paginas_por_extenso no template_PeriTAB
                    string autotextName = "Tipo de unidade de criminalistic";
                    BuildingBlockEntries buildingBlockEntries = Ribbon.Variables.Template_PeriTAB.BuildingBlockEntries;
                    for (int i = 1; i <= buildingBlockEntries.Count; i++)
                    {
                        BuildingBlock bb = buildingBlockEntries.Item(i);
                        if (bb.Name == autotextName)
                        {
                            if (paragraph == null)
                            {
                                controle_Unidade_da_PF.Range.Paragraphs[1].Range.InsertParagraphAfter();
                                bb.Insert(controle_Unidade_da_PF.Range.Paragraphs[1].Next().Range);
                            }
                            else
                            {
                                paragraph.Range.Text = "";
                                bb.Insert(paragraph.Range);
                            }
                        }
                    }
                }
                    
            }

        }

    }
}
