using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Linq;

namespace PeriTAB
{
    internal class Class_ContentControlOnExit_Event
    {
        public static Dictionary<string, string> dict_Unidade_e_Unidade_da_PF = new Dictionary<string, string>()
        {
             { "INC/DITEC/PF", "DITEC - INSTITUTO NACIONAL DE CRIMINALÍSTICA" }
            ,{ "SETEC/SR/PF/AC" , "SUPERINTENDÊNCIA REGIONAL NO ACRE" }
            ,{ "SETEC/SR/PF/AL" , "SUPERINTENDÊNCIA REGIONAL EM ALAGOAS" }
            ,{ "SETEC/SR/PF/AP" , "SUPERINTENDÊNCIA REGIONAL NO AMAPÁ" }
            ,{ "SETEC/SR/PF/AM" , "SUPERINTENDÊNCIA REGIONAL NO AMAZONAS" }
            ,{ "SETEC/SR/PF/BA" , "SUPERINTENDÊNCIA REGIONAL NA BAHIA" }
            ,{ "SETEC/SR/PF/CE" , "SUPERINTENDÊNCIA REGIONAL NO CEARÁ" }
            ,{ "SETEC/SR/PF/DF" , "SUPERINTENDÊNCIA REGIONAL NO DISTRITO FEDERAL" }
            ,{ "SETEC/SR/PF/ES" , "SUPERINTENDÊNCIA REGIONAL NO ESPÍRITO SANTO" }
            ,{ "SETEC/SR/PF/GO" , "SUPERINTENDÊNCIA REGIONAL EM GOIÁS" }
            ,{ "SETEC/SR/PF/MA" , "SUPERINTENDÊNCIA REGIONAL NO MARANHÃO" }
            ,{ "SETEC/SR/PF/MT" , "SUPERINTENDÊNCIA REGIONAL NO MATO GROSSO" }
            ,{ "SETEC/SR/PF/MS" , "SUPERINTENDÊNCIA REGIONAL NO MATO GROSSO DO SUL" }
            ,{ "SETEC/SR/PF/MG" , "SUPERINTENDÊNCIA REGIONAL EM MINAS GERAIS" }
            ,{ "SETEC/SR/PF/PA" , "SUPERINTENDÊNCIA REGIONAL NO PARÁ" }
            ,{ "SETEC/SR/PF/PB" , "SUPERINTENDÊNCIA REGIONAL NA PARAÍBA" }
            ,{ "SETEC/SR/PF/PR" , "SUPERINTENDÊNCIA REGIONAL NO PARANÁ" }
            ,{ "SETEC/SR/PF/PE" , "SUPERINTENDÊNCIA REGIONAL EM PERNAMBUCO" }
            ,{ "SETEC/SR/PF/PI" , "SUPERINTENDÊNCIA REGIONAL NO PIAUÍ" }
            ,{ "SETEC/SR/PF/RJ" , "SUPERINTENDÊNCIA REGIONAL NO RIO DE JANEIRO" }
            ,{ "SETEC/SR/PF/RN" , "SUPERINTENDÊNCIA REGIONAL NO RIO GRANDE DO NORTE" }
            ,{ "SETEC/SR/PF/RS" , "SUPERINTENDÊNCIA REGIONAL NO RIO GRANDE DO SUL" }
            ,{ "SETEC/SR/PF/RO" , "SUPERINTENDÊNCIA REGIONAL EM RONDÔNIA" }
            ,{ "SETEC/SR/PF/RR" , "SUPERINTENDÊNCIA REGIONAL EM RORAIMA" }
            ,{ "SETEC/SR/PF/SC" , "SUPERINTENDÊNCIA REGIONAL EM SANTA CATARINA" }
            ,{ "SETEC/SR/PF/SP" , "SUPERINTENDÊNCIA REGIONAL EM SÃO PAULO" }
            ,{ "SETEC/SR/PF/SE" , "SUPERINTENDÊNCIA REGIONAL EM SERGIPE" }
            ,{ "SETEC/SR/PF/TO" , "SUPERINTENDÊNCIA REGIONAL EM TOCANTINS" }
            ,{ "NUTEC/DPF/ARU/SP" , "DELEGACIA DE POLÍCIA FEDERAL EM ARAÇATUBA" }
            ,{ "NUTEC/DPF/CAS/SP" , "DELEGACIA DE POLÍCIA FEDERAL EM CAMPINAS" }
            ,{ "NUTEC/DPF/DRS/MS" , "DELEGACIA DE POLÍCIA FEDERAL EM DOURADOS" }
            ,{ "NUTEC/DPF/FIG/PR" , "DELEGACIA DE POLÍCIA FEDERAL EM FOZ DO IGUAÇU" }
            ,{ "NUTEC/DPF/GRA/PR" , "DELEGACIA DE POLÍCIA FEDERAL EM GUAÍRA" }
            ,{ "NUTEC/DPF/JFA/MG" , "DELEGACIA DE POLÍCIA FEDERAL EM JUIZ DE FORA" }
            ,{ "NUTEC/DPF/JZO/BA" , "DELEGACIA DE POLÍCIA FEDERAL EM JUAZEIRO" }
            ,{ "NUTEC/DPF/JNE/CE" , "DELEGACIA DE POLÍCIA FEDERAL EM JUAZEIRO DO NORTE" }
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

        public static List<string> Lista_Unidade = new List<string>(dict_Unidade_e_Unidade_da_PF.Keys);

        public static List<string> Lista_Subtitulos = new List<string>()
        {
             { "CONTÁBIL-FINANCEIRO" }
            ,{ "REGISTROS DE ÁUDIO E IMAGENS" }
            ,{ "ENGENHARIA" }
            ,{ "INFORMÁTICA" }
            ,{ "QUÍMICA FORENSE" }
            ,{ "LOCAL DE CRIME" }
            ,{ "MEIO AMBIENTE" }
            ,{ "PATRIMÔNIO HISTÓRICO, ARTÍSTICO E CULTURA" }
            ,{ "VEÍCULOS" }
            ,{ "DOCUMENTOSCOPIA" }
            ,{ "BIOMETRIA FORENSE" }
            ,{ "MERCEOLOGIA" }
            ,{ "BALÍSTICA E CARACTERIZAÇÃO FÍSICA DE MATERIAIS" }
            ,{ "GENÉTICA FORENSE" }
            ,{ "BOMBAS E EXPLOSIVOS" }
            ,{ "MEDICINA E ODONTOLOGIA FORENSE" }
            ,{ "ELETROELETRÔNICOS" }
            ,{ "CARACTERIZAÇÃO FÍSICA DE MATERIAIS" }
            ,{ "CONSTATAÇÃO DE DROGA (PRISÃO EM FLAGRANTE)" }
            ,{ "GEOLOGIA" }
        };

        private static Dictionary<string, string> dict_Secao_de_conclusao_e_Fim_do_preambulo = new Dictionary<string, string>()
        {
             { "RESPOSTA AOS QUESITOS", "respondendo aos quesitos formulados, abaixo transcritos" }
            ,{ "CONCLUSÃO" , "atendendo ao abaixo transcrito" }
        };

        public static List<string> Lista_secao_de_conclusao = new List<string>(dict_Secao_de_conclusao_e_Fim_do_preambulo.Keys);
        public static List<string> Lista_fim_do_preambulo = new List<string>(dict_Secao_de_conclusao_e_Fim_do_preambulo.Values);

        public static Dictionary<string, string> dict_Fim_do_preambulo_e_Secao_de_conclusao = dict_Secao_de_conclusao_e_Fim_do_preambulo.ToDictionary(par => par.Value, par => par.Key); // Dicionario invertido

        public void Metodo_ContentControlOnExit()
        {
            // Configura o evento global para monitorar quando o controle é alterado
            Globals.ThisAddIn.Application.ActiveDocument.ContentControlOnExit += (ContentControl contentControl, ref bool cancel) =>
            {
                VincularLista(contentControl, "Unidade", "Unidade da PF", dict_Unidade_e_Unidade_da_PF);
                if (contentControl.Title == "Unidade")
                {
                    Add_or_remove_ultima_linha_cabecalho1();
                    Muda_Tipo_de_unidade_de_criminalistica();
                }
                VincularLista(contentControl, "Seção de conclusão", "Fim do preâmbulo", dict_Secao_de_conclusao_e_Fim_do_preambulo);
                VincularLista(contentControl, "Fim do preâmbulo", "Seção de conclusão", dict_Fim_do_preambulo_e_Secao_de_conclusao);
            };

        }

        private void VincularLista(ContentControl contentControl_lista1, string titulo_lista1, string titulo_lista2, Dictionary<string, string> dicionario)
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

        public ContentControl GetContentControl(string titulo_do_controle)
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

        public void ChangeEntry(ContentControl ContentControl, string valor_da_lista)
        {
            foreach (ContentControlListEntry entry in ContentControl.DropdownListEntries)
            {
                if (entry.Text == valor_da_lista)
                {
                    if (ContentControl.LockContents)
                    {
                        ContentControl.LockContents = false;
                        ContentControl.DropdownListEntries[entry.Index].Select();
                        ContentControl.LockContents = true;
                    }
                    else
                    {
                        ContentControl.DropdownListEntries[entry.Index].Select();
                    }
                    break;
                }
            }
        }

        public void Add_or_remove_ultima_linha_cabecalho1()
        {
            //if (ContentControl.Title == "Unidade")
            //{
            ContentControl controle_Unidade_da_PF = GetContentControl("Unidade da PF");

            Paragraph paragraph = controle_Unidade_da_PF.Range.Paragraphs[1].Next();
            if (controle_Unidade_da_PF.Range.Text == "DITEC - INSTITUTO NACIONAL DE CRIMINALÍSTICA")
            {
                if (paragraph != null)
                {
                    if (paragraph.Range.ContentControls.Count > 0) paragraph.Range.ContentControls[1].LockContentControl = false;
                    paragraph.Range.Delete();
                }
            }
            else
            {
                string autotextName = "tipo_unidade_crim_PeriTAB";
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
                            if (paragraph.Range.ContentControls.Count > 0) paragraph.Range.ContentControls[1].LockContentControl = false;
                            paragraph.Range.Delete();
                            controle_Unidade_da_PF.Range.Paragraphs[1].Range.InsertParagraphAfter();
                            bb.Insert(controle_Unidade_da_PF.Range.Paragraphs[1].Next().Range);
                        }
                    }
                }
            }
            //}
        }

        public void Muda_Tipo_de_unidade_de_criminalistica()
        {
            ContentControl controle_Unidade = GetContentControl("Unidade");
            ContentControl controle_Tipo_de_unidade_de_criminalistica = GetContentControl("Tipo de unidade de criminalistica");
            if (controle_Unidade.Range.Text.StartsWith("SETEC")) ChangeEntry(controle_Tipo_de_unidade_de_criminalistica, "SETOR TÉCNICO-CIENTÍFICO");
            if (controle_Unidade.Range.Text.StartsWith("NUTEC")) ChangeEntry(controle_Tipo_de_unidade_de_criminalistica, "NÚCLEO TÉCNICO-CIENTÍFICO");
        }

    }
}
