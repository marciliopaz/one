﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace One
{
    public partial class Form1 : Form
    {
        /*
        - Pedir o arquivo do ano anterior ECD 2018
        - Registo I050 (tirar duvidas do preenchimento)
           - revisar a planilha excel... Nivel da conta....tem registros com codigo 5... Pela tabela vai até 4

            Nível	Grupo/Conta
            1	    Ativo
            2	    Ativo Circulante
            3	    Disponível
            4	    Caixa


        - Registro I150 (verificar qual data inicial e data final)
        - Verificar se é necessário preencher os registros I051, I052, I100
        - Solicitar dados do Bloco J
	        J005
	        J100
	        J150
	        J900
	        J990
         */

        string nomeEmpresa = "ONE";
        string cnpj = "00.000.000/0001-00";
        string endereco = "RUA XYZ, 999";

        public Form1()
        {
            InitializeComponent();
        }

        // GERA TXT
        public bool geraTxt(string texto, string caminhoTxt)
        {
            try
            {
                // Check if file already exists. If yes, delete it.     
                if (System.IO.File.Exists(caminhoTxt))
                {
                    System.IO.File.Delete(caminhoTxt);
                }

                // REGISTRO 0
                texto += "";

                // REGISTRO 7

                System.IO.File.WriteAllText(caminhoTxt, texto);

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocorreu um erro ao salvar o Txt \n " + "Detalhes:" + ex.Message);
                return false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string strDados = "";

            //Nome do Arquivo
            //string strPathFile = @"C:\temp\DesenvTeste\One\One\teste.xlsx";
            //string strPathFile = @"C:\temp\DesenvTeste\One\One\Blocos.xlsx";
            string strPathFile = @"C:\temp\DesenvTeste\One\One\Blocos\Blocos2013.xls";
            string strPathTxt = @"C:\temp\DesenvTeste\One\One\WriteText.txt";


            string con =
              @"Provider=Microsoft.Jet.OLEDB.4.0;" +
              @"Data Source=" + strPathFile + "; " +
              @"Extended Properties='Excel 8.0;HDR=Yes;'";

            //con = @"Provider = Microsoft.ACE.OLEDB.12.0; " +
            //    @"Data Source = " + strPathFile + "; " +
            //    @"Extended Properties = Excel 8.0; " +
            //    @"providerName=Provider = Microsoft.ACE.OLEDB.12.0";

            //String de conexão Excel 2003:
            //@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=CAMINHO_DO_XLS;Extended Properties='Excel 8.0;HDR=YES;'"

            //String de conexão Excel 2007:
            //@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=CAMINHO_DO_XLSX;Extended Properties='Excel 12.0 Xml;HDR=YES;'"

            //String de conexão Excel 2013
            //con =   @"Provider=Microsoft.ACE.OLEDB.14.0;" + 
            //        @"Data Source = " + strPathFile + "; " +
            //        @"Extended Properties='Excel 14.0;HDR=YES;IMEX=1'";

            try
            {
                using (OleDbConnection connection = new OleDbConnection(con))
                {
                    connection.Open();

                    strDados += preencheRegistro0000();
                    strDados += preencheRegistro0001();
                    strDados += preencheRegistro0007();
                    strDados += preencheRegistro0990();
                    strDados += preencheRegistoI001();
                    strDados += preencheRegistroI010();
                    strDados += preencheRegistroI030();
                    strDados += preencheRegistroI200(connection);

                    connection.Close();
                }

                if (!string.IsNullOrEmpty(strDados))
                {
                    geraTxt(strDados, strPathTxt);
                    MessageBox.Show("Arquivo Gerado");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocorreu um erro, analise o erro abaixo:\r\n" + ex.Message);
                throw;
            }

        }

        public string preencheBloco1(OleDbConnection arqExcel)
        {
            string strTxt = "";
            string stLimit = "|";

            #region REGISTRO I050   ###################### tirar duvida preenchimento plano de conta
            //01  REG
            strTxt += stLimit;
            strTxt += "I050";

            //02  DT_ALT  Data da Inclusão/Alteração: Representa a data da inclusão/alteração da conta no plano de contas.  ###########
            strTxt += stLimit;
            strTxt += "01012018";

            //03  COD_NAT   ############################
            strTxt += stLimit;
            strTxt += "01";

            //04  IND_CTA (S/A)   ############################
            strTxt += stLimit;
            strTxt += "";

            //05  NIVEL
            strTxt += stLimit;
            strTxt += "";
            
            //06  COD_CTA
            strTxt += stLimit;
            strTxt += "";

            //07  COD_CTA_SUP
            strTxt += stLimit;
            strTxt += "";

            //08  CTA
            strTxt += stLimit;
            strTxt += "";

            strTxt += System.Environment.NewLine;
            #endregion

            #region REGISTRO I051 (fazer?????)
            //01  REG
            strTxt += stLimit;
            strTxt += "I051";

            //02  COD_PLAN_REF

            //03  COD_CCUS

            //04  COD_CTA_REF



            strTxt += System.Environment.NewLine;
            #endregion

            #region REGISTRO I052 (fazer????)
            //01  REG
            strTxt += stLimit;
            strTxt += "I052";

            strTxt += System.Environment.NewLine;
            #endregion REGISTRO I100

            #region REGISTRO I100 (fazer????)

            #endregion

            #region REGISTRO I150  ##################
            //01  REG
            strTxt += stLimit;
            strTxt += "I150";

            //02  DT_INI
            strTxt += stLimit;
            strTxt += "01012018";

            //03  DT_FIN
            strTxt += stLimit;
            strTxt += "31012018";



            strTxt += System.Environment.NewLine;
            #endregion

            #region REGISTRO I200 (pendente) - lancamentos
            //01  REG
            strTxt += stLimit;
            strTxt += "I200";

            //02  NUM_LCTO
            strTxt += stLimit;
            strTxt += "";

            //03  DT_LCTO
            strTxt += stLimit;
            strTxt += "";

            //04  VL_LCTO
            strTxt += stLimit;
            strTxt += "";

            //05  IND_LCTO
            strTxt += stLimit;
            strTxt += "";

            //06  DT_LCTO_EXT
            strTxt += stLimit;
            strTxt += "";


            strTxt += System.Environment.NewLine;
            #endregion

            #region REGISTRO I250 (pendente)
            //01  REG
            //02  COD_CTA
            //03  COD_CCUS
            //04  VL_DC
            //05  IND_DC
            //06  NUM_ARQ
            //07  COD_HIST_PAD
            //08  HIST
            //09  COD_PART


            strTxt += System.Environment.NewLine;
            #endregion

            #region REGISTRO I355 (pendente)

            strTxt += System.Environment.NewLine;
            #endregion

            //ENCERRAMENTO DO BLOCO I   
            #region REGISTRO I990 ##############
            //01  REG
            strTxt += stLimit;
            strTxt += "I990";

            //02  QTD_LIN_I
            strTxt += stLimit;
            strTxt += "4";

            strTxt += System.Environment.NewLine;
            #endregion

            #region REGISTRO J001
            //01  REG
            strTxt += stLimit;
            strTxt += "J001";

            //02  IND_DAD
            strTxt += stLimit;
            strTxt += "0";


            strTxt += System.Environment.NewLine;
            #endregion

            #region REGISTRO J005 (PEGAR DADOS COM A ELLOA)
            //01  REG
            //02  DT_INI
            //03  DT_FIN
            //04  ID_DEM
            //05  CAB_DEM

            strTxt += System.Environment.NewLine;
            #endregion

            #region REGISTRO J100 (PEGAR DADOS COM A ELLOA)
            //01  REG
            //02  COD_AGL
            //03  IND_COD_AGL
            //04  NIVEL_AGL
            //05  COD_AGL_SUP
            //06  IND_GRP_BAL
            //07  DESCR_COD_AGL
            //08  VL_CTA_INI
            //09  IND_DC_CTA_INI
            //10  VL_CTA_FIN
            //11  IND_DC_CTA_FIN
            //12  NOTA_EXP_REF


            strTxt += System.Environment.NewLine;
            #endregion

            #region REGISTRO J150 (PEGAR DADOS COM A ELLOA)
            //01 REG
            strTxt += stLimit;
            strTxt += "J150";

            //02  COD_AGL
            //03  IND_COD_AGL
            //04  NIVEL_AGL
            //05  COD_AGL_SUP
            //06  DESCR_COD_AGL
            //07  VL_CTA
            //08  IND_DC_CTA
            //09  IND_GRP_DRE
            //10  NOTA_EXP_REF




            strTxt += System.Environment.NewLine;
            #endregion

            #region REGISTRO J900 (PEGAR DADOS COM A ELLOA)
            //01    REG
            strTxt += stLimit;
            strTxt += "J900";

            //02  DNRC_ENCER
            strTxt += stLimit;
            strTxt += "TERMO DE ENCERRAMENTO";
            
            //03  NUM_ORD
            //04  NAT_LIVRO
            //05  NOME
            //06  QTD_LIN
            //07  DT_INI_ESCR
            //08  DT_FIN_ESCR

            strTxt += System.Environment.NewLine;
            #endregion

            #region REGISTRO J930 (PEGAR DADOS DA ELLOA)
            //01    REG
            strTxt += stLimit;
            strTxt += "J930";

            //02  IDENT_NOM
            //03  IDENT_CPF_CNPJ
            //04  IDENT_QUALIF
            //05  COD_ASSIN
            //06  IND_CRC
            //07  EMAIL
            //08  FONE
            //09  UF_CRC
            //10  NUM_SEQ_CRC
            //11  DT_CRC
            //12  IND_RESP_LEGAL


            strTxt += System.Environment.NewLine;
            #endregion

            #region REGISTRO J990
            //01    REG
            strTxt += stLimit;
            strTxt += "J990";

            //02  QTD_LIN_J
            strTxt += stLimit;
            strTxt += "4"; ///QUANTIDADE LINHAS DO BLOCO J O bloco J tem um total de 100 linhas

            strTxt += System.Environment.NewLine;
            #endregion


            strTxt += stLimit;


            return strTxt;
        }

        public string preencheRegistroI030()
        {
            string strTxt = "";
            string stLimit = "|";

            #region REGISTRO I030
            //01  REG
            strTxt += stLimit;
            strTxt += "I030";

            //02  DNRC_ABERT
            strTxt += stLimit;
            strTxt += "TERMO DE ABERTURA";

            //03  NUM_ORD
            strTxt += stLimit;
            strTxt += "2";

            //04  NAT_LIVR
            strTxt += stLimit;
            strTxt += "DIARIO GERAL";

            //05  QTD_LIN
            strTxt += stLimit;
            strTxt += "718.719";

            //06  NOME
            strTxt += stLimit;
            strTxt += "OCEAN NETWORK EXPRESS (Latin America)";

            //07  NIRE
            strTxt += stLimit;
            strTxt += "35235086630";

            //08  CNPJ
            strTxt += stLimit;
            strTxt += "28689596000106";

            //09  DT_ARQ
            strTxt += stLimit;
            strTxt += "01012018";

            //10  DT_ARQ_CONV
            strTxt += stLimit;
            strTxt += "31122018";

            //11  DESC_MUN
            strTxt += stLimit;
            strTxt += "SAO PAULO";

            //12  DT_EX_SOCIAL
            strTxt += stLimit;
            strTxt += "31122018";


            strTxt += System.Environment.NewLine;
            #endregion

            return strTxt;
        }

        public string preencheRegistroI010()
        {

            string strTxt = "";
            string stLimit = "|";

            #region REGISTRO I010
            //01  REG
            strTxt += stLimit;
            strTxt += "I010";

            //02  IND_ESC
            strTxt += stLimit;
            strTxt += "G";

            //03  COD_VER_LC
            strTxt += stLimit;
            strTxt += "7.00";

            strTxt += stLimit;

            strTxt += System.Environment.NewLine;
            #endregion

            return strTxt;
        }

        public string preencheRegistoI001()
        {
            string strTxt = "";
            string stLimit = "|";

            #region REGISTRO I001
            //01  REG
            strTxt += stLimit;
            strTxt += "I001";

            //02  IND_DAD
            strTxt += stLimit;
            strTxt += "0";

            strTxt += stLimit;

            strTxt += System.Environment.NewLine;
            #endregion

            return strTxt;
        }

        public string preencheRegistro0990()
        {
            string strTxt = "";
            string stLimit = "|";

            #region REGISTRO 0990
            //01  REG
            strTxt += stLimit;
            strTxt += "0990";

            //02  QTD_LIN_0   #########################
            strTxt += stLimit;
            strTxt += "4";

            strTxt += stLimit;

            strTxt += System.Environment.NewLine;
            #endregion

            return strTxt;
        }

        public string preencheRegistro0007()
        {
            string strTxt = "";
            string stLimit = "|";

            #region REGISTRO 0007
            //01  REG
            strTxt += stLimit;
            strTxt += "0007";

            //02  COD_ENT _REF
            strTxt += stLimit;
            strTxt += "00";

            //03  COD_INSCR
            strTxt += stLimit;
            strTxt += "";

            strTxt += System.Environment.NewLine;
            #endregion

            return strTxt;
        }

        public string preencheRegistro0001()
        {
            string strTxt = "";
            string stLimit = "|";

            #region REGISTRO 0001
            //01    REG
            strTxt += stLimit;
            strTxt += "0001";

            //02    IND_DAD
            strTxt += stLimit;
            strTxt += "0";

            strTxt += stLimit;

            strTxt += System.Environment.NewLine;
            #endregion
            return strTxt;
        }

        public string preencheRegistro0000()
        {
            string strTxt = "";
            string stLimit = "|";

            #region REGISTRO 0000
            //01    REG
            strTxt += stLimit;
            strTxt += "0000";

            //02    LECD
            strTxt += stLimit;
            strTxt += "LECD";

            //03    DT_INI
            strTxt += stLimit;
            strTxt += "01012018";

            //04	DT_FIN
            strTxt += stLimit;
            strTxt += "31122019";

            //05	NOME
            strTxt += stLimit;
            strTxt += "OCEAN NETWORK EXPRESS";

            //06	CNPJ
            strTxt += stLimit;
            strTxt += "OCEAN NETWORK EXPRESS (Latin America)";

            //07	UF
            strTxt += stLimit;
            strTxt += "SP";

            //08	IE
            strTxt += stLimit;
            strTxt += "";

            //09	COD_MUN
            strTxt += stLimit;
            strTxt += "3550308";

            //10	IM
            strTxt += stLimit;
            strTxt += "58353747";

            //11	IND_SIT_ESP
            strTxt += stLimit;
            strTxt += "";

            //12	IND_SIT_INI_PER
            strTxt += stLimit;
            strTxt += "0";

            //13	IND_NIRE
            strTxt += stLimit;
            strTxt += "1";

            //14	IND_FIN_ESC
            strTxt += stLimit;
            strTxt += "0";

            //15	COD_HASH_SUB    ################
            strTxt += stLimit;
            strTxt += "";

            //16	IND_GRANDE_PORTE
            strTxt += stLimit;
            strTxt += "0";

            //17	TIP_ECD
            strTxt += stLimit;
            strTxt += "0";

            //18	COD_SCP
            strTxt += stLimit;
            strTxt += "";

            //19  IDENT_MF
            strTxt += stLimit;
            strTxt += "N";

            //20  IND_ESC_CONS
            strTxt += stLimit;
            strTxt += "N";

            // FINALIZA BLOCO 1
            strTxt += stLimit;

            strTxt += System.Environment.NewLine;
            #endregion

            return strTxt;
        }

        public string preencheRegistroI200(OleDbConnection arqExcel)
        {
            string strRetorno = "";

            // Exemplo Retorno Registo I200
            // | I200 | 1000 | 02052015 | 5000,00 | N ||

            DataTable dataTable = new DataTable();
            OleDbCommand command = new OleDbCommand("SELECT * FROM [I200 e I250Lançamentos$] ", arqExcel);
            OleDbDataAdapter adapter = new OleDbDataAdapter(command);
            adapter.Fill(dataTable);

            var rows = (from p in dataTable.AsEnumerable()
                        where p[6].ToString() != "Centro de Custo"
                        orderby p[0]
                        select new
                       {
                           dataLancamento = p[0],
                           contaDebitoCredito = p[1],
                           flgDebitoCredito = p[2],
                           arquivamento = p[3],
                           vlLancamento = p [4],
                           centroCusto = p[6],
                           historico = p[7],
                           tipoLancamento = p[8]
                       }).ToList();

            int i = 0;
            foreach (var item in rows)
            {
                i++;
                strRetorno += "|I200|" +
                    i.ToString() + "|" +
                    String.Format("{0:dd/MM/yyyy}", item.dataLancamento) + "|" + // Data Lcto
                    String.Format("{0:N}", item.vlLancamento).Replace(".","") + "|" +  // Vl Lcto
                    "N" + "|" +  // Lancamento Normal
                    Environment.NewLine;
            }


            return strRetorno;
        }
    }
}



//class Program
//{
//    static void Main(string[] args)
//    {
//        // Replace path for your file
//        readXLS(@"C:\MyExcelFile.xls"); // or "*.xlsx"
//        Console.ReadKey();
//    }

//    public static void readXLS(string PathToMyExcel)
//    {
//        //Open your template file.
//        Workbook wb = new Workbook(PathToMyExcel);

//        //Get the first worksheet.
//        Worksheet worksheet = wb.Worksheets[0];

//        //Get cells
//        Cells cells = worksheet.Cells;

//        // Get row and column count
//        int rowCount = cells.MaxDataRow;
//        int columnCount = cells.MaxDataColumn;

//        // Current cell value
//        string strCell = "";

//        Console.WriteLine(String.Format("rowCount={0}, columnCount={1}", rowCount, columnCount));

//        for (int row = 0; row <= rowCount; row++) // Numeration starts from 0 to MaxDataRow
//        {
//            for (int column = 0; column <= columnCount; column++)  // Numeration starts from 0 to MaxDataColumn
//            {
//                strCell = "";
//                strCell = Convert.ToString(cells[row, column].Value);
//                if (String.IsNullOrEmpty(strCell))
//                {
//                    continue;
//                }
//                else
//                {
//                    // Do your staff here
//                    Console.WriteLine(strCell);
//                }
//            }
//        }
//    }
//}