using System;
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


        - Registro I150 (verificar qual data inicial e data final) 01/01/2018 31/01/2018
        - Embaixo do registro acima informar o registro I155

        - Registro I150 (verificar qual data inicial e data final) 01/02/2018 28/02/2018
        - Embaixo do registro acima informar o registro I155

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
                    strDados += preencheRegistroI001();
                    strDados += preencheRegistroI010();
                    strDados += preencheRegistroI030();
                    strDados += preencheRegistroI050();

                    strDados += preencheRegistroI155();

                    // preencheRegistroI52();// Falta finalizar
                    // preencheRegistroI100(); // Falta finalizar


                    //strDados += preencheRegistroI200(connection);

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


        public string preencheRegistroI155()
        {
            string strTxt = "";
            string stLimit = "|";

            string strPathFile = @"C:\temp\DesenvTeste\One\One\Blocos\BlocoI150.xls";

            string strCon =
              @"Provider=Microsoft.Jet.OLEDB.4.0;" +
              @"Data Source=" + strPathFile + "; " +
              @"Extended Properties='Excel 8.0;HDR=Yes;'";

            DataTable dataTable = new DataTable();
            using (OleDbConnection connI150file = new OleDbConnection(strCon))
            {
                connI150file.Open();

                OleDbCommand command = new OleDbCommand("SELECT * FROM [I150$] ", connI150file);
                OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                adapter.Fill(dataTable);

                connI150file.Close();
            }

            var rowsI155 = (from p in dataTable.AsEnumerable()
                            select new
                            {
                                dtIniSaldo = p[0],
                                dtFimSaldo = p[1],
                                codConta = p[2],
                                codCentroCusto = p[3],
                                saldoInicial = p[4],
                                sitSaldo = p[5],
                                totalDebito = p[6],
                                totalCredito = p[7],
                                saldoFinal = p[8],
                                sitSaldoFinal = p[9]
                            }
                            ).ToList();

            int i = 0;
            foreach (var item in rowsI155)
            {
                if (!string.IsNullOrEmpty(item.dtIniSaldo.ToString()))
                {
                    Console.WriteLine(
                        item.dtIniSaldo + " " +
                        item.dtFimSaldo + " " +
                        item.codConta + " " +
                        item.codCentroCusto + " " +
                        item.saldoInicial + " "+
                        item.saldoFinal + " "
                        );
                    i++;

                    // cria a linha 150

                    // cria a linha 155
                    strTxt +=
                        criaLinhaI155(item.codConta, item.codCentroCusto, item.saldoInicial, item.sitSaldo, 
                        item.totalDebito, item.totalCredito, item.saldoFinal, item.sitSaldoFinal);
                }
            }


            return strTxt;
        }

        public string preencheRegistroI150(string dtIni, string dtFim)
        {
            string strTxt = "";
            string stLimit = "|";

            //01  REG
            strTxt += stLimit;
            strTxt += "I150";

            //02  DT_INI
            strTxt += stLimit;
            strTxt += dtIni;

            //03  DT_FIN
            strTxt += stLimit;
            strTxt += dtFim;


            strTxt += stLimit;
            strTxt += System.Environment.NewLine;

            return strTxt;
        }

        public String preencheRegistroI100() // Centro de Custo
        {
            string strTxt = "";
            string stLimit = "|";

            //01  REG
            strTxt += stLimit;
            strTxt += "I100";

            //02  DT_ALT
            strTxt += stLimit;
            strTxt += "01012018";

            //03  COD_CCUS
            strTxt += stLimit;
            strTxt += " 11101000";

            //04  CCUS
            strTxt += stLimit;
            strTxt += "BR01SAOZ01";

            strTxt += stLimit;
            strTxt += System.Environment.NewLine;

            return strTxt;
        }

        public string preencheRegistroI050()
        {
            string strTxt = "";
            string stLimit = "|";

            string strPathFile = @"C:\temp\DesenvTeste\One\One\Blocos\BlocoI.xls";

            string strCon =
              @"Provider=Microsoft.Jet.OLEDB.4.0;" +
              @"Data Source=" + strPathFile + "; " +
              @"Extended Properties='Excel 8.0;HDR=Yes;'";

            DataTable dataTable = new DataTable();
            using (OleDbConnection connI050file = new OleDbConnection(strCon))
            {
                connI050file.Open();

                OleDbCommand command = new OleDbCommand("SELECT * FROM [I050$] ", connI050file);
                OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                adapter.Fill(dataTable);

                connI050file.Close();
            }

            var rowsI050 = (from p in dataTable.AsEnumerable()
                            select new
                            {
                                dtInclusao = p[0],
                                codConta = p[1],
                                descrConta = p[2],  
                                indicadorConta = p[3],
                                indicadorNatureza = p[4],
                                nivel  = p[5],
                                codNivelSup = p[6],
                                codPlanoRef = p[7],
                                codConta2  = p[8],
                                codCentroCusto  = p[9],   
                                codCentroCusto2 = p[10],   
                                codAglut = p[11]
                            }
                            ).ToList();

            int i = 0;
            foreach (var item in rowsI050)
            {
                if (! string.IsNullOrEmpty(item.codConta.ToString()) )
                {
                    Console.WriteLine(
                        item.dtInclusao + " " +
                        item.codConta + " " +
                        item.descrConta + " " +
                        item.indicadorConta + " "
                        );
                    i++;
                }
            }

            #region Trazer NIVEL 1
            var rowNivel1 = (from n1 in rowsI050
                             where n1.nivel.ToString() == "1"
                             select n1).ToList();

            foreach (var item in rowNivel1)
            {
                strTxt += criaLinhaI050( item.dtInclusao
                                , item.indicadorNatureza
                                , item.indicadorConta
                                , item.nivel
                                , item.codConta
                                , ""
                                , item.descrConta
                             );
            }

            #endregion

            #region Trazer NIVEL 2
            var rowNivel2 = (from n2 in rowsI050
                             where n2.nivel.ToString() == "2"
                             select n2).ToList();

            foreach (var item in rowNivel2)
            {
                strTxt += criaLinhaI050(item.dtInclusao
                                , item.indicadorNatureza
                                , item.indicadorConta
                                , item.nivel
                                , item.codConta
                                , item.codNivelSup
                                , item.descrConta
                             );
            }
            #endregion

            #region Trazer NIVEL 3
            var rowNivel3 = (from n3 in rowsI050
                             where n3.nivel.ToString() == "3"
                             select n3).ToList();

            foreach (var item in rowNivel3)
            {
                strTxt += criaLinhaI050(item.dtInclusao
                                , item.indicadorNatureza
                                , item.indicadorConta
                                , item.nivel
                                , item.codConta
                                , item.codNivelSup
                                , item.descrConta
                             );
            }
            #endregion

            #region Trazer NIVEL 4
            var rowNivel4 = (from n4 in rowsI050
                             where n4.nivel.ToString() == "4"
                             select n4).ToList();

            foreach (var item in rowNivel4)
            {
                strTxt += criaLinhaI050(item.dtInclusao
                                , item.indicadorNatureza
                                , item.indicadorConta
                                , item.nivel
                                , item.codConta
                                , item.codNivelSup
                                , item.descrConta
                             );
            }
            #endregion

            #region Trazer NIVEL 5
            var rowNivel5 = (from n5 in rowsI050
                             where n5.nivel.ToString() == "5"
                             select n5).ToList();

            foreach (var item in rowNivel5)
            {
                // Preenche registro I050 nivel 5
                strTxt += criaLinhaI050(item.dtInclusao
                                , item.indicadorNatureza
                                , item.indicadorConta
                                , item.nivel
                                , item.codConta
                                , item.codNivelSup
                                , item.descrConta
                             );
                // Preenche registro I052
                //01  REG
                strTxt += stLimit;
                strTxt += "I052";

                //02  COD_CCUS
                strTxt += stLimit;
                strTxt +=  item.codCentroCusto2.ToString().Trim();

                //03 COD_AGL
                strTxt += stLimit;
                strTxt += item.codAglut.ToString().ToUpper().Trim();

                strTxt += stLimit;
                strTxt += System.Environment.NewLine;
            }

            #endregion

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
            strTxt += "G - DIARIO GERAL";

            //05  QTD_LIN
            strTxt += stLimit;
            strTxt += "718719";

            //06  NOME
            strTxt += stLimit;
            strTxt += "OCEAN NETWORK EXPRESS (LATIN AMERICA)";

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

            strTxt += stLimit;

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

        public string preencheRegistroI001()
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
            strTxt += "SP";

            //03  COD_INSCR
            strTxt += stLimit;
            strTxt += "112066369116";

            // FINALIZA REGISTRO 0007
            strTxt += stLimit;

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
            strTxt += "31122018";

            //05	NOME
            strTxt += stLimit;
            strTxt += "OCEAN NETWORK EXPRESS (LATIN AMERICA)";

            //06	CNPJ
            strTxt += stLimit;
            strTxt += "28689596000106";

            //07	UF
            strTxt += stLimit;
            strTxt += "SP";

            //08	IE
            strTxt += stLimit;
            strTxt += "112066369116";

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

            var rows = (     from p in dataTable.AsEnumerable()
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
                    string.Format("{0:dd/MM/yyyy}", item.dataLancamento) + "|" + // Data Lcto
                    string.Format("{0:N}", item.vlLancamento).Replace(".","") + "|" +  // Vl Lcto
                    "N" + "|" +  // Lancamento Normal
                    Environment.NewLine;
            }


            return strRetorno;
        }

        public string criaLinhaI050( object dtInclusao
            , object indicadorNatureza
            , object indicadorConta
            , object nivel 
            , object codConta
            , object codNivelSup
            , object descrConta)
        {
            string strTxt = "";
            string stLimit = "|";

            //01  REG
            strTxt += stLimit;
            strTxt += "I050";

            //02  DT_ALT #########################
            strTxt += stLimit;
            strTxt += dtInclusao.ToString().Replace("/", "").Replace("00:00:00", "").Trim();

            //03  COD_NAT (Indicador de Natureza)
            strTxt += stLimit;
            strTxt += indicadorNatureza.ToString();

            //04  IND_CTA (Analitico / Sintetico (A/S))
            strTxt += stLimit;
            strTxt += indicadorConta.ToString().Substring(0, 1);

            //05  NIVEL
            strTxt += stLimit;
            strTxt += nivel.ToString();

            //06  COD_CTA
            strTxt += stLimit;
            strTxt += codConta.ToString();

            //07  COD_CTA_SUP (Codigo Nivel superior)
            strTxt += stLimit;
            strTxt += codNivelSup.ToString();

            //08  CTA (Descricao da Conta)
            strTxt += stLimit;
            strTxt += descrConta.ToString().ToUpper().Trim();

            strTxt += stLimit;

            strTxt += System.Environment.NewLine;

            return strTxt;
        }

        public string criaLinhaI155(object codConta
              , object codCentroCusto
              , object saldoInicial
              , object sitSaldo
              , object totalDebito
              , object totalCredito
              , object saldoFinal
              , object sitSaldoFinal
            )

        {
            string strTxt = "";
            string stLimit = "|";

            #region RegistroI155
            //01  REG
            strTxt += stLimit;
            strTxt += "I155";

            //02  COD_CTA
            strTxt += stLimit;
            strTxt += codConta;

            //03  COD_CCUS
            strTxt += stLimit;
            strTxt += codCentroCusto;

            //04  VL_SLD_INI
            strTxt += stLimit;
            strTxt += string.Format("{0:N}", saldoInicial).Replace(".", "").Replace("-","");

            //05  IND_DC_INI
            strTxt += stLimit;
            strTxt += sitSaldo;

            //06  VL_DEB
            strTxt += stLimit;
            strTxt += string.Format("{0:N}", totalDebito).Replace(".", "").Replace("-", "");

            //07  VL_CRED
            strTxt += stLimit;
            strTxt += string.Format("{0:N}", totalCredito).Replace(".", "").Replace("-", "");

            //08  VL_SLD_FIN
            strTxt += stLimit;
            strTxt += string.Format("{0:N}", saldoFinal).Replace(".", "").Replace("-", "");

            //09  IND_DC_FIN
            strTxt += stLimit;
            strTxt += sitSaldoFinal;


            strTxt += stLimit;
            strTxt += System.Environment.NewLine;
            #endregion

            return strTxt;
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