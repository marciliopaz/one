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
using ExcelDataReader;
using System.IO;

namespace OneSolution
{
    public partial class Form1 : Form
    {
        private string strDir;
        private int cont0000 = 0;
        private int cont0001 = 0;
        private int cont0007 = 0;
        private int cont0990 = 0;
        private int contI001 = 0;
        private int contI010 = 0;
        private int contI030 = 0;
        private int contI050 = 0;
        private int contI051 = 0;
        private int contI052 = 0;
        private int contI100 = 0;
        private int contI150 = 0;
        private int contI155 = 0;
        private int contI200 = 0;
        private int contI250 = 0;
        private int contI350 = 0;
        private int contI355 = 0;
        private int contI990 = 0;
        private int contJ001 = 0;
        private int contJ005 = 0;
        private int contJ100 = 0;
        private int contJ900 = 0;
        private int contJ930 = 0;
        private int contJ990 = 0;
        private int cont9001 = 0;
        private int cont9990 = 0;
        private int cont9999 = 0;
        private int cont9900 = 0;


        /*
        - Pedir o arquivo do ano anterior ECD 2018
        - Registo I050 (tirar duvidas do preenchimento)
        - Registro I150 (verificar qual data inicial e data final)
        - Verificar se é necessário preencher os registros I051, I052
        - Enviar dados do Bloco J
        J005
        J100
        J150
        J900
        J990
        */

        public Form1()
        {
            InitializeComponent();
            //strDir = @"C:\Users\marcilio\source\repos\OneSolution\arquivos\";
            strDir = @"C:\temp\DesenvTeste\One\One\Blocos\";

        }

        private void button1_Click(object sender, EventArgs e)
        {

            //Excel.Application excel = new Excel.Application();
            //Excel.Workbook wb = excel.Workbooks.Open(openFileDialog1.FileName);
            //MessageBox.Show(openFileDialog1.FileName);
            DataTable dtTablesList = default(DataTable);
            string sConnection = null;
            string sSheetName = null;
            OleDbConnection oleExcelConnection = default(OleDbConnection);

            sConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\temp\\Blocos novo.xlsx;Extended Properties=\"Excel 12.0;HDR=No;IMEX=1\"";

            oleExcelConnection = new OleDbConnection(sConnection);
            oleExcelConnection.Open();

            dtTablesList = oleExcelConnection.GetSchema("Tables");

            if (dtTablesList.Rows.Count > 0)
            {
                sSheetName = dtTablesList.Rows[0]["TABLE_NAME"].ToString();
            }
            foreach (DataRow item in dtTablesList.Rows)
            {
                MessageBox.Show(item["TABLE_NAME"].ToString());
            }

            dtTablesList.Clear();
            dtTablesList.Dispose();

        }

        private void button2_Click(object sender, EventArgs e)
        {

            // Excel.Workbook wb = xl.Workbooks.Open("c:\\temp\\bloco.xlsx");
            FileStream stream = File.Open(textBox1.Text, FileMode.Open, FileAccess.Read);

            var extension = Path.GetExtension(textBox1.Text).ToLower();

            IExcelDataReader excelReader = null;

            if (extension == ".xls")
            {
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else if (extension == ".xlsx")
            {
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }
            if (excelReader != null)
            {
                // DataSet result = excelReader.AsDataSet();
                DataTable dt = excelReader.AsDataSet().Tables[0];


            }

            // Exemplo de uso
            //result.Tables[0].Rows[1][3]
            //"Valor do Lançamento" 

        }

        // GERA TXT
        public bool geraTxt(StringBuilder texto, string caminhoTxt)
        {
            try
            {
                // Check if file already exists. If yes, delete it.     
                if (System.IO.File.Exists(caminhoTxt))
                {
                    System.IO.File.Delete(caminhoTxt);
                }

                System.IO.File.WriteAllText(caminhoTxt, texto.ToString());

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocorreu um erro ao salvar o Txt \n " + "Detalhes:" + ex.Message);
                return false;
            }
        }
        public string PreencheBloco1(OleDbConnection arqExcel)
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

            #region REGISTRO 0001
            //01    REG
            strTxt += stLimit;
            strTxt += "0001";

            //02    IND_DAD
            strTxt += stLimit;
            strTxt += "0";

            strTxt += System.Environment.NewLine;
            #endregion

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

            #region REGISTRO 0990  ### verificar qtde linha
            //01  REG
            strTxt += stLimit;
            strTxt += "0990";

            //02  QTD_LIN_0   #########################
            strTxt += stLimit;
            strTxt += "7";

            strTxt += System.Environment.NewLine;
            #endregion

            #region REGISTRO I001
            //01  REG
            strTxt += stLimit;
            strTxt += "I001";

            //02  IND_DAD
            strTxt += stLimit;
            strTxt += "0";
            #endregion

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


            strTxt += System.Environment.NewLine;
            #endregion

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

            #region REGISTRO I050   ######################
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

            #region REGISTRO I200 (pendente)
            //01  REG
            strTxt += stLimit;
            strTxt += "I200";

            //02  NUM_LCTO

            //03  DT_LCTO

            //04  VL_LCTO

            //05  IND_LCTO

            //06  DT_LCTO_EXT


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

        private void button3_Click(object sender, EventArgs e)
        {
            string strDados = "";
            StringBuilder stDados = new StringBuilder();

            //Nome do Arquivo
            //string strPathFile = @"C:\temp\DesenvTeste\One\One\teste.xlsx";
            //string strPathFile = @"C:\temp\DesenvTeste\One\One\Blocos.xlsx";
            //string strPathFile = @"C:\temp\DesenvTeste\One\One\Blocos\Blocos2013.xls";
            string strPathFile = strDir + @"Bloco2013.xls";
            string strPathTxt = strDir + @"WriteText.txt";


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
                    progressBar1.Value += 10;

                    connection.Open();

                    //strDados += preencheRegistro0000();
                    stDados.Append(preencheRegistro0000());
                    progressBar1.Value += 10;
                    //strDados += preencheRegistro0001();
                    stDados.Append(preencheRegistro0001());
                    progressBar1.Value += 10;
                    //strDados += preencheRegistro0007();
                    stDados.Append(preencheRegistro0007());
                    progressBar1.Value += 10;
                    //strDados += preencheRegistro0990();
                    stDados.Append(preencheRegistro0990());
                    progressBar1.Value += 10;
                    //strDados += preencheRegistroI001();
                    stDados.Append(preencheRegistroI001());
                    progressBar1.Value += 10;
                    //strDados += preencheRegistroI010();
                    stDados.Append(preencheRegistroI010());
                    progressBar1.Value += 10;
                    //strDados += preencheRegistroI030();
                    stDados.Append(preencheRegistroI030());
                    progressBar1.Value += 10;
                    //strDados += preencheRegistroI050();
                    stDados.Append(preencheRegistroI050());
                    progressBar1.Value += 10;
                    //strDados += preencheRegistroI100();
                    //strDados += preencheRegistroI155();
                    stDados.Append(preencheRegistroI100());
                    stDados.Append(preencheRegistroI155());

                    progressBar1.Value += 10;
                    //strDados += preencheLancamentos();
                    //strDados += preencheRegistroI350eI355();
                    //strDados += preencheRegistroI990();
                    stDados.Append(preencheLancamentos());
                    stDados.Append(preencheRegistroI350eI355());
                    stDados.Append(preencheRegistroI990());

                    //strDados += preencheRegistroJ001();
                    //strDados += preencheRegistroJ005();
                    //strDados += preencheRegistroJ900();//falta preencher pedir dados
                    stDados.Append(preencheRegistroJ990());

                    // Contadores
                    stDados.Append(preencheRegistro9001());
                    stDados.Append(preencheRegistro9900());
                    stDados.Append(preencheRegistro9990());
                    stDados.Append(preencheRegistro9999());




                    //strDados += preencheRegistroI200(connection);
                    connection.Close();
                }

                //if (!string.IsNullOrEmpty(strDados))
                if (stDados.Length > 0)
                {
                    geraTxt(stDados, strPathTxt);
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

        public StringBuilder preencheRegistroI155()
        {
            //string strTxt = "";
            StringBuilder stTxt = new StringBuilder();
            string strPathFile = strDir + @"BlocoI150.xls";
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
                            )
                            .ToList();


            string ctrlPeriodo = "";
            foreach (var item in rowsI155)
            {
                if (!string.IsNullOrEmpty(item.dtIniSaldo.ToString()))
                {
                    if (item.dtIniSaldo.ToString() != ctrlPeriodo)
                    {
                        ctrlPeriodo = item.dtIniSaldo.ToString();
                        // cria a linha 150
                        stTxt.Append(preencheRegistroI150(item.dtIniSaldo.ToString(), item.dtFimSaldo.ToString()));
                        contI150++;
                    }
                    else
                    {
                        ctrlPeriodo = item.dtIniSaldo.ToString();
                        // cria a linha 155
                        stTxt.Append(criaLinhaI155(item.codConta, item.codCentroCusto, item.saldoInicial, item.sitSaldo,
                            item.totalDebito, item.totalCredito, item.saldoFinal, item.sitSaldoFinal));
                        contI155++;
                    }
                }
            }
            rowsI155 = null;

            return stTxt;
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

            //03  COD_CCUS (nao)
            strTxt += stLimit;
            strTxt += codCentroCusto;

            //04  VL_SLD_INI
            strTxt += stLimit;
            double saldoIni = 0.00;
            double.TryParse(saldoInicial.ToString().Replace("-", ""), out saldoIni);
            strTxt += string.Format("{0:N}", saldoIni);

            //05  IND_DC_INI (nao)
            strTxt += stLimit;
            strTxt += sitSaldo;

            //06  VL_DEB
            strTxt += stLimit;
            double vlDeb = 0.00;
            double.TryParse(totalDebito.ToString().Replace("-", ""), out vlDeb);
            strTxt += string.Format("{0:N}", vlDeb);

            //07  VL_CRED
            strTxt += stLimit;
            double vlCred = 0.00;
            double.TryParse(totalCredito.ToString().Replace("-", ""), out vlCred);
            strTxt += string.Format("{0:N}", vlCred);

            //08  VL_SLD_FIN
            strTxt += stLimit;
            double vlSldFin = 0.00;
            double.TryParse(saldoFinal.ToString().Replace("-", ""), out vlSldFin);
            strTxt += string.Format("{0:N}", vlSldFin);

            //09  IND_DC_FIN (nao)
            strTxt += stLimit;
            strTxt += sitSaldoFinal;


            strTxt += stLimit;
            strTxt += System.Environment.NewLine;
            #endregion

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
            strTxt += dtIni.ToString().Replace("/", "").Replace("00:00:00", "").Trim();

            //03  DT_FIN
            strTxt += stLimit;
            strTxt += dtFim.ToString().Replace("/", "").Replace("00:00:00", "").Trim();


            strTxt += stLimit;
            strTxt += System.Environment.NewLine;

            return strTxt;
        }

        public string preencheRegistroI100() // Centro de Custo
        {
            string strTxt = "";

            string strPathFile = strDir + @"Bloco I100.xlsx";

            using (FileStream stream = File.Open(strPathFile, FileMode.Open, FileAccess.Read))
            {
                var extension = Path.GetExtension(strPathFile).ToLower();

                IExcelDataReader excelReader = null;

                if (extension == ".xls")
                {
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else if (extension == ".xlsx")
                {
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }
                if (excelReader != null)
                {
                    // 1. Opção
                    // Abre a tabela informando que a primeira linha possui o nome do campo
                    DataTable dt = excelReader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                    })
                    .Tables[0];
                    //2. Opção
                    // DataTable dt = excelReader.AsDataSet().Tables[0];

                    stream.Close();
                    // Possui Dados no Excel
                    foreach (DataRow r in dt.Rows)
                    {
                        if (r[0].ToString().Trim().Length > 0) { 
                            strTxt += criaLinhaI100(r[0].ToString(), r[1].ToString(), r[2].ToString());
                            contI100++;
                        }
                    }
                }

            }
            return strTxt;
        }

        public string criaLinhaI100(string dtAlt, string strCodCusto, string strDescrCentroCusto)
        {
            string strTxt = "";
            string stLimit = "|";

            //01  REG
            strTxt += stLimit;
            strTxt += "I100";

            //02  DT_ALT
            strTxt += stLimit;
            //strTxt += "20092017";
            strTxt += dtAlt.ToString().Replace("/", "").Replace("00:00:00", "").Trim();

            //03  COD_CCUS
            strTxt += stLimit;
            //strTxt += " 11101000";
            strTxt += strCodCusto;

            //04  CCUS
            strTxt += stLimit;
            //strTxt += "BR01SAOZ01";
            strTxt += strDescrCentroCusto;

            strTxt += stLimit;
            strTxt += System.Environment.NewLine;

            return strTxt;
        }

        public string criaLinhaI051(object codPlanoRef, object codCentroCusto, object codConta2)
        {
            string strTxt = "";
            string stLimit = "|";

            //01  REG
            strTxt += stLimit;
            strTxt += "I051";

            //02  COD_PLAN_REF
            strTxt += stLimit;
            strTxt += codPlanoRef.ToString().Substring(0, 1);

            //03  COD_CCUS
            strTxt += stLimit;
            strTxt += codCentroCusto.ToString();

            //04  COD_CTA_REF
            strTxt += stLimit;
            strTxt += codConta2.ToString();

            strTxt += stLimit;

            strTxt += System.Environment.NewLine;

            return strTxt;
        }

        public string criaLinhaI050(object dtInclusao
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
            strTxt += indicadorNatureza.ToString().PadLeft(2, '0');

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

        public string preencheRegistroI050()
        {
            string strTxt = "";
            string stLimit = "|";

            string strPathFile = strDir + @"BlocoI.xls";

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
                                nivel = p[5],
                                codNivelSup = p[6],
                                codPlanoRef = p[7],
                                codConta2 = p[8],
                                codCentroCusto = p[9],
                                codCentroCusto2 = p[10],
                                codAglut = p[11]
                            }
                            ).ToList();

            #region Trazer NIVEL 1
            var rowNivel1 = (from n1 in rowsI050
                             where n1.nivel.ToString() == "1"
                             select n1).ToList();

            foreach (var item in rowNivel1)
            {
                strTxt += criaLinhaI050(item.dtInclusao
                                , item.indicadorNatureza
                                , item.indicadorConta
                                , item.nivel
                                , item.codConta
                                , ""
                                , item.descrConta
                             );
                contI050++;

                strTxt += criaLinhaI051(item.codPlanoRef, item.codCentroCusto, item.codConta2);
                contI051++;
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
                contI050++;

                strTxt += criaLinhaI051(item.codPlanoRef, item.codCentroCusto, item.codConta2);
                contI051++;
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
                contI050++;

                strTxt += criaLinhaI051(item.codPlanoRef, item.codCentroCusto, item.codConta2);
                contI051++;
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
                contI050++;

                strTxt += criaLinhaI051(item.codPlanoRef, item.codCentroCusto, item.codConta2);
                contI051++;
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
                contI050++;

                strTxt += criaLinhaI051(item.codPlanoRef, item.codCentroCusto, item.codConta2);
                contI051++;

                // Preenche registro I052
                //01  REG
                strTxt += stLimit;
                strTxt += "I052";

                //02  COD_CCUS
                strTxt += stLimit;
                strTxt += item.codCentroCusto2.ToString().Trim();

                //03 COD_AGL
                strTxt += stLimit;
                strTxt += item.codAglut.ToString().ToUpper().Trim();

                strTxt += stLimit;
                strTxt += System.Environment.NewLine;
                contI052++;
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
            //strTxt += "OCEAN NETWORK EXPRESS (LATIN AMERICA)";
            strTxt += "Ocean Network Express(Latin America) Agencia Maritima LTDA";

            //07  NIRE
            strTxt += stLimit;
            strTxt += "35235086630";

            //08  CNPJ
            strTxt += stLimit;
            strTxt += "28689596000106";

            //09  DT_ARQ
            strTxt += stLimit;
            strTxt += "20092017";

            //10  DT_ARQ_CONV
            strTxt += stLimit;
            strTxt += "";

            //11  DESC_MUN
            strTxt += stLimit;
            strTxt += "SAO PAULO";

            //12  DT_EX_SOCIAL (DATA ENCERRAMENTO EXERCICIO SOCIAL)
            strTxt += stLimit;
            strTxt += "31122018";

            strTxt += stLimit;

            strTxt += System.Environment.NewLine;
            contI030++;
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
            contI010++;
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
            contI001++;
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
            cont0990++;
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
            cont0007++;
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
            cont0001++;
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
            //strTxt += "OCEAN NETWORK EXPRESS (LATIN AMERICA)";
            strTxt += "Ocean Network Express(Latin America) Agencia Maritima LTDA";

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
            strTxt += "1";

            //15	COD_HASH_SUB   
            strTxt += stLimit;
            strTxt += "4A.19.5B.1A.92.CC.13.EC.46.C3.74.D9.3C.BE.C9.AC.2E.04.7D.64".Replace(".", "");

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
            cont0000++;
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
                            vlLancamento = p[4],
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
                    string.Format("{0:N}", item.vlLancamento).Replace(".", "") + "|" +  // Vl Lcto
                    "N" + "|" +  // Lancamento Normal
                    Environment.NewLine;
            }


            return strRetorno;
        }

        #region I200 e I250
        public StringBuilder preencheLancamentos()
        {
            //string strTxt = "";
            StringBuilder stTxtGeral = new StringBuilder();

            string strPathFile = strDir + @"Bloco I200 2250 novom.xlsx";

            using (FileStream stream = File.Open(strPathFile, FileMode.Open, FileAccess.Read))
            {
                var extension = Path.GetExtension(strPathFile).ToLower();

                IExcelDataReader excelReader = null;

                if (extension == ".xls")
                {
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else if (extension == ".xlsx")
                {
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }
                if (excelReader != null)
                {
                    // 1. Opção
                    // Abre a tabela informando que a primeira linha possui o nome do campo
                    DataTable dt = excelReader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                    })
                    .Tables[0];
                    //2. Opção
                    // DataTable dt = excelReader.AsDataSet().Tables[0];

                    stream.Close();

                    //  Exemplo rapido
                    //  Referencia
                    //  https://chrisbitting.com/2016/05/04/datatable-row-loop-csharp-performance-testing-linq-vs-select-vs-parallel-vs-for/
                    int i = 0;
                    string i200 = "";

                    var query = from DataRow row in dt.AsEnumerable()
                                select row;
                    //Parallel.For(0, dt.Rows.Count, r =>
                    // Parallel.ForEach(query, r =>
                    foreach (var r in query)
                    {
                        if (r[0].ToString().Trim() != "2")
                        {
                            // Cria Registro I200
                            if (i200 != r[1].ToString() + r[2].ToString() || i == 2)
                            {
                                stTxtGeral.Append(criaLinhaI200(r[1].ToString(), r[2].ToString(), r[3].ToString(), r[4].ToString()));
                                contI200++;
                            }

                            i200 = r[1].ToString() + r[2].ToString();

                            // Cria Registro I250
                            stTxtGeral.Append(criaLinhaI250(r[5].ToString(), r[6].ToString(), r[7].ToString(), r[8].ToString(), r[10].ToString()));
                            contI250++;
                        }
                    }
                    //});

                    //var query = from DataRow row in dt.AsEnumerable()
                    //            select row;

                    //Parallel.ForEach(query, row =>
                    //{
                    //    string sval = row["Field2"].ToString();
                    //    stTxt.Append(sval);
                    //});


                    // Possui Dados no Excel
                    //int i = 0;
                    //string i200 = "";
                    //foreach (DataRow r in dt.Rows)
                    //{
                    //    i++;
                    //    textBox1.Text = i.ToString();
                    //    if (r[0].ToString().Trim() != "2")
                    //    {
                    //        // Cria Registro I200
                    //        if (i200 != r[1].ToString() + r[2].ToString() || i == 2)
                    //        {
                    //            stTxt.Append(criaLinhaI200(r[1].ToString(), r[2].ToString(), r[3].ToString(), r[4].ToString()));
                    //            //strTxt += criaLinhaI200(r[1].ToString(), r[2].ToString(), r[3].ToString(), r[4].ToString());
                    //        }
                    //        i200 = r[1].ToString() + r[2].ToString();

                    //        // Cria Registro I250
                    //        stTxt.Append(criaLinhaI250(r[5].ToString(), r[6].ToString(), r[7].ToString(), r[8].ToString(), r[10].ToString()));
                    //        //strTxt += criaLinhaI250(r[5].ToString(), r[6].ToString(), r[7].ToString(), r[8].ToString(), r[10].ToString());

                    //    }
                    //    if (i == 110)
                    //    {
                    //        break;
                    //    }
                    //}
                }

            }
            return stTxtGeral;
        }

        private StringBuilder criaLinhaI200(string dtLcto, string numLcto, string vlLcto, string tipoLcto)
        {
            StringBuilder stTxt = new StringBuilder();
            string stLimit = "|";

            //01  REG
            stTxt.Append(stLimit);
            stTxt.Append("I200");

            //02  NUM_LCTO
            stTxt.Append(stLimit);
            stTxt.Append(numLcto.Replace("'", ""));

            //03  DT_LCTO
            stTxt.Append(stLimit);
            stTxt.Append(dtLcto.ToString().Replace("/", "").Replace("00:00:00", "").Trim());

            //04  VL_LCTO
            double tmpVlLcto = 0.00;
            double.TryParse(vlLcto, out tmpVlLcto);
            stTxt.Append(stLimit);
            stTxt.Append(tmpVlLcto.ToString("F"));

            //05  IND_LCTO
            if (tipoLcto.Trim() == "Normal")
            {
                tipoLcto = "N";
            }
            else
            {
                tipoLcto = tipoLcto.Substring(0, 1);
            }
            stTxt.Append(stLimit);
            stTxt.Append(tipoLcto);

            //06  DT_LCTO_EXT
            stTxt.Append(stLimit);
            stTxt.Append("");

            stTxt.Append(stLimit);

            stTxt.Append(System.Environment.NewLine);

            return stTxt;
        }

        private StringBuilder criaLinhaI250(string codCta, string centroCusto, string valorDc, string indDC, string hist)
        {
            //string strTxt = "";
            StringBuilder stTxt = new StringBuilder();
            string stLimit = "|";

            //01  REG
            stTxt.Append(stLimit);
            stTxt.Append("I250");

            //02  COD_CTA
            stTxt.Append(stLimit);
            stTxt.Append(codCta.Replace("'", ""));

            //03  COD_CCUS
            stTxt.Append(stLimit);
            stTxt.Append(centroCusto);

            //04  VL_DC
            stTxt.Append(stLimit);
            stTxt.Append(Convert.ToDouble(valorDc).ToString("F").Replace("-", ""));

            //05  IND_DC
            stTxt.Append(stLimit);
            stTxt.Append(indDC);

            //06  NUM_ARQ
            stTxt.Append(stLimit);
            stTxt.Append("");

            //07  COD_HIST_PAD
            stTxt.Append(stLimit);
            stTxt.Append(centroCusto);

            //08  HIST
            stTxt.Append(stLimit);
            stTxt.Append(hist.Replace("\r\n", string.Empty).Replace("\n", string.Empty).Replace("\r", string.Empty).Replace(Environment.NewLine, "").Replace("|",""));

            //09  COD_PART
            stTxt.Append(stLimit);
            stTxt.Append("");

            stTxt.Append(stLimit);

            stTxt.Append(System.Environment.NewLine);

            return stTxt;

        }

        #endregion

        #region I350 e I350
        public string preencheRegistroI350eI355()
        {
            string strTxt = "";

            string strPathFile = strDir + @"Bloco I350.xls";

            using (FileStream stream = File.Open(strPathFile, FileMode.Open, FileAccess.Read))
            {
                var extension = Path.GetExtension(strPathFile).ToLower();

                IExcelDataReader excelReader = null;

                if (extension == ".xls")
                {
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else if (extension == ".xlsx")
                {
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }
                if (excelReader != null)
                {
                    // 1. Opção
                    // Abre a tabela informando que a primeira linha possui o nome do campo
                    DataTable dt = excelReader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                    })
                    .Tables[0];
                    //2. Opção
                    // DataTable dt = excelReader.AsDataSet().Tables[0];

                    stream.Close();

                    // Possui Dados no Excel
                    foreach (DataRow r in dt.Rows)
                    {
                        if (r[0].ToString().Trim() != "Data")
                        {
                            strTxt += criaLinhaI350(r[0].ToString());
                            strTxt += criaLinhaI355(r[1].ToString(), r[2].ToString(), r[3].ToString(), r[4].ToString());
                            contI350++;
                            contI355++;
                        }
                    }
                }

            }
            return strTxt;
        }

        public string criaLinhaI350(string dtRes)
        {
            string strTxt = "";
            string stLimit = "|";

            //01    REG
            strTxt += stLimit;
            strTxt += "I350";

            //02	DT_RES
            strTxt += stLimit;
            strTxt += dtRes.ToString().Replace("/", "").Replace("00:00:00", "").Trim();


            strTxt += stLimit;

            strTxt += System.Environment.NewLine;

            return strTxt;

        }

        public string criaLinhaI355(string codConta, string codCentroCusto, string vlCta, string indDC)
        {
            string strTxt = "";
            string stLimit = "|";

            //01	REG
            strTxt += stLimit;
            strTxt += "I355";

            //02	COD_CTA
            strTxt += stLimit;
            strTxt += codConta;

            //03	COD_CCUS
            strTxt += stLimit;
            strTxt += codCentroCusto;

            //04	VL_CTA
            strTxt += stLimit;
            strTxt += Convert.ToDouble(vlCta).ToString("F");

            //05	IND_DC
            strTxt += stLimit;
            strTxt += indDC;

            strTxt += stLimit;

            strTxt += System.Environment.NewLine;

            return strTxt;
        }
        #endregion

        #region I990 (IMPORTANTE)

        public string preencheRegistroI990()
        {
            string strTxt = "";
            string stLimit = "|";

            contI990++;

            //01    REG
            strTxt += stLimit;
            strTxt += "I990";

            //02	QTD_LIN_i
            strTxt += stLimit;
            //strTxt += "734160";
            strTxt += (contI001 
                + contI010
                + contI030 
                + contI050
                + contI051
                + contI052
                + contI100 
                + contI150 + contI155 + contI200 + contI250 + contI350 + contI355 + contI990).ToString();

            strTxt += stLimit;

            strTxt += System.Environment.NewLine;

            return strTxt;
        }
        #endregion

        #region J001
        public string preencheRegistroJ001()
        {
            string strTxt = "";
            string stLimit = "|";

            //01    REG
            strTxt += stLimit;
            strTxt += "J001";

            //02	IND_DAD
            strTxt += stLimit;
            strTxt += "0";

            strTxt += stLimit;

            strTxt += System.Environment.NewLine;

            return strTxt;
        }
        #endregion

        #region J005
        public string preencheRegistroJ005()
        {
            string strTxt = "";

            string strPathFile = strDir + @"J005.xlsx";

            using (FileStream stream = File.Open(strPathFile, FileMode.Open, FileAccess.Read))
            {
                var extension = Path.GetExtension(strPathFile).ToLower();

                IExcelDataReader excelReader = null;

                if (extension == ".xls")
                {
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else if (extension == ".xlsx")
                {
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }
                if (excelReader != null)
                {
                    // 1. Opção
                    // Abre a tabela informando que a primeira linha possui o nome do campo
                    DataTable dt = excelReader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                    })
                    .Tables[0];

                    //  Exemplo rapido
                    //  Referencia
                    //  https://chrisbitting.com/2016/05/04/datatable-row-loop-csharp-performance-testing-linq-vs-select-vs-parallel-vs-for/

                    string sval = "";
                    StringBuilder sb = new StringBuilder();
                    Parallel.For(0, dt.Rows.Count, rowi =>
                    {
                        sval = rowi.ToString() + "|" + dt.Rows[rowi][4];
                        sb.Append(sval);
                    });

                    //fim exemplo rapido

                    //2. Opção
                    // DataTable dt = excelReader.AsDataSet().Tables[0];

                    // Possui Dados no Excel
                    bool fgPreencheuJ005 = false;
                    foreach (DataRow r in dt.Rows)
                    {

                        if (r[0].ToString() != "Data Inicial")
                        {
                            if (!fgPreencheuJ005)
                            {
                                strTxt += criaLinhaIJ005(r[0].ToString(), r[1].ToString(), "1", r[3].ToString());
                                fgPreencheuJ005 = true;
                            }

                            strTxt += criaLinhaIJ100(
                                r[4].ToString()
                                , "D"
                                , r[6].ToString()
                                , r[14].ToString()
                                , r[7].ToString()
                                , r[15].ToString()
                                , r[9].ToString()
                                , r[12].ToString()
                                , r[11].ToString()
                                , ""
                                , "");

                            //strTxt += criaLinhaJ150(
                            //    "codAgl",
                            //    "indCodAgl_TD",
                            //    "nivelAgl",
                            //    "codAglSup",
                            //    "descrCodAgl",
                            //    "vlCta",
                            //    "indDCcta",
                            //    "indGrpDre",
                            //    "notaExpRef");
                        }
                    }
                }
                stream.Close();
            }
            return strTxt;
        }

        private string criaLinhaJ150(string codAgl, string indCodAgl_TD, string nivelAgl, string codAglSup,
            string descrCodAgl, string vlCta, string indDCcta, string indGrpDre, string notaExpRef)
        {
            string strTxt = "";
            string stLimit = "|";

            //01	REG
            strTxt += stLimit;
            strTxt += "J150";

            //02	COD_AGL
            strTxt += stLimit;
            strTxt += codAgl;

            //03	IND_COD_AGL
            strTxt += stLimit;
            strTxt += indCodAgl_TD;

            //04	NIVEL_AGL
            strTxt += stLimit;
            strTxt += nivelAgl;

            //05	COD_AGL_SUP
            strTxt += stLimit;
            strTxt += codAglSup;

            //06	DESCR_COD_AGL
            strTxt += stLimit;
            strTxt += descrCodAgl;

            //07	VL_CTA
            strTxt += stLimit;
            strTxt += vlCta;

            //08	IND_DC_CTA
            strTxt += stLimit;
            strTxt += indDCcta;

            //09	IND_GRP_DRE
            strTxt += stLimit;
            strTxt += indGrpDre;

            //10	NOTA_EXP_REF
            strTxt += stLimit;
            strTxt += notaExpRef;

            strTxt += stLimit;

            strTxt += System.Environment.NewLine;

            return strTxt;
        }

        private string criaLinhaIJ005(string dtIni, string dtFim, string idDem, string cabecDem)
        {
            string strTxt = "";
            string stLimit = "|";

            //01	REG
            strTxt += stLimit;
            strTxt += "J005";

            //02	DT_INI
            strTxt += stLimit;
            strTxt += dtIni.ToString().Replace("/", "").Replace("00:00:00", "").Trim();

            //03	DT_FIN
            strTxt += stLimit;
            strTxt += dtFim.ToString().Replace("/", "").Replace("00:00:00", "").Trim();

            //04	ID_DEM
            strTxt += stLimit;
            strTxt += idDem;

            //05	CAB_DEM
            strTxt += stLimit;
            strTxt += cabecDem;

            strTxt += stLimit;

            strTxt += System.Environment.NewLine;

            return strTxt;
        }
        #endregion

        #region J100
        private string criaLinhaIJ100(string codAgl, string indCodAgl, string nivelAgl, string codAglSup, string indGrpBal, string descrCodAgl,
            string vlCtaIni, string indDCctaIni, string vlCtaFin, string indDCctaFin, string notaExpRef)
        {
            string strTxt = "";
            string stLimit = "|";

            //01	REG
            strTxt += stLimit;
            strTxt += "J100";

            //02	COD_AGL
            strTxt += stLimit;
            strTxt += codAgl;

            //03	IND_COD_AGL ( T ou D )
            strTxt += stLimit;
            strTxt += indCodAgl;

            //04	NIVEL_AGL
            strTxt += stLimit;
            strTxt += nivelAgl;

            //05	COD_AGL_SUP
            strTxt += stLimit;
            strTxt += codAglSup;

            //06	IND_GRP_BAL
            strTxt += stLimit;
            strTxt += indGrpBal;

            //07	DESCR_COD_AGL
            strTxt += stLimit;
            strTxt += descrCodAgl;

            //08	VL_CTA_INI
            strTxt += stLimit;
            double ConvVlCtaIni = 0.00;
            double.TryParse(vlCtaIni, out ConvVlCtaIni);
            strTxt += ConvVlCtaIni.ToString("F");
            //strTxt += Convert.ToDouble(vlCtaIni).ToString("F");

            //09	IND_DC_CTA_INI
            strTxt += stLimit;
            strTxt += indDCctaIni;

            //10	VL_CTA_FIN
            strTxt += stLimit;
            double ConvVlCtaFim = 0.00;
            double.TryParse(vlCtaFin, out ConvVlCtaFim);
            strTxt += ConvVlCtaFim.ToString("F");
            //strTxt += Convert.ToDouble(vlCtaFin).ToString("F");

            //11	IND_DC_CTA_FIN
            strTxt += stLimit;
            strTxt += indDCctaFin;

            //12	NOTA_EXP_REF
            strTxt += stLimit;
            strTxt += notaExpRef;


            strTxt += stLimit;

            strTxt += System.Environment.NewLine;

            return strTxt;
        }
        #endregion

        #region J150
        //01	REG
        //02	COD_AGL
        //03	IND_COD_AGL
        //04	NIVEL_AGL
        //05	COD_AGL_SUP
        //06	DESCR_COD_AGL
        //07	VL_CTA
        //08	IND_DC_CTA
        //09	IND_GRP_DRE
        //10	NOTA_EXP_REF
        #endregion

        #region J900
        public string preencheRegistroJ900()
        {
            string strTxt = "";
            string stLimit = "|";

            //01    REG
            strTxt += stLimit;
            strTxt += "J900";

            //02	DNRC_ENCER
            //03	NUM_ORD
            //04	NAT_LIVRO
            //05	NOME
            //06	QTD_LIN
            //07	DT_INI_ESCR
            //08	DT_FIN_ESCR
            strTxt += stLimit;

            strTxt += System.Environment.NewLine;

            return strTxt;
        }

        #endregion

        #region J930
        public string preencheRegistroJ930()
        {
            string strTxt = "";
            string stLimit = "|";

            //01    REG
            strTxt += stLimit;
            strTxt += "I930";

            //02	IDENT_NOM
            //03	IDENT_CPF_CNPJ
            //04	IDENT_QUALIF
            //05	COD_ASSIN
            //06	IND_CRC
            //07	EMAIL
            //08	FONE
            //09	UF_CRC
            //10	NUM_SEQ_CRC
            //11	DT_CRC
            //12	IND_RESP_LEGAL

            strTxt += stLimit;

            strTxt += System.Environment.NewLine;

            return strTxt;
        }
        #endregion

        #region J990 (IMPORTANTE)
        public string preencheRegistroJ990()
        {
            string strTxt = "";
            string stLimit = "|";

            //01    REG
            strTxt += stLimit;
            strTxt += "J990";

            //02	QTD_LIN_J
            strTxt += stLimit;
            strTxt += "1";


            strTxt += stLimit;

            strTxt += System.Environment.NewLine;
            contJ990++;
            return strTxt;
        }
        #endregion

        #region 9001 (IMPORTANTE)
        public string preencheRegistro9001()
        {
            string strTxt = "";
            string stLimit = "|";

            //01    REG
            strTxt += stLimit;
            strTxt += "9001";

            //02	IND_MOV
            strTxt += stLimit;
            strTxt += "0";


            strTxt += stLimit;

            strTxt += System.Environment.NewLine;
            cont9001++;

            return strTxt;
        }
        #endregion

        #region 9900 (IMPORTANTE)
        public string preencheRegistro9900()
        {
            string strTxt = "";
            string stLimit = "|";

            ////01    REG
            //strTxt += stLimit;
            //strTxt += "9900";

            ////02	IND_MOV
            //strTxt += "0";

            //strTxt += stLimit;

            //strTxt += System.Environment.NewLine;

            strTxt += "|9900|0000|" + cont0000.ToString() + "|" + System.Environment.NewLine;
            strTxt += "|9900|0001|" + cont0001.ToString() + "|" + System.Environment.NewLine;
            strTxt += "|9900|0007|" + cont0007.ToString() + "|" + System.Environment.NewLine;
            strTxt += "|9900|0990|" + cont0990.ToString() + "|" + System.Environment.NewLine;
            strTxt += "|9900|I001|" + contI001.ToString() + "|"  + System.Environment.NewLine;
            strTxt += "|9900|I010|" + contI010.ToString() + "|" + System.Environment.NewLine;
            strTxt += "|9900|I030|" + contI030.ToString() + "|" + System.Environment.NewLine;
            strTxt += "|9900|I050|" + contI050.ToString() + "|" + System.Environment.NewLine;
            strTxt += "|9900|I051|" + contI051.ToString() + "|" + System.Environment.NewLine;
            strTxt += "|9900|I052|" + contI052.ToString() + "|" + System.Environment.NewLine;
            strTxt += "|9900|I100|" + contI100.ToString() + "|" + System.Environment.NewLine;
            strTxt += "|9900|I150|" + contI150.ToString() + "|" + System.Environment.NewLine;
            strTxt += "|9900|I155|" + contI155.ToString() + "|" + System.Environment.NewLine;
            strTxt += "|9900|I200|" + contI200.ToString() + "|" + System.Environment.NewLine;
            strTxt += "|9900|I250|" + contI250.ToString() + "|" + System.Environment.NewLine;
            //strTxt += "|9900|I200|18|" + System.Environment.NewLine;
            //strTxt += "|9900|I250|38|" + System.Environment.NewLine;
            strTxt += "|9900|I350|" + contI350.ToString() + "|" + System.Environment.NewLine;
            strTxt += "|9900|I355|" + contI355.ToString() + "|" + System.Environment.NewLine;
            strTxt += "|9900|I990|" + contI990.ToString() + "|" + System.Environment.NewLine;
            strTxt += "|9900|J001|0|" + System.Environment.NewLine;
            strTxt += "|9900|J005|0|" + System.Environment.NewLine;
            strTxt += "|9900|J100|0|" + System.Environment.NewLine;
            strTxt += "|9900|J900|0|" + System.Environment.NewLine;
            strTxt += "|9900|J930|0|" + System.Environment.NewLine;
            strTxt += "|9900|J990|1|" + System.Environment.NewLine;
            strTxt += "|9900|9001|1|" + System.Environment.NewLine;
            strTxt += "|9900|9990|1|" + System.Environment.NewLine;
            strTxt += "|9900|9999|1|" + System.Environment.NewLine;
            strTxt += "|9900|9900|27|" + System.Environment.NewLine;


            return strTxt;
        }
        #endregion

        #region 9990 (IMPORTANTE)
        public string preencheRegistro9990()
        {
            string strTxt = "";
            string stLimit = "|";

            //01    REG
            strTxt += stLimit;
            strTxt += "9990";

            //02	QTD_LIN_9
            strTxt += stLimit;
            strTxt += "30";


            strTxt += stLimit;

            strTxt += System.Environment.NewLine;
            cont9990++;

            return strTxt;
        }
        #endregion

        #region 9999 (IMPORTANTE)
        public string preencheRegistro9999()
        {
            string strTxt = "";
            string stLimit = "|";

            //01    REG
            strTxt += stLimit;
            strTxt += "9999";

            //02	QTD_LIN (total do arquivo)
            strTxt += stLimit;
            strTxt += "730538";


            strTxt += stLimit;

            strTxt += System.Environment.NewLine;
            cont9999++;

            return strTxt;
        }
        #endregion

        private void button4_Click(object sender, EventArgs e)
        {
            var result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }
    }




    /// <summary>
    /// Responsible for loading a WorkSheet from Sheet2 with
    /// a condition for a column of dates.
    /// </summary>
    /// <param name="FileName"></param>
    /// <param name="SheetName"></param>
    /// <param name="TheDate"></param>
    /// <returns></returns>
    //public DataTable LoadData(string FileName, string SheetName, DateTime TheDate)
    //{
    //    System.Text.StringBuilder sb = new System.Text.StringBuilder();
    //    DataTable dt = new DataTable();

    //    using (OleDbConnection cn = new OleDbConnection
    //    { ConnectionString = ConnectionString(FileName, "Yes") })
    //    {

    //        cn.Open();

    //        using (OleDbCommand cmd = new OleDbCommand
    //        {
    //            CommandText = "SELECT [Dates], [Office Plan] FROM [Sheet2$] WHERE [Dates] = " + TheDate.ToString(),
    //            Connection = cn
    //        }
    //         )

    //        {
    //            OleDbDataReader dr = cmd.ExecuteReader();
    //            dt.Load(dr);
    //        }

    //        return dt;
    //    }
    // }
}