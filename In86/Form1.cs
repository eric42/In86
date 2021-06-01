using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using excel = Microsoft.Office.Interop.Excel;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace In86
{
    public partial class frmIn86 : Form
    {
        private string arquivo;
        private string arquivo1;
        private string mensagem;
        List<string> C100 = new List<string>();
        List<string> A100 = new List<string>();
        List<string> C113 = new List<string>();
        List<string> R200 = new List<string>();
        List<string> R150 = new List<string>();
        List<string> R1100 = new List<string>();

        public frmIn86()
        {
            InitializeComponent();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            //define as propriedades do controle 
            //OpenFileDialog
            this.openFileDialog1.Multiselect = false;
            this.openFileDialog1.Title = "Selecionar Arquivo";
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Filter = "Texto(*.txt)|*.txt";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.ReadOnlyChecked = true;
            openFileDialog1.ShowReadOnly = true;
            openFileDialog1.DefaultExt = "txt";


            DialogResult dr = this.openFileDialog1.ShowDialog();

            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                txtArquivo.Text = openFileDialog1.FileName;
                arquivo = openFileDialog1.FileName;

                txtCarga.Enabled = true;
                btnSearch1.Enabled = true;
            }
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnConverter_Click(object sender, EventArgs e)
        {

            int counter = 0;
            string line;

            line = CarregaListaDados(ref counter);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                excelPackage.Workbook.Properties.Title = "IN86";

                int i, num;
                string[] titulos;

                GerarBlocoC(excelPackage, out i, out titulos, out num);

                num = GerarBlocoA(excelPackage, out i, num, out titulos);

                num = GerarBlocoR200(excelPackage, out i, num, out titulos);

                num = GerarBlocoR150(excelPackage, out i, num, out titulos);

                GerarBloco431(excelPackage, out i, out titulos);

                GerarBloco432(excelPackage, out i, out titulos);

                GerarBloco433(excelPackage, out i, out titulos);

                GerarBloco434(excelPackage, out i, out titulos);

                GerarBloco438(excelPackage, out i, out titulos);

                GerarBloco439(excelPackage, out i, out titulos);

                GerarBloco4103(excelPackage, out i, out titulos);

                GerarBloco4104(excelPackage, out i, out titulos);

                GerarBloco4105(excelPackage, out i, out titulos);

                GerarBloco4106(excelPackage, out i, out titulos);

                GerarBloco1CE(excelPackage, out i, out num, out titulos);

                GerarBloco441(excelPackage, out i, out titulos);

                GerarBloco442(excelPackage, out i, out titulos);

                GerarBlocoFaturamento(excelPackage); string caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                
                string path = caminho + @"\IN86.xlsx";
                File.WriteAllBytes(path, excelPackage.GetAsByteArray());
                MessageBox.Show("Concluído. Verifique em " + path + ".xls");
            }
        }

        private static void GerarBlocoFaturamento(ExcelPackage excelPackage)
        {
            // Aqui simplesmente adiciono a planilha inicial
            var sheet17 = excelPackage.Workbook.Worksheets.Add("Faturamento");
            sheet17.Name = "Faturamento";

            sheet17.Cells[1, 1].Value = "Informações do FATURAMENTO (Registro 0111 - Sped Contribuições)";
            sheet17.Cells[3, 1].Value = "Receita Bruta Não Cumulativa TRIBUTADA mercado interno (campo 02)";
            sheet17.Cells[5, 1].Value = "Receita Bruta Não Cumulativa NÃO TIBUTADA mercado interno (campo 03)";
            sheet17.Cells[7, 1].Value = "Receita Bruta Não Cumulativa EXPORTAÇÃO (campo 04)";
            sheet17.Cells[9, 1].Value = "Receita Bruta Total (campo 06)";

            sheet17.Cells[2, 2].Value = "Valores R$";
            sheet17.Cells[3, 2].Value = "172577237,77";
            sheet17.Cells[5, 2].Value = "655653,72";
            sheet17.Cells[7, 2].Value = "17154200,75";
            sheet17.Cells[9, 2].Value = "190387091,74";

            sheet17.Cells[2, 3].Value = "Indice Part";
            sheet17.Cells[3, 3].Formula = "B3/B9";
            sheet17.Cells[5, 3].Formula = "B5/B9";
            sheet17.Cells[7, 3].Formula = "B7/B9";
            sheet17.Cells[9, 3].Value = "=SOMA(C3:C8)";

            sheet17.Cells[3, 4].Formula = "C3*100";
            sheet17.Cells[5, 4].Formula = "C5*100";
            sheet17.Cells[7, 4].Formula = "C7*100";

        }

        private void GerarBloco442(ExcelPackage excelPackage, out int i, out string[] titulos)
        {
            // Aqui simplesmente adiciono a planilha inicial
            var sheet16 = excelPackage.Workbook.Worksheets.Add("4.4.2");
            sheet16.Name = "4.4.2";

            // Títulos
            i = 1;
            titulos = new String[] { "01 - Modelo", "02 - Série / sub", "03 - Num docto", "4 - Data emissão", "5 - Numero DI", "Linha preenchida IN86 - 4.4.2" };
            foreach (var titulo in titulos)
            {
                sheet16.Cells[1, i++].Value = titulo;
            }

            i = 2;
            foreach (string y in C100)
            {
                for (int j = 1; j < 7; j++)
                {
                    if (j == 1)
                    {
                        sheet16.Cells[i, 1].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C120\");EXT.TEXTO(\'Bloco C\'!B" + i + ";3;2))";
                    }
                    if (j == 2)
                    {
                        sheet16.Cells[i, 2].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C120\");SUBSTITUIR(EXT.TEXTO(\'Bloco C\'!B" + i + ";5;3);\" * \";\"\"))";
                    }
                    if (j == 3)
                    {
                        sheet16.Cells[i, 3].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C120\");EXT.TEXTO(\'Bloco C\'!B" + i + ";8;9))";
                    }
                    if (j == 4)
                    {
                        sheet16.Cells[i, 4].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C120\");EXT.TEXTO(\'Bloco C\'!B" + i + ";17;8))";
                    }
                    if (j == 5)
                    {
                        sheet16.Cells[i, 5].Value = "=SE(\'Bloco C\'!C" + i + "=\"C120\";\'Bloco C\'!E" + i + ")";
                    }
                    if (j == 6)
                    {
                        sheet16.Cells[i, 6].Value = "=CONCATENAR(A" + i + ";ESQUERDA(CONCATENAR(B" + i + ";REPT(\" \";5));5);C" + i + ";D" + i + ";REPT(0;10-NÚM.CARACT(SUBSTITUIR(E" + i + ";\",\";\"\")))&SUBSTITUIR(E" + i + ";\",\";))";
                    }
                }
            }
        }

        private void GerarBloco441(ExcelPackage excelPackage, out int i, out string[] titulos)
        {
            // Aqui simplesmente adiciono a planilha inicial
            var sheet15 = excelPackage.Workbook.Worksheets.Add("4.4.1");
            sheet15.Name = "4.4.1";

            // Títulos
            i = 1;
            titulos = new String[] { "01 - Modelo", "02 - Série / sub", "03 - Num docto", "4 - Data emissão", "5 - Numero do registro", "6 - Numero do despacho", "Linha preenchida IN86 - 4.4.1" };
            foreach (var titulo in titulos)
            {
                sheet15.Cells[1, i++].Value = titulo;
            }

            i = 2;
            foreach (string y in R1100)
            {
                for (int j = 1; j < 8; j++)
                {
                    if (j == 1)
                    {
                        sheet15.Cells[i, 1].Value = "=SE(\'Bloco 1 CE\'!D" + i + "=\"1105\";\'Bloco 1 CE\'!E" + i + ")";
                    }
                    if (j == 2)
                    {
                        sheet15.Cells[i, 2].Value = "=SE(\'Bloco 1 CE\'!D" + i + "=\"1105\";\'Bloco 1 CE\'!F" + i + ")";
                    }
                    if (j == 3)
                    {
                        sheet15.Cells[i, 3].Value = "=SE((\'Bloco 1 CE\'!D" + i + "=\"1105\");DIREITA(\'Bloco 1 CE\'!G" + i + ";6))";
                    }
                    if (j == 4)
                    {
                        sheet15.Cells[i, 4].Value = "=SE(\'Bloco 1 CE\'!D" + i + "=\"1105\";\'Bloco 1 CE\'!I" + i + ")";
                    }
                    if (j == 5)
                    {
                        sheet15.Cells[i, 5].Value = "=SE(\'Bloco 1 CE\'!D" + i + "=\"1105\";EXT.TEXTO(\'Bloco 1 CE\'!C" + i + ";1;12))";
                    }
                    if (j == 6)
                    {
                        sheet15.Cells[i, 6].Value = "=SE(\'Bloco 1 CE\'!D" + i + "=\"1105\";EXT.TEXTO(\'Bloco 1 CE\'!C" + i + ";13;12))";
                    }
                    if (j == 7)
                    {
                        sheet15.Cells[i, 7].Value = "=CONCATENAR(A" + i + ";ESQUERDA(CONCATENAR(B" + i + ";REPT(\" \";5));5);ESQUERDA(CONCATENAR(C" + i + ";REPT(\" \";9));9);D" + i + ";E" + i + ";REPT(0;12-NÚM.CARACT(SUBSTITUIR(F" + i + ";\",\";\"\")))&SUBSTITUIR(F" + i + ";\",\";))";
                    }
                }
            }
        }

        private void GerarBloco1CE(ExcelPackage excelPackage, out int i, out int num, out string[] titulos)
        {
            // Aqui simplesmente adiciono a planilha inicial
            var sheet14 = excelPackage.Workbook.Worksheets.Add("Bloco 1 CE");
            sheet14.Name = "Bloco 1 CE";

            // Títulos
            i = 1;
            titulos = new String[] { "Nr registro", "Nr despacho", "", "registro", "", "03 -NRO_DE", "", "", "06 - NRO_RE" };
            foreach (var titulo in titulos)
            {
                sheet14.Cells[1, i++].Value = titulo;
            }

            // Valores
            i = 2;
            num = 0;
            foreach (string y in R1100)
            {
                num = 0;
                string[] value = y.Split('|');//.Where(x => x != "");
                for (int j = 2; j < value.Count(); j++)
                {
                    if (!value[num].Equals(""))
                    {
                        sheet14.Cells[i, j].Value = value[num];
                    }
                    num++;
                }

                sheet14.Cells[i, 1].Value = "=SE(D" + i + "=\"1100\";I" + i + ";\"\")";
                sheet14.Cells[i, 2].Value = "=SE(D" + i + "=\"1100\";F" + i + ";\"\")";

                i++;
            }
        }

        private void GerarBloco4106(ExcelPackage excelPackage, out int i, out string[] titulos)
        {
            // Aqui simplesmente adiciono a planilha inicial
            var sheet13 = excelPackage.Workbook.Worksheets.Add("4.10.6");
            sheet13.Name = "4.10.6";

            // Títulos
            i = 1;
            titulos = new String[] { "1 - Série", "2 - Nr docto", "3 - DT Emissão", "4 - Participante", "5 - Nr item", "6 - CST PIS",
                    "7 - Alíquota", "8 - Base Calc", "09 - Vlr Crédito PIS - Receita Exportação", "10 - Vlr Crédito PIS - Receita Mercado interno",
                    "11 - Vlr Crédito PIS - Receita não tributada", "12 - Vlr PIS", "13 - CST COFINS", "14 - Alíq Cofins", "15 - BC Cofins",
                    "16 - Vlr Créd Cofins Receita Exportação", "17 - Vlr Créd Cofins - Receita Mercado interno", "18 - Vlr Créd Cofins Receita não tributada",
                    "19 - Valor Cofins", "20 - Dt Apropriação", "Linha preenchida IN25/10 - 4.10.6" };
            foreach (var titulo in titulos)
            {
                sheet13.Cells[1, i++].Value = titulo;
            }

            i = 2;
            foreach (string y in A100)
            {
                for (int j = 1; j < 22; j++)
                {
                    if (j == 1)
                    {
                        sheet13.Cells[i, 1].Value = "=SE((\'BLOCO A\'!C" + i + "=\"A170\");SUBSTITUIR(EXT.TEXTO(\'BLOCO A\'!B" + i + ";3;3);\" * \";\"\"))";
                    }
                    if (j == 2)
                    {
                        sheet13.Cells[i, 2].Value = "=SE((\'BLOCO A\'!C" + i + "=\"A170\");EXT.TEXTO(\'BLOCO A\'!B" + i + ";6;9))";
                    }
                    if (j == 3)
                    {
                        sheet13.Cells[i, 3].Value = "=SE((\'BLOCO A \'!C" + i + "=\"A170\");EXT.TEXTO(\'BLOCO A\'!B" + i + ";15;8))";
                    }
                    if (j == 4)
                    {
                        sheet13.Cells[i, 4].Value = "=SE((\'BLOCO A \'!C" + i + "=\"A170\");EXT.TEXTO(\'BLOCO A \'!B" + i + ";23;15))";
                    }
                    if (j == 5)
                    {
                        sheet13.Cells[i, 5].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A170\";TEXTO(\'BLOCO A \'!D" + i + ";\"000\"))";
                    }
                    if (j == 6)
                    {
                        sheet13.Cells[i, 6].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A170\";\'BLOCO A \'!K" + i + ")";
                    }
                    if (j == 7)
                    {
                        sheet13.Cells[i, 7].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A170\";TEXTO(\'BLOCO A \'!M" + i + ";\"0,0000\"))";
                    }
                    if (j == 8)
                    {
                        sheet13.Cells[i, 8].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A170\";TEXTO(\'BLOCO A \'!L" + i + ";\"0,000\"))";
                    }
                    if (j == 9)
                    {
                        sheet13.Cells[i, 9].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A170\";TEXTO(L" + i + "*Faturamento.C$7;\"0,00\"))";
                    }
                    if (j == 10)
                    {
                        sheet13.Cells[i, 10].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A170\";TEXTO(L" + i + "*Faturamento.C$3;\"0,00\"))";
                    }
                    if (j == 11)
                    {
                        sheet13.Cells[i, 11].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A170\";TEXTO(L" + i + "*Faturamento.C$5;\"0,00\"))";
                    }
                    if (j == 12)
                    {
                        sheet13.Cells[i, 12].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A170\";TEXTO(\'BLOCO A \'!N" + i + ";\"0,00\"))";
                    }
                    if (j == 13)
                    {
                        sheet13.Cells[i, 13].Value = "=F" + i;
                    }
                    if (j == 14)
                    {
                        sheet13.Cells[i, 14].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A170\";TEXTO(\'BLOCO A \'!Q" + i + ";\"0,0000\"))";
                    }
                    if (j == 15)
                    {
                        sheet13.Cells[i, 15].Value = "=H" + i;
                    }
                    if (j == 16)
                    {
                        sheet13.Cells[i, 16].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A170\";TEXTO(S" + i + "*Faturamento.C$7;\"0,00\"))";
                    }
                    if (j == 17)
                    {
                        sheet13.Cells[i, 17].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A170\";TEXTO(S" + i + "*Faturamento.C$3;\"0,00\"))";
                    }
                    if (j == 18)
                    {
                        sheet13.Cells[i, 18].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A170\";TEXTO(S" + i + "*Faturamento.C$5;\"0,00\"))";
                    }
                    if (j == 19)
                    {
                        sheet13.Cells[i, 19].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A170\";TEXTO(\'BLOCO A \'!R" + i + ";\"0,00\"))";
                    }
                    if (j == 20)
                    {
                        sheet13.Cells[i, 20].Value = "=SE((\'BLOCO A \'!C" + i + "=\"A170\");EXT.TEXTO(\'BLOCO A \'!B" + i + ";15;8))";
                    }
                    if (j == 21)
                    {
                        sheet13.Cells[i, 21].Value = "=CONCATENAR(ESQUERDA(CONCATENAR(A" + i + ";REPT(\" \"; 5));5);ESQUERDA(CONCATENAR(B" + i + ";REPT(\" \"; 9));9);C" + i + ";ESQUERDA(CONCATENAR(D" + i + ";REPT(\" \"; 14));14);E" + i + ";ESQUERDA(CONCATENAR(F" + i + ";REPT(\" \"; 2));2);" +
                            "REPT(0;8-NÚM.CARACT(SUBSTITUIR(G" + i + ";\",\";\"\")))&SUBSTITUIR(G" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(H" + i + ";\",\";\"\")))&SUBSTITUIR(H" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(I" + i + ";\",\";\"\")))&SUBSTITUIR(I" + i + ";\",\";);" +
                            "REPT(0;17-NÚM.CARACT(SUBSTITUIR(J" + i + ";\",\";\"\")))&SUBSTITUIR(J" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(K" + i + ";\",\";\"\")))&SUBSTITUIR(K" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(L" + i + ";\",\";\"\")))&SUBSTITUIR(L" + i + ";\",\";);" +
                            "ESQUERDA(CONCATENAR(M" + i + ";REPT(\" \"; 2));2);REPT(0;8-NÚM.CARACT(SUBSTITUIR(N" + i + ";\",\";\"\")))&SUBSTITUIR(N" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(O" + i + ";\",\";\"\")))&SUBSTITUIR(O" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(P" + i + ";\",\";\"\")))" +
                            "&SUBSTITUIR(P" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(Q" + i + ";\",\";\"\")))&SUBSTITUIR(Q" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(R" + i + ";\",\";\"\")))&SUBSTITUIR(R" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(S" + i + ";\",\";\"\")))" +
                            "&SUBSTITUIR(S" + i + ";\",\";);T" + i + ")";
                    }
                }
                i++;
            }
        }

        private void GerarBloco4105(ExcelPackage excelPackage, out int i, out string[] titulos)
        {
            // Aqui simplesmente adiciono a planilha inicial
            var sheet12 = excelPackage.Workbook.Worksheets.Add("4.10.5");
            sheet12.Name = "4.10.5";

            // Títulos
            i = 1;
            titulos = new String[] { "1 - Modelo docto", "2 - Série", "3 - Num do dcto", "4 - Dt Emissão", "5 - Cod Participante" ,
                "6 - Nr item", "7 - CST PIS", "8 - Alíquota", "9 - Base Calc", "10 - Vlr Crédito PIS - Receita Exportação", "11 - Vlr Crédito PIS - Receita Mercado interno",
                "12 - Vlr Crédito PIS - Receita não tributada", "13 - Vlr PIS", "14 - CST COFINS", "15 - Alíq Cofins", "16 - BC Cofins", "17 - Vlr Créd Cofins Receita Exportação",
                "18 - Vlr Créd Cofins - Receita Mercado interno", "19 - Vlr Créd Cofins Receita não tributada", "20 - Valor Cofins", "21 - Dt Apropriação", "Linha preenchida IN25/10 - 4.10.5"};

            foreach (var titulo in titulos)
            {
                sheet12.Cells[1, i++].Value = titulo;
            }

            i = 2;
            foreach (string y in C100)
            {
                for (int j = 1; j < 23; j++)
                {
                    if (j == 1)
                    {
                        sheet12.Cells[i, 1].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");EXT.TEXTO(\'Bloco C\'!B" + i + ";3;2))";
                    }
                    if (j == 2)
                    {
                        sheet12.Cells[i, 2].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");SUBSTITUIR(EXT.TEXTO(\'Bloco C\'!B" + i + ";14;3);\" * \";\"\"))";
                    }
                    if (j == 3)
                    {
                        sheet12.Cells[i, 3].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");EXT.TEXTO(\'Bloco C\'!B" + i + ";5;9))";
                    }
                    if (j == 4)
                    {
                        sheet12.Cells[i, 4].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");EXT.TEXTO(\'Bloco C\'!B" + i + ";17;8))";
                    }
                    if (j == 5)
                    {
                        sheet12.Cells[i, 5].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");EXT.TEXTO(\'Bloco C\'!B" + i + ";33;14))";
                    }
                    if (j == 6)
                    {
                        sheet12.Cells[i, 6].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!D" + i + ";\"000\"))";
                    }
                    if (j == 7)
                    {
                        sheet12.Cells[i, 7].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");SE(\'Bloco C\'!AA" + i + "=\"\";\"\";\'Bloco C\'!AA" + i + "))";
                    }
                    if (j == 8)
                    {
                        sheet12.Cells[i, 8].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!AC" + i + ";\"0,0000\"))";
                    }
                    if (j == 9)
                    {
                        sheet12.Cells[i, 9].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!AB" + i + ";\"0,000\"))";
                    }
                    if (j == 10)
                    {
                        sheet12.Cells[i, 10].Value = "=SE(G" + i + "=\"50\";\"0,00\";SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(M" + i + "*Faturamento!C$7;\"0,00\")))";
                    }
                    if (j == 11)
                    {
                        sheet12.Cells[i, 11].Value = "=SE(G" + i + "=\"50\";M" + i + ";SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(M" + i + "*Faturamento!C$3;\"0,00\")))";
                    }
                    if (j == 12)
                    {
                        sheet12.Cells[i, 12].Value = "=SE(G" + i + "=\"50\";\"0,00\";SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(M" + i + "*Faturamento!C$5;\"0,00\")))";
                    }
                    if (j == 13)
                    {
                        sheet12.Cells[i, 13].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!.C" + i + "=\"C170\");TEXTO(\'Bloco C\'!AF" + i + ";\"0,00\"))";
                    }
                    if (j == 14)
                    {
                        sheet12.Cells[i, 14].Value = "=G" + i;
                    }
                    if (j == 15)
                    {
                        sheet12.Cells[i, 15].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!AI" + i + ";\"0,0000\"))";
                    }
                    if (j == 16)
                    {
                        sheet12.Cells[i, 16].Value = "=I" + i;
                    }
                    if (j == 17)
                    {
                        sheet12.Cells[i, 17].Value = "=SE(N" + i + "=\"50\";\"0,00\";SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(T2*Faturamento!C$7;\"0,00\")))";
                    }
                    if (j == 18)
                    {
                        sheet12.Cells[i, 18].Value = "=SE(N" + i + "=\"50\";T" + i + ";SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(T" + i + "*Faturamento!C$3;\"0,00\")))";
                    }
                    if (j == 19)
                    {
                        sheet12.Cells[i, 19].Value = "=SE(N" + i + "=\"50\";\"0,00\";SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(T" + i + "*Faturamento!C$5;\"0,00\")))";
                    }
                    if (j == 20)
                    {
                        sheet12.Cells[i, 20].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!AL" + i + ";\"0,00\"))";
                    }
                    if (j == 21)
                    {
                        sheet12.Cells[i, 21].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");EXT.TEXTO(\'Bloco C\'!B" + i + ";25;8))";
                    }
                    if (j == 22)
                    {
                        sheet12.Cells[i, 22].Value = "=CONCATENAR(A" + i + ";ESQUERDA(CONCATENAR(B" + i + ";REPT(\" \";5));5);C" + i + ";D" + i + ";ESQUERDA(CONCATENAR(E" + i + ";REPT(\" \";14));14);F" + i + ";ESQUERDA(CONCATENAR(G" + i + ";REPT(\" \";2));2);REPT(0;8-NÚM.CARACT(SUBSTITUIR(H" + i + ";\",\";\"\")))" +
                            "&SUBSTITUIR(H" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(I" + i + ";\",\";\"\")))&SUBSTITUIR(I" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(J" + i + ";\",\";\"\")))&SUBSTITUIR(J" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(K" + i + ";\",\";\"\")))&SUBSTITUIR(K" + i + ";\",\";);" +
                            "REPT(0;17-NÚM.CARACT(SUBSTITUIR(L" + i + ";\",\";\"\")))&SUBSTITUIR(L" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(M" + i + ";\",\";\"\")))&SUBSTITUIR(M" + i + ";\",\";);ESQUERDA(CONCATENAR(N" + i + ";REPT(\" \";2));2);REPT(0;8-NÚM.CARACT(SUBSTITUIR(O" + i + ";\",\";\"\")))" +
                            "&SUBSTITUIR(O" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(P" + i + ";\",\";\"\")))&SUBSTITUIR(P" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(Q" + i + ";\",\";\"\")))&SUBSTITUIR(Q" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(R" + i + ";\",\";\"\")))&SUBSTITUIR(R" + i + ";\",\";);" +
                            "REPT(0;17-NÚM.CARACT(SUBSTITUIR(S" + i + ";\",\";\"\")))&SUBSTITUIR(S" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(T" + i + ";\",\";\"\")))&SUBSTITUIR(T" + i + ";\",\";);U" + i + ")";
                    }
                }
                i++;
            }
        }

        private void GerarBloco4104(ExcelPackage excelPackage, out int i, out string[] titulos)
        {
            // Aqui simplesmente adiciono a planilha inicial
            var sheet11 = excelPackage.Workbook.Worksheets.Add("4.10.4");
            sheet11.Name = "4.10.4";

            // Títulos
            i = 1;
            titulos = new String[] { "1 - Modelo docto", "2 - Série", "3 - Num do dcto", "4 - Dt Emissão", "5 - Nr item" ,
                "6 - CST PIS", "7 - Alíquota", "8 - Base Calc ", "9 - Vlr Crédito PIS - Receita Exportação", "10 - Vlr Crédito PIS - Receita Mercado interno", "11 - Vlr Crédito PIS - Receita não tributada",
                "12 - Vlr PIS", "13 - CST COFINS", "14 - Alíq Cofins", "15 - BC Cofins", "16 - Vlr Créd Cofins Receita Exportação", "17 - Vlr Créd Cofins - Receita Mercado interno",
                "18 - Vlr Créd Cofins Receita não tributada", "19 - Valor Cofins", "20 - Dt Apropriação", "Linha preenchida IN25/10 - 4.10.4"};

            foreach (var titulo in titulos)
            {
                sheet11.Cells[1, i++].Value = titulo;
            }

            i = 2;
            foreach (string y in C100)
            {
                for (int j = 1; j < 22; j++)
                {
                    if (j == 1)
                    {
                        sheet11.Cells[i, 1].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";1;1)=\"0\";EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");EXT.TEXTO(\'Bloco C\'!B" + i + ";3;2))";
                    }
                    if (j == 2)
                    {
                        sheet11.Cells[i, 2].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";1;1)=\"0\";EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");SUBSTITUIR(EXT.TEXTO(\'Bloco C\'!B" + i + ";14;3);\" * \";\"\"))";
                    }
                    if (j == 3)
                    {
                        sheet11.Cells[i, 3].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";1;1)=\"0\";EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"0\";\'Bloco C\'!C2=\"C170\");EXT.TEXTO(\'Bloco C\'!B" + i + ";5;9))";
                    }
                    if (j == 4)
                    {
                        sheet11.Cells[i, 4].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";1;1)=\"0\";EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");EXT.TEXTO(\'Bloco C\'!B" + i + ";17;8))";
                    }
                    if (j == 5)
                    {
                        sheet11.Cells[i, 5].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";1;1)=\"0\";EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!D" + i + ";\"000\"))";
                    }
                    if (j == 6)
                    {
                        sheet11.Cells[i, 6].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";1;1)=\"0\";EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");\'Bloco C\'!AA" + i + ")";
                    }
                    if (j == 7)
                    {
                        sheet11.Cells[i, 7].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";1;1)=\"0\";EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!AC" + i + ";\"#0,0000\"))";
                    }
                    if (j == 8)
                    {
                        sheet11.Cells[i, 8].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";1;1)=\"0\";EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!AB" + i + ";\"0,000#\"))";
                    }
                    if (j == 9)
                    {
                        sheet11.Cells[i, 9].Value = "=SE(F" + i + "=\"50\";\"0,00\";SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";1;1)=\"0\";EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(L" + i + "*Faturamento!C$7;\"0,00\")))";
                    }
                    if (j == 10)
                    {
                        sheet11.Cells[i, 10].Value = "=SE(F" + i + "=\"50\";L" + i + ";SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";1;1)=\"0\";EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(L" + i + "*Faturamento!C$3;\"0,00\")))";
                    }
                    if (j == 11)
                    {
                        sheet11.Cells[i, 11].Value = "=SE(F" + i + "=\"50\";\"0,00\";SE(F" + i + "=\"50\";\"0,00\";SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";1;1)=\"0\";EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(L" + i + "*Faturamento!C$5;\"0,00\"))))";
                    }
                    if (j == 12)
                    {
                        sheet11.Cells[i, 12].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";1;1)=\"0\";EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!AF" + i + ";\"#0,00#\"))";
                    }
                    if (j == 13)
                    {
                        sheet11.Cells[i, 13].Value = "=F2" + i;
                    }
                    if (j == 14)
                    {
                        sheet11.Cells[i, 14].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";1;1)=\"0\";EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!AI" + i + ";\"#0,0000#\"))";
                    }
                    if (j == 15)
                    {
                        sheet11.Cells[i, 15].Value = "=H" + i;
                    }
                    if (j == 16)
                    {
                        sheet11.Cells[i, 16].Value = "=SE(M" + i + "=\"50\";\"0,00\";SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";1;1)=\"0\";EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(S" + i + "*Faturamento!C$7;\"0,00\")))";
                    }
                    if (j == 17)
                    {
                        sheet11.Cells[i, 17].Value = "=SE(M" + i + "=\"50\";S" + i + ";SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";1;1)=\"0\";EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(S" + i + "*Faturamento!C$3;\"0,00\")))";
                    }
                    if (j == 18)
                    {
                        sheet11.Cells[i, 18].Value = "=SE(M" + i + "=\"50\";\"0,00\";SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";1;1)=\"0\";EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(S" + i + "*Faturamento!C$5;\"0,00\")))";
                    }
                    if (j == 19)
                    {
                        sheet11.Cells[i, 19].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";1;1)=\"0\";EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!AL" + i + ";\"#0,00#\"))";
                    }
                    if (j == 20)
                    {
                        sheet11.Cells[i, 20].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";1;1)=\"0\";EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");EXT.TEXTO(\'Bloco C\'!B" + i + ";25;8))";
                    }
                    if (j == 21)
                    {
                        sheet11.Cells[i, 21].Value = "=CONCATENAR(A" + i + ";ESQUERDA(CONCATENAR(B" + i + ";REPT(\" \";5));5);C" + i + ";D" + i + ";E" + i + ";ESQUERDA(CONCATENAR(F" + i + ";REPT(\" \";2));2);REPT(0;8-NÚM.CARACT(SUBSTITUIR(G" + i + ";\",\";\"\")))&SUBSTITUIR(G" + i + ";\",\";);" +
                            "REPT(0;17-NÚM.CARACT(SUBSTITUIR(H" + i + ";\",\";\"\")))&SUBSTITUIR(H" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(I" + i + ";\",\";\"\")))&SUBSTITUIR(I" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(J" + i + ";\",\";\"\")))&SUBSTITUIR(J" + i + ";\",\";);" +
                            "REPT(0;17-NÚM.CARACT(SUBSTITUIR(K" + i + ";\",\";\"\")))&SUBSTITUIR(K" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(L" + i + ";\",\";\"\")))&SUBSTITUIR(L" + i + ";\",\";);ESQUERDA(CONCATENAR(M" + i + ";REPT(\" \";2));2);REPT(0;8-NÚM.CARACT(SUBSTITUIR(N" + i + ";\",\";\"\")))" +
                            "&SUBSTITUIR(N" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(O" + i + ";\",\";\"\")))&SUBSTITUIR(O" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(P" + i + ";\",\";\"\")))&SUBSTITUIR(P" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(Q" + i + ";\",\";\"\")))&SUBSTITUIR(Q" + i + ";\",\";);" +
                            "REPT(0;17-NÚM.CARACT(SUBSTITUIR(R" + i + ";\",\";\"\")))&SUBSTITUIR(R" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(S" + i + ";\",\";\"\")))&SUBSTITUIR(S" + i + ";\",\";);T" + i + ")";
                    }
                }
                i++;
            }
        }

        private void GerarBloco4103(ExcelPackage excelPackage, out int i, out string[] titulos)
        {
            // Aqui simplesmente adiciono a planilha inicial
            var sheet10 = excelPackage.Workbook.Worksheets.Add("4.10.1");
            sheet10.Name = "4.10.1";

            // Títulos
            i = 1;
            titulos = new String[] { "1 - Modelo docto", "2 - Série", "3 - Num do dcto", "4 - Dt Emissão", "5 - Nr item" ,
                "6 - CST PIS", "7 - Alíquota", "8 - Base Calc ", "9 - Vlr PIS", "10 - CST Cofins", "11 - Alíq Cofins",
                "12 - BC Cofins", "13 - Valor Cofins", "14 - Dt Apropriação", "Linha preenchida IN25/10 - 4.10.1"};

            foreach (var titulo in titulos)
            {
                sheet10.Cells[1, i++].Value = titulo;
            }

            i = 2;
            foreach (string y in C100)
            {
                for (int j = 1; j < 16; j++)
                {
                    if (j == 1)
                    {
                        sheet10.Cells[i, 1].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");EXT.TEXTO(\'Bloco C\'!B" + i + ";3;2))";
                    }
                    if (j == 2)
                    {
                        sheet10.Cells[i, 2].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");SUBSTITUIR(EXT.TEXTO(\'Bloco C\'!B" + i + ";14;3);\" * \";\"\"))";
                    }
                    if (j == 3)
                    {
                        sheet10.Cells[i, 3].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");EXT.TEXTO(\'Bloco C\'!B" + i + ";5;9))";
                    }
                    if (j == 4)
                    {
                        sheet10.Cells[i, 4].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");EXT.TEXTO(\'Bloco C\'!B" + i + ";17;8))";
                    }
                    if (j == 5)
                    {
                        sheet10.Cells[i, 5].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!D" + i + ";\"000\"))";
                    }
                    if (j == 6)
                    {
                        sheet10.Cells[i, 6].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");\'Bloco C\'!AA" + i + ")";
                    }
                    if (j == 7)
                    {
                        sheet10.Cells[i, 7].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!AC" + i + ";\"#0,0000\"))";
                    }
                    if (j == 8)
                    {
                        sheet10.Cells[i, 8].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!AB" + i + ";\"0,000#\"))";
                    }
                    if (j == 9)
                    {
                        sheet10.Cells[i, 9].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!AF" + i + ";\"#0,00#\"))";
                    }
                    if (j == 10)
                    {
                        sheet10.Cells[i, 10].Value = "=F" + i;
                    }
                    if (j == 11)
                    {
                        sheet10.Cells[i, 11].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!AI" + i + ";\"#0,0000#\"))";
                    }
                    if (j == 12)
                    {
                        sheet10.Cells[i, 12].Value = "=H" + i;
                    }
                    if (j == 13)
                    {
                        sheet10.Cells[i, 13].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!AL" + i + ";\"#0,00#\"))";
                    }
                    if (j == 14)
                    {
                        sheet10.Cells[i, 14].Value = "=SE(E(EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");EXT.TEXTO(\'Bloco C\'!B" + i + ";25;8))";
                    }
                    if (j == 15)
                    {
                        sheet10.Cells[i, 15].Value = "=CONCATENAR(A" + i + ";ESQUERDA(CONCATENAR(B" + i + ";REPT(\" \";5));5);C" + i + ";D" + i + ";E" + i + ";ESQUERDA(CONCATENAR(F" + i + ";REPT(\" \";2));2);REPT(0;8-NÚM.CARACT(SUBSTITUIR(G" + i + ";\",\";\"\")))&SUBSTITUIR(G" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(H" + i + ";\",\";\"\")))" +
                            "&SUBSTITUIR(H" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(I" + i + ";\",\";\"\")))&SUBSTITUIR(I" + i + ";\",\";);ESQUERDA(CONCATENAR(J" + i + ";REPT(\" \";2));2);REPT(0;8-NÚM.CARACT(SUBSTITUIR(K" + i + ";\",\";\"\")))&SUBSTITUIR(K" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(L" + i + ";\",\";\"\")))&SUBSTITUIR(L" + i + ";\",\";);" +
                            "REPT(0;17-NÚM.CARACT(SUBSTITUIR(M" + i + ";\",\";\"\")))&SUBSTITUIR(M" + i + ";\",\";);N" + i + ")";
                    }

                }
                i++;
            }
        }

        private void GerarBloco439(ExcelPackage excelPackage, out int i, out string[] titulos)
        {
            // Aqui simplesmente adiciono a planilha inicial
            var sheet9 = excelPackage.Workbook.Worksheets.Add("4.3.9");
            sheet9.Name = "4.3.9";

            // Títulos
            i = 1;
            titulos = new String[] { "1 - Série", "2 - Nr docto", "3 - DT Emissão", "4 - Participante", "5 - Nr item" ,
                "6 - Código Serviço", "7 - Descrição compl", "8 - Valor do serviço", "9 - Desconto", "10 - Aliq ISS", "11 - Base Calculo ISS",
                "12 - VL ISS", "Linha preenchida IN25/10 - 4.3.9"};

            foreach (var titulo in titulos)
            {
                sheet9.Cells[1, i++].Value = titulo;
            }

            i = 2;
            foreach (string y in A100)
            {
                for (int j = 1; j < 14; j++)
                {
                    if (j == 1)
                    {
                        sheet9.Cells[i, 1].Value = "=SE((\'BLOCO A \'!C" + i + "=\"A170\");SUBSTITUIR(EXT.TEXTO(\'BLOCO A \'!B" + i + ";3;3);\" * \";\"\"))";
                    }
                    if (j == 2)
                    {
                        sheet9.Cells[i, 2].Value = "=SE((\'BLOCO A \'!C" + i + "=\"A170\");EXT.TEXTO(\'BLOCO A \'!B" + i + ";6;9))";
                    }
                    if (j == 3)
                    {
                        sheet9.Cells[i, 3].Value = "=SE((\'BLOCO A \'!C" + i + "=\"A170\");EXT.TEXTO(\'BLOCO A \'!B" + i + ";15;8))";
                    }
                    if (j == 4)
                    {
                        sheet9.Cells[i, 4].Value = "=SE((\'BLOCO A \'!C" + i + "=\"A170\");EXT.TEXTO(\'BLOCO A \'!B" + i + ";23;15))";
                    }
                    if (j == 5)
                    {
                        sheet9.Cells[i, 5].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A170\";TEXTO(\'BLOCO A \'!D" + i + ";\"000\"))";
                    }
                    if (j == 6)
                    {
                        sheet9.Cells[i, 6].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A170\";\'BLOCO A \'!E" + i + ")";
                    }
                    if (j == 8)
                    {
                        sheet9.Cells[i, 8].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A170\";TEXTO(\'BLOCO A \'!G" + i + ";\"#0,00#\"))";
                    }
                    if (j == 9)
                    {
                        sheet9.Cells[i, 9].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A170\";TEXTO(\'BLOCO A \'!H" + i + ";\"#0,00#\"))";
                    }
                    if (j == 10)
                    {
                        sheet9.Cells[i, 10].Value = "=TEXTO(((L" + i + "/K" + i + ")*100);\"#0,00#\")";
                    }
                    if (j == 11)
                    {
                        sheet9.Cells[i, 11].Value = "=H" + i;
                    }
                    if (j == 12)
                    {
                        sheet9.Cells[i, 12].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A170\";TEXTO(\'BLOCO A \'!W" + i + ";\"#0,00#\"))";
                    }
                    if (j == 13)
                    {
                        sheet9.Cells[i, 13].Value = "=CONCATENAR(ESQUERDA(CONCATENAR(A" + i + ";REPT(\" \"; 5));5);ESQUERDA(CONCATENAR(B" + i + ";REPT(\" \"; 9));9);C" + i + ";ESQUERDA(CONCATENAR(D" + i + ";REPT(\" \"; 14));14);E" + i + ";ESQUERDA(CONCATENAR(F" + i + ";REPT(\" \"; 20));20);" +
                            "ESQUERDA(CONCATENAR(G" + i + ";REPT(\" \"; 45));45);REPT(0;17-NÚM.CARACT(SUBSTITUIR(H" + i + ";\",\";\"\")))&SUBSTITUIR(H" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(I" + i + ";\",\";\"\")))&SUBSTITUIR(I" + i + ";\",\";);;REPT(0;5-NÚM.CARACT(SUBSTITUIR(J" + i + ";\",\";\"\")))" +
                            "&SUBSTITUIR(J" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(K" + i + ";\",\";\"\")))&SUBSTITUIR(K" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(L" + i + ";\",\";\"\")))&SUBSTITUIR(L" + i + ";\",\";))";
                    }
                }
                i++;
            }
        }

        private void GerarBloco438(ExcelPackage excelPackage, out int i, out string[] titulos)
        {
            // Aqui simplesmente adiciono a planilha inicial
            var sheet8 = excelPackage.Workbook.Worksheets.Add("4.3.8");
            sheet8.Name = "4.3.8";

            // Títulos
            i = 1;
            titulos = new String[] { "1 - Série", "2 - Nr docto", "3 - DT Emissão", "4 - Participante", "5 - Valor do serviço" ,
                "6 - Desconto", "7 - Aliq IRR", "8 - Base Calculo IRRF", "9 - VL IRRF", "Linha preenchida IN25/10 - 4.3.8"};

            foreach (var titulo in titulos)
            {
                sheet8.Cells[1, i++].Value = titulo;
            }

            i = 2;
            foreach (string y in A100)
            {
                for (int j = 1; j < 11; j++)
                {
                    if (j == 1)
                    {
                        sheet8.Cells[i, 1].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A100\";SUBSTITUIR(\'BLOCO A\'!H" + i + ";\" * \";\" \"))";
                    }
                    if (j == 2)
                    {
                        sheet8.Cells[i, 2].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A100\"; TEXTO(\'BLOCO A \'!J" + i + ";\"000000000\"))";
                    }
                    if (j == 3)
                    {
                        sheet8.Cells[i, 3].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A100\";\'BLOCO A \'!L" + i + ")";
                    }
                    if (j == 4)
                    {
                        sheet8.Cells[i, 4].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A100\";\'BLOCO A \'!F" + i + ")";
                    }
                    if (j == 5)
                    {
                        sheet8.Cells[i, 5].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A100\";TEXTO(\'BLOCO A \'!N" + i + ";\"#0,00#\"))";
                    }
                    if (j == 6)
                    {
                        sheet8.Cells[i, 6].Value = "=SE(\'BLOCO A \'!C" + i + "=\"A100\";TEXTO(\'BLOCO A \'!P" + i + ";\"#0,00#\"))";
                    }
                    if (j == 10)
                    {
                        sheet8.Cells[i, 10].Value = "=CONCATENAR(ESQUERDA(CONCATENAR(A" + i + ";REPT(\" \"; 5));5);B" + i + ";C" + i + ";ESQUERDA(CONCATENAR(D" + i + ";REPT(\" \"; 14));14);REPT(0;17-NÚM.CARACT(SUBSTITUIR(E" + i + ";\",\";\"\")))&SUBSTITUIR(E" + i + ";\",\";);" +
                            "REPT(0;17-NÚM.CARACT(SUBSTITUIR(F" + i + ";\",\";\"\")))&SUBSTITUIR(F" + i + ";\",\";);REPT(0;5-NÚM.CARACT(SUBSTITUIR(G" + i + ";\",\";\"\")))&SUBSTITUIR(G" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(H" + i + ";\",\";\"\")))&SUBSTITUIR(H" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(I" + i + ";\",\";\"\")))&SUBSTITUIR(I" + i + ";\",\";))";
                    }
                }
                i++;
            }
        }

        private void GerarBloco434(ExcelPackage excelPackage, out int i, out string[] titulos)
        {
            // Aqui simplesmente adiciono a planilha inicial
            var sheet7 = excelPackage.Workbook.Worksheets.Add("4.3.4");
            sheet7.Name = "4.3.4";

            // Títulos
            i = 1;
            titulos = new String[] { "1 - Modelo", "2 - Série", "3 - Nr docto", "4 - DT Emissão", "5 - Participante" ,
                "6 - Nr item", "7 - Cód Merc/Serv", "8 - Descrição compl", "9 - CFOP", "10 - Cod Nat", "11 - Clas Fisc Merc", "12 - Qtdade", "13 - unid",
                "14 - Vlr Unit", "15 -Vlr Tot Item", "16 - Desconto", "17 - Ind Trib IPI", "18 - Aliq IPI", "19 - BC IPI", "20 - Vlr IPI", "21 - CST ICMS",
                "22 - Ind ICMS", "23 - Aliq ICMS", "24 - BC ICMS ", "25 - Vlr ICMS Pr", "26 - BC ICMS ST", "27 -Vlr ICMS ST", "28 - Ind Mov", "29 - CST IPI",
                "Linha preenchida IN86 - 4.3.4"};

            foreach (var titulo in titulos)
            {
                sheet7.Cells[1, i++].Value = titulo;
            }


            i = 2;
            foreach (string y in C100)
            {
                for (int j = 1; j < 31; j++)
                {
                    if (j == 1)
                    {
                        sheet7.Cells[i, 1].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");EXT.TEXTO(\'Bloco C\'!B" + i + ";3;2))";
                    }
                    if (j == 2)
                    {
                        sheet7.Cells[i, 2].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");SUBSTITUIR(EXT.TEXTO(\'Bloco C\'!B" + i + ";14;3);\" * \";\"\"))";
                    }
                    if (j == 3)
                    {
                        sheet7.Cells[i, 3].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");EXT.TEXTO(\'Bloco C\'!B" + i + ";5;9))";
                    }
                    if (j == 4)
                    {
                        sheet7.Cells[i, 4].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");EXT.TEXTO(\'Bloco C\'!B" + i + ";17;8))";
                    }
                    if (j == 5)
                    {
                        sheet7.Cells[i, 5].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");EXT.TEXTO(\'Bloco C\'!B" + i + ";33;14))";
                    }
                    if (j == 6)
                    {
                        sheet7.Cells[i, 6].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!D" + i + ";\"000\"))";
                    }
                    if (j == 7)
                    {
                        sheet7.Cells[i, 7].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");\'Bloco C\'!E" + i + ")";
                    }
                    if (j == 8)
                    {
                        sheet7.Cells[i, 8].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");SE(\'Bloco C\'!F" + i + "=\"\";\"\";\'Bloco C\'!F" + i + "))";
                    }
                    if (j == 9)
                    {
                        sheet7.Cells[i, 9].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");\'Bloco C\'!M" + i + ")";
                    }
                    if (j == 10)
                    {
                        sheet7.Cells[i, 10].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");\'Bloco C\'!N" + i + ")";
                    }
                    if (j == 11)
                    {
                        sheet7.Cells[i, 11].Value = "=PROCV(G" + i + ";\'0200\'.A$2:G$20302;7;0)";
                    }
                    if (j == 12)
                    {
                        sheet7.Cells[i, 12].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!G" + i + ";\"#0,000#\"))";
                    }
                    if (j == 13)
                    {
                        sheet7.Cells[i, 13].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");\'Bloco C\'!H" + i + ")";
                    }
                    if (j == 14)
                    {
                        sheet7.Cells[i, 14].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!I" + i + "/L" + i + ";\"#,0000\"))";
                    }
                    if (j == 15)
                    {
                        sheet7.Cells[i, 15].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!I" + i + ";\"#0,00#\"))";
                    }
                    if (j == 16)
                    {
                        sheet7.Cells[i, 16].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!J" + i + ";\"#0,00#\"))";
                    }
                    if (j == 17)
                    {
                        sheet7.Cells[i, 17].Value = "=SE(T" + i + "=\"0,00\";\"2\";\"1\")";
                    }
                    if (j == 18)
                    {
                        sheet7.Cells[i, 18].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!Y" + i + ";\"#0,00#\"))";
                    }
                    if (j == 19)
                    {
                        sheet7.Cells[i, 19].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!X" + i + ";\"#0,00#\"))";
                    }
                    if (j == 20)
                    {
                        sheet7.Cells[i, 20].Value = "=SE(E(ESQUERDA('Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!Z" + i + ";\"#0,00#\"))";
                    }
                    if (j == 21)
                    {
                        sheet7.Cells[i, 21].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");(TEXTO(\'Bloco C\'!L" + i + ";\"000\")))";
                    }
                    if (j == 22)
                    {
                        sheet7.Cells[i, 22].Value = "=SE(EXT.TEXTO(U" + i + ";2;1)<\"3\";1;SE(EXT.TEXTO(U" + i + ";2;1)=\"9\";3;SE(EXT.TEXTO(U" + i + ";2;1)=\"7\";1;2)))";
                    }
                    if (j == 23)
                    {
                        sheet7.Cells[i, 23].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");;TEXTO(\'Bloco C\'!P" + i + ";\"#0,00#\"))";
                    }
                    if (j == 24)
                    {
                        sheet7.Cells[i, 24].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");\'Bloco C\'!O" + i + ")";
                    }
                    if (j == 25)
                    {
                        sheet7.Cells[i, 25].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!Q" + i + ";\"#0,00#\"))";
                    }
                    if (j == 26)
                    {
                        sheet7.Cells[i, 26].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!R" + i + ";\"#0,00#\"))";
                    }
                    if (j == 27)
                    {
                        sheet7.Cells[i, 27].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!T" + i + ";\"#0,00#\"))";
                    }
                    if (j == 28)
                    {
                        sheet7.Cells[i, 28].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"1\";\'Bloco C\'!C" + i + "=\"C170\");SE(\'Bloco C\'!K" + i + "=\"0\";\"S\";\"N\"))";
                    }
                    if (j == 29)
                    {
                        sheet7.Cells[i, 29].Value = "=SE(Q" + i + "=\"1\";\"00\";\"02\")";
                    }
                    if (j == 30)
                    {
                        sheet7.Cells[i, 30].Value = "=CONCATENAR(A" + i + ";ESQUERDA(CONCATENAR(B+i+;REPT(\" \";5));5);C" + i + ";D" + i + ";ESQUERDA(CONCATENAR(E" + i + ";REPT(\" \";14));14);ESQUERDA(CONCATENAR(F" + i + ";REPT(\" \";3));3);" +
                            "ESQUERDA(CONCATENAR(G" + i + ";REPT(\" \";20));20);ESQUERDA(CONCATENAR(H" + i + ";REPT(\" \";45));45);I" + i + ";ESQUERDA(CONCATENAR(J" + i + ";REPT(\" \";6));6);ESQUERDA(CONCATENAR(K" + i + ";REPT(\" \";8));8);" +
                            "REPT(0;17-NÚM.CARACT(SUBSTITUIR(L" + i + ";\",\";\"\")))&SUBSTITUIR(L" + i + ";\",\";);ESQUERDA(CONCATENAR(M" + i + ";REPT(\" \";3));3);REPT(0;17-NÚM.CARACT(SUBSTITUIR(N" + i + ";\",\";\"\")))&SUBSTITUIR(N" + i + ";\",\";);" +
                            "REPT(0;17-NÚM.CARACT(SUBSTITUIR(O" + i + ";\",\";\"\")))&SUBSTITUIR(O" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(P" + i + ";\",\";\"\")))&SUBSTITUIR(P" + i + ";\",\";);Q" + i + "; REPT(0;5-NÚM.CARACT(SUBSTITUIR(R" + i + ";\",\";\"\")))" +
                            "&SUBSTITUIR(R" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(S" + i + ";\",\";\"\")))&SUBSTITUIR(S" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(T" + i + ";\",\";\"\")))&SUBSTITUIR(T" + i + ";\",\";);U" + i + ";V" + i + ";" +
                            "REPT(0;5-NÚM.CARACT(SUBSTITUIR(W" + i + ";\",\";\"\")))&SUBSTITUIR(W" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(X" + i + ";\",\";\"\")))&SUBSTITUIR(X" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(Y" + i + ";\",\";\"\")))" +
                            "&SUBSTITUIR(Y" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(Z" + i + ";\",\";\"\")))&SUBSTITUIR(Z" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(AA" + i + ";\",\";\"\")))&SUBSTITUIR(AA" + i + ";\",\";);AB" + i + ";" +
                            "ESQUERDA(CONCATENAR(AC" + i + ";REPT(\" \";2));2))";
                    }
                }
                i++;
            }
        }

        private void GerarBloco433(ExcelPackage excelPackage, out int i, out string[] titulos)
        {
            // Aqui simplesmente adiciono a planilha inicial
            var sheet6 = excelPackage.Workbook.Worksheets.Add("4.3.3");
            sheet6.Name = "4.3.3";

            // Títulos
            i = 1;
            titulos = new String[] { "1 - Modelo", "2 - Série", "3 - Nr docto", "4 - DT Emissão", "5 - Participante" ,
                "6 - DT Entrada", "7 - VL Mercadorias", "8 - Desc", "9 - Vlr Frete", "10 - Vlr Seguro", "11 - Vlr Out Despesas", "12 - Vlr IPI", "13 - Vlr ICMS ST",
                "14 -Vlr T NF", "15 - IE Sub", "16 - Tipo Fat", "17 - Observ", "18 - Ato Declaratorio", "19 - Mod Doc Ref", "20 - Ser/Sub Doc Ref", "21 - Num Doc Ref",
                "22 - Data Em Doc Ref", "23 -Part Doc Ref", "Linha preenchida IN86 - 4.3.3"};

            foreach (var titulo in titulos)
            {
                sheet6.Cells[1, i++].Value = titulo;
            }

            i = 2;
            foreach (string y in C100)
            {
                for (int j = 1; j < 25; j++)
                {
                    if (j == 1)
                    {
                        sheet6.Cells[i, 1].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"1\");\'Bloco C\'!G" + i + ")";
                    }
                    if (j == 2)
                    {
                        sheet6.Cells[i, 2].Value = "=TEXTO(SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"1\");SUBSTITUIR(\'Bloco C\'!I" + i + ";\" * \";\"\"));\"000\")";
                    }
                    if (j == 3)
                    {
                        sheet6.Cells[i, 3].Value = "=SE(E(\'Bloco C\'.C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"1\");(TEXTO(\'Bloco C\'!J" + i + ";\"000000000\")))";
                    }
                    if (j == 4)
                    {
                        sheet6.Cells[i, 4].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"1\");SE(\'Bloco C\'!L" + i + "=\"\";\"\";(\'Bloco C\'!L" + i + ")))";
                    }
                    if (j == 5)
                    {
                        sheet6.Cells[i, 5].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"1\");SE(\'Bloco C\'!F" + i + "=\"\";\"\";DIREITA(\'Bloco C\'!F" + i + ";14)))";
                    }
                    if (j == 6)
                    {
                        sheet6.Cells[i, 6].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"1\");SE(\'Bloco C\'!L" + i + "=\"\";\"\";(\'Bloco C\'!M" + i + ")))";
                    }
                    if (j == 7)
                    {
                        sheet6.Cells[i, 7].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"1\");TEXTO(\'Bloco C\'!R" + i + ";\"#0,00#\"))";
                    }
                    if (j == 8)
                    {
                        sheet6.Cells[i, 8].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"1\");TEXTO(\'Bloco C\'!P" + i + ";\"#0,00#\"))";
                    }
                    if (j == 9)
                    {
                        sheet6.Cells[i, 9].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"1\");TEXTO(\'Bloco C\'.T" + i + ";\"#0,00#\"))";
                    }
                    if (j == 10)
                    {
                        sheet6.Cells[i, 10].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"1\");TEXTO(\'Bloco C\'!U" + i + ";\"#0,00#\"))";
                    }
                    if (j == 11)
                    {
                        sheet6.Cells[i, 11].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"1\");TEXTO(\'Bloco C\'!V" + i + ";\"#0,00#\"))";
                    }
                    if (j == 12)
                    {
                        sheet6.Cells[i, 12].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"1\");TEXTO(\'Bloco C\'!AA" + i + ";\"#0,00#\"))";
                    }
                    if (j == 13)
                    {
                        sheet6.Cells[i, 13].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"1\");TEXTO(\'Bloco C\'!Z" + i + ";\"#0,00#\"))";
                    }
                    if (j == 14)
                    {
                        sheet6.Cells[i, 14].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"1\");TEXTO(\'Bloco C\'!N" + i + ";\"#0,00#\"))";
                    }
                    if (j == 16)
                    {
                        sheet6.Cells[i, 16].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"1\");SE(\'Bloco C\'!O" + i + "=\"0\";\"1\";SE(\'Bloco C\'!O" + i + "=\"1\";\"2\";\"\")))";
                    }
                    if (j == 19)
                    {
                        sheet6.Cells[i, 19].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"1\");SE(\'Bloco C\'!AN" + i + "=\"\";\"\";\'Bloco C\'!AN" + i + "))";
                    }
                    if (j == 20)
                    {
                        sheet6.Cells[i, 20].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"1\");SE(\'Bloco C\'!AO" + i + "=\"\";\"\";\'Bloco C\'!AO" + i + "))";
                    }
                    if (j == 21)
                    {
                        sheet6.Cells[i, 21].Value = "=SE(E(\'Bloco C\'.C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"1\");SE(\'Bloco C\'!AP" + i + "=\"\";\"\";\'Bloco C\'!AP" + i + "))";
                    }
                    if (j == 22)
                    {
                        sheet6.Cells[i, 22].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"1\");SE(\'Bloco C\'!AQ" + i + "=\"\";\"\";\'Bloco C\'!AQ" + i + "))";
                    }
                    if (j == 23)
                    {
                        sheet6.Cells[i, 23].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"1\");SE(\'Bloco C\'!AR" + i + "=\"\";\"\";\'Bloco C\'!AR" + i + "))";
                    }
                    if (j == 24)
                    {
                        sheet6.Cells[i, 24].Value = "=CONCATENAR(A" + i + "; ESQUERDA(CONCATENAR(B" + i + ";REPT(\" \"; 5));5);C" + i + ";ESQUERDA(CONCATENAR(D" + i + ";REPT(\" \"; 8));8);ESQUERDA(CONCATENAR" +
                            "(E" + i + ";REPT(\" \"; 14));14);ESQUERDA(CONCATENAR(F" + i + ";REPT(\" \"; 8));8);REPT(0;17-NÚM.CARACT(SUBSTITUIR(G" + i + ";\",\";\"\")))&SUBSTITUIR(G" + i + ";\",\";);" +
                            "REPT(0;17-NÚM.CARACT(SUBSTITUIR(H" + i + ";\",\";\"\")))&SUBSTITUIR(H" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(I" + i + ";\",\";\"\")))&SUBSTITUIR(I" + i + ";\",\";);" +
                            "REPT(0;17-NÚM.CARACT(SUBSTITUIR(J" + i + ";\",\";\"\")))&SUBSTITUIR(J" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(K" + i + ";\",\";\"\")))&SUBSTITUIR(K" + i + ";\",\";);" +
                            "REPT(0;17-NÚM.CARACT(SUBSTITUIR(L" + i + ";\",\";\"\")))&SUBSTITUIR(L" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(M" + i + ";\",\";\"\")))&SUBSTITUIR(M" + i + ";\",\";);" +
                            "REPT(0;17-NÚM.CARACT(SUBSTITUIR(N" + i + ";\",\";\"\")))&SUBSTITUIR(N" + i + ";\",\";); ESQUERDA(CONCATENAR(O" + i + ";REPT(\" \"; 14));14);ESQUERDA(CONCATENAR(P" + i + ";" +
                            "REPT(\" \"; 1));1);ESQUERDA(CONCATENAR(Q" + i + ";REPT(\" \"; 45));45);ESQUERDA(CONCATENAR(R" + i + ";REPT(\" \"; 50));50);ESQUERDA(CONCATENAR(S" + i + ";REPT(\" \"; 2));2);" +
                            "ESQUERDA(CONCATENAR(T" + i + ";REPT(\" \"; 5));5);ESQUERDA(CONCATENAR(U" + i + ";REPT(\" \"; 9));9);ESQUERDA(CONCATENAR(V" + i + ";REPT(\" \"; 8));8);ESQUERDA(CONCATENAR" +
                            "(W" + i + ";REPT(\" \"; 14));14))";
                    }
                }
                i++;
            }
        }

        private void GerarBloco432(ExcelPackage excelPackage, out int i, out string[] titulos)
        {
            // Aqui simplesmente adiciono a planilha inicial
            var sheet5 = excelPackage.Workbook.Worksheets.Add("4.3.2");
            sheet5.Name = "4.3.2";

            // Títulos
            i = 1;
            titulos = new String[] { "1 - Ind Movto", "2 - Modelo", "3 - Série", "4 - Nr docto", "5 - DT Emissão" ,
                "6 - Nr item", "7 - Cód Merc/Serv", "8 - Descrição compl", "9 - CFOP", "10 - Cod Nat", "11 - Clas Fisc Merc", "12 - Qtdade", "13 - unid",
                "14 - Vlr Unit", "15 - Vlr Tot Item", "16 - Desconto", "17 - Ind Trib IPI", "18 - Aliq IPI", "19 - BC IPI", "20 - Vlr IPI", "21 - CST ICMS",
                "22 - Ind ICMS", "23 - Aliq ICMS", "24 - BC ICMS", "25 - Vlr ICMS Pr", "26 - BC ICMS ST", "27 - Vlr ICMS ST", "28 - Ind Mov", "29 - CST IPI",
                "Linha Preenchida IN86 - 4.3.2"};

            foreach (var titulo in titulos)
            {
                sheet5.Cells[1, i++].Value = titulo;
            }


            i = 2;
            foreach (string y in C100)
            {
                for (int j = 1; j < 31; j++)
                {
                    if (j == 1)
                    {
                        sheet5.Cells[i, 1].Value = "= SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\"); SE(EXT.TEXTO(\'Bloco C\'!B" + i + ";2;1)=\"0\";\"E\";\"S\"))";
                    }
                    if (j == 2)
                    {
                        sheet5.Cells[i, 2].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");EXT.TEXTO(\'Bloco C\'!B" + i + ";3;2))";
                    }
                    if (j == 3)
                    {
                        sheet5.Cells[i, 3].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");SUBSTITUIR(EXT.TEXTO(\'Bloco C\'!B" + i + ";14;3);\" * \";\"\"))";
                    }
                    if (j == 4)
                    {
                        sheet5.Cells[i, 4].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");EXT.TEXTO(\'Bloco C\'!B" + i + ";5;9))";
                    }
                    if (j == 5)
                    {
                        sheet5.Cells[i, 5].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");EXT.TEXTO(\'Bloco C\'!B" + i + ";17;8))";
                    }
                    if (j == 6)
                    {
                        sheet5.Cells[i, 6].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!D" + i + ";\"000\"))";
                    }
                    if (j == 7)
                    {
                        sheet5.Cells[i, 7].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");\'Bloco C\'!E" + i + ")";
                    }
                    if (j == 8)
                    {
                        sheet5.Cells[i, 8].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");\'Bloco C\'!F" + i + ")";
                    }
                    if (j == 9)
                    {
                        sheet5.Cells[i, 9].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");\'Bloco C\'!M" + i + ")";
                    }
                    if (j == 10)
                    {
                        sheet5.Cells[i, 10].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");\'Bloco C\'!N" + i + ")";
                    }
                    if (j == 11)
                    {
                        sheet5.Cells[i, 11].Value = "=PROCV(G" + i + ";'0200'!A$2:G$35873;7;FALSO)";
                    }
                    if (j == 12)
                    {
                        sheet5.Cells[i, 12].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!G" + i + ";\"#0,000#\"))";
                    }
                    if (j == 13)
                    {
                        sheet5.Cells[i, 13].Value = "=SE(E(ESQUERDA(\'Bloco C\'.B" + i + ";1)=\"0\";\'Bloco C\'.C" + i + "=\"C170\");\'Bloco C\'.H" + i + ")";
                    }
                    if (j == 14)
                    {
                        sheet5.Cells[i, 14].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!I" + i + "/L" + i + ";\"#,0000\"))";
                    }
                    if (j == 15)
                    {
                        sheet5.Cells[i, 15].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!I" + i + ";\"#0,00#\"))";
                    }
                    if (j == 16)
                    {
                        sheet5.Cells[i, 16].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";z'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!J" + i + ";\"#0,00#\"))";
                    }
                    if (j == 17)
                    {
                        sheet5.Cells[i, 17].Value = "=SE(T" + 1 + "=\"0,00\";\"2\";\"1\")";
                    }
                    if (j == 18)
                    {
                        sheet5.Cells[i, 18].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!Y" + i + ";\"#0,00#\"))";
                    }
                    if (j == 19)
                    {
                        sheet5.Cells[i, 19].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C1=\"C170\");TEXTO(\'Bloco C\'!X" + i + ";\"#0,00#\"))";
                    }
                    if (j == 20)
                    {
                        sheet5.Cells[i, 20].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!Z" + i + ";\"#0,00#\"))";
                    }
                    if (j == 21)
                    {
                        sheet5.Cells[i, 21].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");\'Bloco C\'!L" + i + ")";
                    }
                    if (j == 22)
                    {
                        sheet5.Cells[i, 22].Value = "=SE(EXT.TEXTO(U" + i + ";2;1)<\"3\";1;SE(EXT.TEXTO(U" + i + ";2;1)=\"9\";3;SE(EXT.TEXTO(U" + i + ";2;1)=\"7\";1;2)))";
                    }
                    if (j == 23)
                    {
                        sheet5.Cells[i, 23].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");;TEXTO(\'Bloco C\'!P" + i + ";\"#0,00#\"))";
                    }
                    if (j == 24)
                    {
                        sheet5.Cells[i, 24].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");\'Bloco C\'!O" + i + ")";
                    }
                    if (j == 25)
                    {
                        sheet5.Cells[i, 25].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!Q" + i + ";\"#0,00#\"))";
                    }
                    if (j == 26)
                    {
                        sheet5.Cells[i, 26].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!R" + i + ";\"#0,00#\"))";
                    }
                    if (j == 27)
                    {
                        sheet5.Cells[i, 27].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");TEXTO(\'Bloco C\'!T" + i + ";\"#0,00#\"))";
                    }
                    if (j == 28)
                    {
                        sheet5.Cells[i, 28].Value = "=SE(E(ESQUERDA(\'Bloco C\'!B" + i + ";1)=\"0\";\'Bloco C\'!C" + i + "=\"C170\");SE(\'Bloco C\'!K" + i + "=\"0\";\"S\";\"N\"))";
                    }
                    if (j == 29)
                    {
                        sheet5.Cells[i, 29].Value = "=SE(E(Q" + i + "=\"1\";A" + i + "=\"S\");\"50\";SE(E(Q" + i + "=\"2\";A" + i + "=\"S\");\"52\";SE(E(Q" + i + "=\"1\";A" + i + "=\"E\");\"00\";\"02\")))";
                    }
                    if (j == 30)
                    {
                        sheet5.Cells[i, 30].Value = "=CONCATENAR(A" + i + ";B" + i + ";ESQUERDA(CONCATENAR(C" + i + ";REPT(\" \";5));5);D" + i + ";E" + i + ";F" + i + ";ESQUERDA(CONCATENAR(G" + i + ";REPT(\" \";20));20);ESQUERDA(CONCATENAR(H" + i + ";" +
                            "REPT(\" \";45));45);I" + i + ";ESQUERDA(CONCATENAR(J" + i + ";REPT(\" \";6));6);REPT(0;8-NÚM.CARACT(SUBSTITUIR(K" + i + ";\",\";\"\")))&SUBSTITUIR(K" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(L" + i + ";\",\";" +
                            "\"\")))&SUBSTITUIR(L" + i + ";\",\";);ESQUERDA(CONCATENAR(M" + i + ";REPT(\" \";3));3);REPT(0;17-NÚM.CARACT(SUBSTITUIR(N" + i + ";\",\";\"\")))&SUBSTITUIR(N" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(O" + i + ";\",\";\"\")))" +
                            "&SUBSTITUIR(O" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(P" + i + ";\",\";\"\")))&SUBSTITUIR(P" + i + ";\",\";);Q" + i + "; REPT(0;5-NÚM.CARACT(SUBSTITUIR(R" + i + ";\",\";\"\")))&SUBSTITUIR(R" + i + ";\",\";);REPT(0;17-NÚM.CARACT" +
                            "(SUBSTITUIR(S" + i + ";\",\";\"\")))&SUBSTITUIR(S" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(T" + i + ";\",\";\"\")))&SUBSTITUIR(T" + i + ";\",\";);U" + i + ";V" + i + ";REPT(0;5-NÚM.CARACT(SUBSTITUIR(W" + i + ";\",\";\"\")))" +
                            "&SUBSTITUIR(W" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(X" + i + ";\",\";\"\")))&SUBSTITUIR(X" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(Y" + i + ";\",\";\"\")))&SUBSTITUIR(Y" + i + ";\",\";);REPT(0;17-NÚM.CARACT" +
                            "(SUBSTITUIR(Z" + i + ";\",\";\"\")))&SUBSTITUIR(Z" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(AA" + i + ";\",\";\"\")))&SUBSTITUIR(AA" + i + ";\",\";);AB" + i + ";ESQUERDA(CONCATENAR(AC" + i + ";REPT(\" \";2));2))";
                    }
                }
                i++;
            }
        }

        private void GerarBloco431(ExcelPackage excelPackage, out int i, out string[] titulos)
        {
            // Aqui simplesmente adiciono a planilha inicial
            var sheet4 = excelPackage.Workbook.Worksheets.Add("4.3.1");
            sheet4.Name = "4.3.1";

            // Títulos
            i = 1;
            titulos = new String[] { "1 - Ind Movto", "2 - Modelo", "3 - Série", "4 - Nr docto", "5 - DT Emissão" ,
                "6 - Participante", "7 - DT Entrada", "8 - Vl Mercadorias", "9 - Desc", "10 - Vlr Frete", "11 - Vlr Seguro", "12 - Vlr Out Despesas", "13 - Vlr IPI",
                "14 - Vlr ICMS ST", "15 - VlrNF", "16 - IE Sub", "17 - Via Transp", "18 - Código Transp", "19 - Qt Vol", "20 - Esp Volume", "21 - Peso Bruto",
                "22 - Peso Liq", "23 - Mod Frete", "24 - Ident Veic", "25 - Ind Canc", "26 - Tipo Fat", "27 - Observ", "28 - ADE", "29 - Mod doc Ref",
                "30 - Ser Sub", "31 - Nr doc ref", "32 - DT Emis Ref",  "33 - Cod Part Ref", "Linha preenchida IN86 - 4.3.1"};

            foreach (var titulo in titulos)
            {
                sheet4.Cells[1, i++].Value = titulo;
            }

            i = 2;
            foreach (string y in C100)
            {
                for (int j = 1; j < 35; j++)
                {
                    if (j == 1)
                    {
                        sheet4.Cells[i, 1].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"0\");SE(\'Bloco C\'!D" + i + "=\"0\";\"E\";\"S\"))";
                    }
                    if (j == 2)
                    {
                        sheet4.Cells[i, 2].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"0\");\'Bloco C\'!G" + i + ")";
                    }
                    if (j == 3)
                    {
                        sheet4.Cells[i, 3].Value = "=TEXTO(SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"0\");SUBSTITUIR(\'Bloco C\'!I" + i + ";\"*\";\"\"));\"000\")";
                    }
                    if (j == 4)
                    {
                        sheet4.Cells[i, 4].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"0\");TEXTO(\'Bloco C\'!J" + i + ";\"000000000\")))";
                    }
                    if (j == 5)
                    {
                        sheet4.Cells[i, 5].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"0\");SE(\'Bloco C\'!L" + i + "=\"\";\"\";(\'Bloco C\'!L" + i + ")))";
                    }
                    if (j == 6)
                    {
                        sheet4.Cells[i, 6].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"0\");SE(\'Bloco C\'!F" + i + "=\"\";\"\";DIREITA(\'Bloco C\'!F" + i + ";14)))";
                    }
                    if (j == 7)
                    {
                        sheet4.Cells[i, 7].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"0\");SE(\'Bloco C\'!M" + i + "=\"\";\"\";(\'Bloco C\'!M" + i + ")))";
                    }
                    if (j == 8)
                    {
                        sheet4.Cells[i, 8].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"0\");TEXTO(\'Bloco C\'!R" + i + ";\"#0,00#\"))";
                    }
                    if (j == 9)
                    {
                        sheet4.Cells[i, 9].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"0\");TEXTO(\'Bloco C\'!P" + i + ";\"#0,00#\"))";
                    }
                    if (j == 10)
                    {
                        sheet4.Cells[i, 10].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"0\");TEXTO(\'Bloco C'!T" + i + ";\"#0,00#\"))";
                    }
                    if (j == 11)
                    {
                        sheet4.Cells[i, 11].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "\"0\");TEXTO(\'Bloco C\'!U" + i + ";\"#0,00#\"))";
                    }
                    if (j == 12)
                    {
                        sheet4.Cells[i, 12].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "\"0\");TEXTO(\'Bloco C\'!V" + i + ";\"#0,00#\"))";
                    }
                    if (j == 13)
                    {
                        sheet4.Cells[i, 13].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "\"0\");TEXTO(\'Bloco C\'!AA" + i + ";\"#0,00#\"))";
                    }
                    if (j == 14)
                    {
                        sheet4.Cells[i, 14].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"0\");TEXTO(\'Bloco C\'!Z" + i + ";\"#0,00#\"))";
                    }
                    if (j == 15)
                    {
                        sheet4.Cells[i, 15].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"0\");TEXTO(\'Bloco C\'!N" + i + ";\"#0,00#\"))";
                    }
                    if (j == 23)
                    {
                        sheet4.Cells[i, 23].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"0\");SE(OU(\'Bloco C\'!S" + i + "=\"1\";\'Bloco C\'!S" + i + "=\"2\";\'Bloco C\'!S" + i + "=\"9\");\"FOB\";\"CIF\"))";
                    }
                    if (j == 25)
                    {
                        sheet4.Cells[i, 25].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"0\");SE(OU(\'Bloco C\'!H" + i + "=\"02\";\'Bloco C\'!H" + i + "=\"03\";\'Bloco C\'!H" + i + "=\"04\";\'Bloco C\'!H" + i + "=\"05\");\"S\";\"N\"))";
                    }
                    if (j == 26)
                    {
                        sheet4.Cells[i, 26].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "\"0\");SE(\'Bloco C\'!O" + i + "=\"0\";\"1\";SE(\'Bloco C\'!O" + i + "=\"1\";\"2\";\"\")))";
                    }
                    if (j == 29)
                    {
                        sheet4.Cells[i, 29].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C'!E" + i + "=\"0\");SE(\'Bloco C\'!AN" + i + "\"\";\"\";\'Bloco C\'!AN" + i + "))";
                    }
                    if (j == 30)
                    {
                        sheet4.Cells[i, 30].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"0\");SE(\'Bloco C\'!AO" + i + "=\"\";\"\";\'Bloco C\'!AO" + i + "))";
                    }
                    if (j == 31)
                    {
                        sheet4.Cells[i, 31].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"0\");\'Bloco C\'!AP" + i + "=\"\";\"\";\'Bloco C\'!AP" + i + "))";
                    }
                    if (j == 32)
                    {
                        sheet4.Cells[i, 32].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"0\");SE(\'Bloco C\'!AQ" + i + "\"\";\"\";\'Bloco C\'!AQ" + i + "))";
                    }
                    if (j == 33)
                    {
                        sheet4.Cells[i, 33].Value = "=SE(E(\'Bloco C\'!C" + i + "=\"C100\";\'Bloco C\'!E" + i + "=\"0\");SE(\'Bloco C\'!AR" + i + "=\"\";\"\";\'Bloco C\'!AR" + i + "))";
                    }
                    if (j == 34)
                    {
                        sheet4.Cells[i, 34].Value = "=CONCATENAR(A" + i + ";B" + i + ";ESQUERDA(CONCATENAR(C" + i + ";REPT(\" \";5));5);D" + i + ";ESQUERDA(CONCATENAR(E" + i + ";REPT(\" \";8));8);ESQUERDA(CONCATENAR(F" + i +
                            ";REPT(\" \";14));14);ESQUERDA(CONCATENAR(G" + i + ";REPT(\" \";8));8);REPT(0;17-NÚM.CARACT(SUBSTITUIR(H" + i + ";\",\";\"\")))&SUBSTITUIR(H" + i + "\",\";\"\");REPT(0;17 - NÚM.CARACT(SUBSTITUIR(I" + i +
                            ";\",\";\"\")))&SUBSTITUIR(I" + i + ";\",\";);REPT(0;17-NÚM.CARACT(J" + i + ";\",\";\"\")))&SUBSTITUIR(J" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(K" + i + ";\",\";\"\")))&SUBSTITUIR(K" + i + ";\",\";);" +
                            "REPT(0;17-NÚM.CARACT(L" + i + ";\",\";\"\")))&SUBSTITUIR(L" + i + ";\",\";\"\");REPT(0;17-NÚM.CARACT(SUBSTITUIR(M" + i + ";\",\";\"\")))&SUBSTITUIR(M" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(N" + i + ";" +
                            "\",\";\"\")))&SUBSTITUIR(N" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(O" + i + ";\",\";\"\")))&SUBSTITUIR(O" + i + ";\",\";);ESQUERDA(CONCATENAR(P" + i + ";REPT(\" \";14))14;);ESQUERDA(CONCATENAR(Q" + i +
                            ";REPT(\" \";15));15);ESQUERDA(CONCATENAR(R" + i + ";REPT(\" \";14));14);REPT(0;17-NÚM.CARACT(SUBSTITUIR(S" + i + ";\",\";\"\")))&SUBSTITUIR(S" + i + ";\",\";);ESQUERDA(CONCATENAR(T" + i + ";REPT(\" \";10));10);" +
                            "REPT(0;17-NÚM.CARACT(SUBSTITUIR(U" + i + ";\",\";\"\")))&SUBSTITUIR(U" + i + ";\",\";);REPT(0;17-NÚM.CARACT(SUBSTITUIR(V" + i + ";\",\";\"\")))&SUBSTITUIR(V" + i + ";\",\";);ESQUERDA(CONCATENAR(W" + i + ";REPT(\" \";3)" +
                            ");3);ESQUERDA(CONCATENAR(X" + i + ";REPT(\" \";15));15);Y" + i + ";ESQUERDA(CONCATENAR(Z" + i + ";REPT(\" \";1));1);ESQUERDA(CONCATENAR(AA" + i + ";REPT(REPT(\" \";45));45);ESQUERDA(CONCATENAR(AB" + i + ";REPT(" +
                            "\" \";50));50);ESQUERDA(CONCATENAR(AC" + i + ";REPT(\" \";2));2);ESQUERDA(CONCATENAR(AD" + i + ";REPT(\" \";5));5);REPT(0;9-NÚM.CARACT(SUBSTITUIR(AE" + i + ";\",\";\"\")))&SUBSTITUIR(AE" + i + ";\",\";);REPT(0;8-NÚM.CARACT" +
                            "(SUBSTITUIR(AF" + i + ";\",\";\"\")))&SUBSTITUIR(AF" + i + ";\",\";);ESQUERDA(CONCATENAR(AG" + i + ";REPT(\" \";14));14))";
                    }
                }
                i++;
            }
        }

        private int GerarBlocoR150(ExcelPackage excelPackage, out int i, int num, out string[] titulos)
        {
            // Aqui simplesmente adiciono a planilha inicial
            var sheet3 = excelPackage.Workbook.Worksheets.Add("0150");
            sheet3.Name = "0150";

            // Títulos
            i = 1;
            titulos = new String[] { "Registro", "Código", "Razão Social", "Código Pais", "CNPJ" ,
                "CPF", "IE", "Municipio", "SUFRAMA", "ENDEREÇO", "Numero", "Complemento", "Bairro" };
            foreach (var titulo in titulos)
            {
                sheet3.Cells[1, i++].Value = titulo;
            }

            i = 2;
            //int arrayTotal;
            foreach (string y in R150)
            {
                num = 0;
                string[] value = y.Split('|');//.Where(x => x != "");
                for (int j = 0; j < value.Length; j++)
                {
                    if (!value[num].Equals(""))
                        sheet3.Cells[i, j].Value = value[num];

                    num++;
                }
                i++;
            }

            return num;
        }

        private int GerarBlocoR200(ExcelPackage excelPackage, out int i, int num, out string[] titulos)
        {
            // Aqui simplesmente adiciono a planilha inicial
            var sheet2 = excelPackage.Workbook.Worksheets.Add("0200");
            sheet2.Name = "0200";

            // Títulos
            i = 1;
            titulos = new String[] { "Código", "", "", "", "" ,
                "", "NCM", "", "", "", "", "Reg" };
            foreach (var titulo in titulos)
            {
                sheet2.Cells[1, i++].Value = titulo;
            }

            i = 2;
            foreach (string y in R200)
            {
                num = 0;
                string[] value = y.Split('|');//.Where(x => x != "");
                for (int j = 1; j < value.Count() + 1; j++)
                {
                    if (!value[num].Equals(""))
                    {
                        if (value[num].Equals("0200"))
                        {
                            sheet2.Cells[i, 12].Value = value[num];
                        }

                        if ((value[num].Length == 1 || value[num].Length == 2 ||
                            value[num].Length == 3) && !Regex.IsMatch(value[num], @"^[0-9]+$"))
                        {
                            sheet2.Cells[i, 5].Value = value[num];
                        }

                        if (value[num].Length == 2 && Regex.IsMatch(value[num], @"^[0-9]+$"))
                        {
                            sheet2.Cells[i, 5].Value = value[num];
                        }
                        if (value[num].Length >= 7)
                        {
                            if (value[num].Length == 8 && Regex.IsMatch(value[num], @"^[0-9]+$"))
                            {
                                sheet2.Cells[i, 6].Value = value[num];
                            }
                            else
                            {
                                string[] count = value[num].Split(' ');
                                if (count.Count() == 1)
                                {
                                    sheet2.Cells[i, 1].Value = value[num];
                                }
                                else
                                {
                                    sheet2.Cells[i, 2].Value = value[num];
                                }
                            }
                        }
                    }
                    num++;
                }
                i++;
            }

            return num;
        }

        private int GerarBlocoA(ExcelPackage excelPackage, out int i, int num, out string[] titulos)
        {
            // Aqui simplesmente adiciono a planilha inicial
            var sheet1 = excelPackage.Workbook.Worksheets.Add("Bloco A");
            sheet1.Name = "Bloco A";

            // Títulos
            i = 1;
            titulos = new String[] { "Código", "Código Geral", "REG", "IND_OPER", "IND_EMIT" ,
                "COd_PART", "COD_SIT", "SER", "SUB", "NUM_DOC", "CHV_NFSE",
                "DT_DOC", "DT_EXE_SERV", "VL_DOC", "IND_PGTO", "VL_DESC", "VL_BC_PIS",
                "VL_PIS", "VL_BC_CONFINS", "VL_PIS_RET", "VL_CONFINS_RE", "VL_ISS" };

            foreach (var titulo in titulos)
            {
                sheet1.Cells[1, i++].Value = titulo;
            }

            i = 2;
            foreach (string y in A100)
            {
                num = 0;
                string[] value = y.Split('|');//.Where(x => x != "");
                for (int j = 2; j <= value.Count(); j++)
                {
                    if (!value[num].Equals(""))
                    {
                        sheet1.Cells[i, j].Value = value[num];
                    }
                    num++;
                }

                sheet1.Cells[i, 1].Value = "=SE(C" + i + "=\"A100\";CONCATENAR(E" + i + ";D" + i + ";SE(H" + i + "=\"\"; \" \");(TEXTO(J" + i + "\"000000000\"));L" + i + ";F" + i + ");\"\")";
                sheet1.Cells[i, 2].Value = "=SE(A" + i + "= \"\";B" + (i - 1) + ";A" + i + ")";

                i++;
            }

            return num;
        }

        private void GerarBlocoC(ExcelPackage excelPackage, out int i, out string[] titulos, out int num)
        {
            // Aqui simplesmente adiciono a planilha inicial
            var sheet = excelPackage.Workbook.Worksheets.Add("Bloco C");
            sheet.Name = "Bloco C";

            // Títulos
            i = 1;
            titulos = new String[] { "Código", "Código", "1 - Registro", "2 - OPERAÇÃO", "3 - TIPO EMITENTE" ,
                "4 - PARTICIPANTE", "5 - MODELO NFE", "6 - SITUAÇÃO TRIBUTARIA", "7 - S**ERIE", "8 - NUMERO", "9 - CHAVE NFE",
                "10 - DT.EMISSÃO", "11 - DT.SAIDA", "12 - VL.TOTAL", "13 - TP.PAGTO", "14 - VL.DESCONTO", "15 - ABATIMENTOS ZFM",
                "16 - VL.MERCADORIA", "17 - FRETE", "18 - VL.FRETE", "19 - VL.SEGURO", "20 - VL.DESP.ACESS", "21 - VL.BASE ICMS",
                "22 - VL ICMS", "23 - VL.BASE.ICMS.ST", "24 - VL.ICMS.ST", "25 - VALOR IPI", "26 - VALOR PIS", "27 - VALOR COFINS",
                "28 - VL.PIS.ST", "29 - VL.COFINS.ST", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42"};
            foreach (var titulo in titulos)
            {
                sheet.Cells[1, i++].Value = titulo;
            }

            // Valores
            i = 2;
            num = 0;
            foreach (string y in C100)
            {
                num = 0;
                string[] value = y.Split('|');//.Where(x => x != "");
                for (int j = 2; j < value.Count(); j++)
                {
                    if (!value[num].Equals(""))
                    {
                        sheet.Cells[i, j].Value = value[num];
                    }
                    num++;
                }

                sheet.Cells[i, 1].Value = "=SE(C" + i + "=\"C100\";CONCATENAR(E" + i + ";D" + i + ";G" + i + ";(TEXTO(J" + i + ";\"000000000\"));SE(I" + i + "=\"\";\" \";TEXTO(I" + i + ";\"000\"));L" + i + ";M" + i + ";F" + i + ");\"\")";
                sheet.Cells[i, 2].Value = "=SE(A" + i + "=\"\";B" + i + ";A" + i + ")";

                i++;
            }
        }

        private string CarregaListaDados(ref int counter)
        {
            string line;
            // Read the file and display it line by line.  
            System.IO.StreamReader file =
                new System.IO.StreamReader(openFileDialog1.FileName);

            while ((line = file.ReadLine()) != null)
            {
                string[] inicioPalavra = line.Split('|');
                foreach (var X in inicioPalavra)
                {
                    if (X.Equals("C100") || X.Equals("C170") || X.Equals("C113") || X.Equals("C120"))
                    {
                        C100.Add(line);
                    }

                    if (X.Equals("0150"))
                    {
                        R150.Add(line);
                    }

                    if (X.Equals("A100") || X.Equals("A170") || X.Equals("A120"))
                    {
                        A100.Add(line);
                    }

                    if (X.Equals("0200"))
                    {
                        R200.Add(line);
                    }

                    if (X.Equals("C113"))
                    {
                        C113.Add(line);
                    }

                    if (X.Equals("1100") || X.Equals("1105"))
                    {
                        R1100.Add(line);
                    }
                }
                counter++;
            }

            if (txtCarga.Text != "")
            {
                System.IO.StreamReader reader =
                    new System.IO.StreamReader(openFileDialog2.FileName);

                while ((line = reader.ReadLine()) != null)
                {
                    string[] inicioPalavra = C100.ToArray();
                    if (inicioPalavra != null)
                    {
                        foreach (var X in inicioPalavra)
                        {
                            if (line.Equals(X))
                                C100.Add(line);
                        }
                    }
                    else
                    {
                        string[] verificaPalavra = line.Split('|');
                        foreach (var X in inicioPalavra)
                        {
                            if (X.Equals("C100") || X.Equals("C170") || X.Equals("C113") || X.Equals("C120"))
                            {
                                C100.Add(line);
                            }
                        }
                    }

                    inicioPalavra = A100.ToArray();
                    if (inicioPalavra != null)
                    {
                        foreach (var X in inicioPalavra)
                        {
                            if (line.Equals(X))
                                A100.Add(line);
                        }
                    }
                    else
                    {
                        string[] verificaPalavra = line.Split('|');
                        foreach (var X in inicioPalavra)
                        {
                            if (X.Equals("A100") || X.Equals("A170") || X.Equals("A120"))
                            {
                                A100.Add(line);
                            }
                        }
                    }

                    inicioPalavra = R150.ToArray();
                    if (inicioPalavra != null)
                    {
                        foreach (var X in inicioPalavra)
                        {
                            if (line.Equals(X))
                                R150.Add(line);
                        }
                    }
                    else
                    {
                        string[] verificaPalavra = line.Split('|');
                        foreach (var X in inicioPalavra)
                        {
                            if (X.Equals("0150"))
                            {
                                R150.Add(line);
                            }
                        }
                    }


                    inicioPalavra = R200.ToArray();
                    if (inicioPalavra != null)
                    {
                        foreach (var X in inicioPalavra)
                        {
                            if (line.Equals(X))
                                R200.Add(line);
                        }
                    }
                    else
                    {
                        string[] verificaPalavra = line.Split('|');
                        foreach (var X in inicioPalavra)
                        {
                            if (X.Equals("0200"))
                            {
                                R200.Add(line);
                            }
                        }
                    }

                    inicioPalavra = C113.ToArray();
                    if (inicioPalavra != null)
                    {
                        foreach (var X in inicioPalavra)
                        {
                            if (line.Equals(X))
                                C113.Add(line);
                        }
                    }
                    else
                    {
                        string[] verificaPalavra = line.Split('|');
                        foreach (var X in inicioPalavra)
                        {
                            if (X.Equals("C113"))
                            {
                                C113.Add(line);
                            }
                        }
                    }

                    inicioPalavra = R1100.ToArray();
                    if (inicioPalavra != null)
                    {
                        foreach (var X in inicioPalavra)
                        {
                            if (line.Equals(X))
                                R1100.Add(line);
                        }
                    }
                    else
                    {
                        string[] verificaPalavra = line.Split('|');
                        foreach (var X in inicioPalavra)
                        {
                            if (X.Equals("1100") || X.Equals("1105"))
                            {
                                R1100.Add(line);
                            }
                        }
                    }
                }

                reader.Close();
            }
            file.Close();
            return line;
        }

        private void btnSearch1_Click(object sender, EventArgs e)
        {
            //define as propriedades do controle 
            //OpenFileDialog
            this.openFileDialog2.Multiselect = false;
            this.openFileDialog2.Title = "Selecionar Arquivo";
            openFileDialog2.InitialDirectory = @"C:\";
            openFileDialog2.Filter = "Texto(*.txt)|*.txt";
            openFileDialog2.CheckFileExists = true;
            openFileDialog2.CheckPathExists = true;
            openFileDialog2.FilterIndex = 2;
            openFileDialog2.RestoreDirectory = true;
            openFileDialog2.ReadOnlyChecked = true;
            openFileDialog2.ShowReadOnly = true;
            openFileDialog2.DefaultExt = "txt";


            DialogResult dr = this.openFileDialog2.ShowDialog();

            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                txtCarga.Text = openFileDialog2.FileName;
                arquivo1 = openFileDialog2.FileName;
            }
        }
    }
}
