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
using Microsoft.VisualBasic;

namespace In86
{
    public partial class frmIn86 : Form
    {
        private string arquivo;
        private string arquivo1;
        List<string> calcTempC = new List<string>();
        List<string> calcTempB = new List<string>();
        List<string> calcR200 = new List<string>();
        List<string> C100 = new List<string>();
        List<string> A100 = new List<string>();
        List<string> C113 = new List<string>();
        List<string> R200 = new List<string>();
        List<string> R150 = new List<string>();
        List<string> R1100 = new List<string>();
        List<string> R0 = new List<string>();
        List<string> C100Temp = new List<string>();
        public bool carregado = false;
        public bool carregadoPis = false;

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
            }
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnConverter_Click(object sender, EventArgs e)
        {
            if (carregado && carregadoPis)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var excelPackage = new ExcelPackage())
                {
                    excelPackage.Workbook.Properties.Title = "IN86";
                    var sheet = excelPackage.Workbook.Worksheets.Add("Bloco A");

                    int i, num;

                    string caminho, path;

                    float calcFatValRBNCT, calcFatValRBNCNT, calcFatValRBNCE, calcIndRBNCT, calcIndRBNCNT, calcIndRBNCE, calcFatValTotal, calcFacIndTotal, calcPercIndRBNCT, calcPercIndRBNCNT, calcPercIndRBNCE;

                    CalculoFaturamento(out calcFatValRBNCT, out calcFatValRBNCNT, out calcFatValRBNCE, out calcIndRBNCT, out calcIndRBNCNT, out calcIndRBNCE, out calcFatValTotal, out calcFacIndTotal, out calcPercIndRBNCT, out calcPercIndRBNCNT, out calcPercIndRBNCE);

                    GerarBlocoC(sheet, out i, out caminho, out path);

                    GerarBLocoA(sheet, out i, out num, caminho, path);

                    GerarBlocoR0200(sheet, out i, num, out caminho, out path);

                    GerarBlocoR0150(out i, out caminho, out path);

                    GerarBloco431(sheet, out i, num, out caminho, out path);

                    GerarBloco432(sheet, out i, out caminho, out path);

                    GerarBloco433(sheet, out i, out caminho, out path);

                    GerarBloco434(sheet, out i, out caminho, out path);

                    GerarBloco438(sheet, out i, out caminho, out path);

                    GerarBloco439(sheet, out i, out caminho, out path);

                    GerarBloco4101(sheet, out i, out caminho, out path);

                    GerarBLoco4104(sheet, out i, out caminho, out path, calcIndRBNCT, calcIndRBNCNT, calcIndRBNCE);

                    GerarBloco4105(sheet, out i, out caminho, out path, calcIndRBNCT, calcIndRBNCNT, calcIndRBNCE, calcPercIndRBNCT);

                    GerarBloco4106(sheet, out i, out caminho, out path, calcIndRBNCT, calcIndRBNCNT, calcIndRBNCE);

                    GerarBloco1CE(sheet, out i, out num, out caminho, out path);

                    GerarBloco441(sheet, out i, num, out caminho, out path);

                    GerarBloco442(sheet, out i, num, out caminho, out path);

                    MessageBox.Show("Concluído. Verifique em " + Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                }
            }
            else
            {
                MessageBox.Show("Favor carregar os dados para dar inicio ao processo de conversão", "Aviso", MessageBoxButtons.OK);
            }
        }

        private void GerarBloco442(ExcelWorksheet sheet, out int i, int num, out string caminho, out string path)
        {
            // Inicio do bloco 4.4.2
            StreamWriter x;

            caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            path = caminho + @"\Bloco4.4.2.txt";
            x = File.CreateText(path);

            // Títulos
            x.WriteLine("|01 - Modelo|02 - Série / sub|03 - Num docto|4 - Data emissão|5 - Numero DI|Linha preenchida IN86 - 4.4.2|");

            i = 2;
            string calcA, calcB, calcC, calcD, calcE, calcF;
            int countTemp = 0;

            foreach (string y in C100)
            {
                calcA = "";
                calcB = "";
                calcC = "";
                calcD = "";
                calcE = "";
                calcF = "";

                num = 0;

                string[] value = y.Split('|');//.Where(x => x != "");

                CarregaBlocoCSheet40(sheet, i, value);

                if (sheet.Cells[i, 1].Value.ToString() == "C120" && calcTempC[countTemp].Substring(1, 1).ToString() == "0")
                {
                    calcA = calcTempC[countTemp].Substring(2, 2).ToString();
                    calcB = calcTempC[countTemp].Substring(4, 3).ToString();
                    calcC = calcTempC[countTemp].Substring(8, 9).ToString();
                    calcD = calcTempC[countTemp].Substring(17, 8).ToString();
                    calcE = sheet.Cells[i, 3].Value.ToString();
                    countTemp++;

                }
                else
                {
                    calcA = "FALSO";
                    calcB = "FALSO";
                    calcC = "FALSO";
                    calcD = "FALSO";
                    calcE = "FALSO";
                    countTemp++;
                }

                calcF = calcA + calcB.PadLeft(3, ' ') + ' ' + calcC + calcD + calcE.Replace(',', ' ');

                x.WriteLine("|" + calcA + "|" + calcB + "|" + calcC + "|" + calcD + "|" + calcE + "|" + calcF + "|");
                i++;
            }
        }

        private void GerarBloco441(ExcelWorksheet sheet, out int i, int num, out string caminho, out string path)
        {
            // Inicio do bloco 4.4.1
            StreamWriter x;

            caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            path = caminho + @"\Bloco4.4.1.txt";
            x = File.CreateText(path);


            // Títulos
            x.WriteLine("|01 - Modelo|02 - Série / sub|03 - Num docto|4 - Data emissão|5 - Numero do registro|6 - Numero do despacho|Linha preenchida IN86 - 4.4.1|");

            i = 2;
            string calcA, calcB, calcC, calcD, calcE, calcF, calcG;

            foreach (string y in R1100)
            {
                calcA = "";
                calcB = "";
                calcC = "";
                calcD = "";
                calcE = "";
                calcF = "";
                calcG = "";

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

                if (sheet.Cells[i, 3].Value.ToString() == "1105")
                {
                    calcA = sheet.Cells[i, 4].Value.ToString();
                    calcB = sheet.Cells[i, 5].Value.ToString();
                    calcC = sheet.Cells[i, 6].Value.ToString();
                    calcD = sheet.Cells[i, 8].Value.ToString();
                    calcE = sheet.Cells[i, 7].Value.ToString().Substring(1, 12);
                    calcF = sheet.Cells[i, 7].Value.ToString().Substring(13, 12);
                    calcG = calcA + calcB.PadRight(5, ' ') + calcC.PadLeft(9, ' ') + calcD + calcE + calcF.Replace(',', ' ');
                }
                else
                {
                    calcA = "FALSO";
                    calcB = "FALSO";
                    calcC = "FALSO";
                    calcD = "FALSO";
                    calcE = "FALSO";
                    calcF = "FALSO";
                    calcG = "FALSO";
                }

                x.WriteLine("|" + calcA + "|" + calcB + "|" + calcC + "|" + calcD + "|" + calcE + "|" + calcF + "|" + calcG + "|");
                i++;
            }
        }

        private void GerarBloco1CE(ExcelWorksheet sheet, out int i, out int num, out string caminho, out string path)
        {
            // Inicio do bloco 1 CE
            StreamWriter x;

            caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            path = caminho + @"\Bloco1CE.txt";
            x = File.CreateText(path);

            // Títulos
            x.WriteLine("|Nr registro|Nr despacho|||registro|||03 -NRO_DE|||06 - NRO_RE|");

            // Valores
            i = 2;
            num = 0;
            sheet.Cells.Clear();
            string calcA, calcB, calcC;
            int countTemp = 0;

            foreach (string y in R1100)
            {
                calcA = "";
                calcB = "";
                calcC = "";
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

                if (sheet.Cells[i, 3].Value.ToString() != "0150")
                {
                    if (sheet.Cells[i, 3].Value.ToString() == "1100")
                    {
                        calcA = sheet.Cells[i, 6].Value.ToString();
                        calcB = sheet.Cells[i, 3].Value.ToString();
                    }
                    else
                    {
                        calcA = "";
                        calcB = "";
                    }

                    if (calcA != "" & calcB != "")
                    {
                        calcC = calcA + calcB;
                        sheet.Cells[i, 13].Value = calcC.ToString();
                        countTemp = i;
                    }
                    else
                    {
                        calcC = sheet.Cells[countTemp, 13].Value.ToString();
                    }

                    x.WriteLine("|" + calcA + "|" + calcB + "|" + calcC + "|" + y);
                }
                i++;
            }
        }

        private void GerarBloco4106(ExcelWorksheet sheet, out int i, out string caminho, out string path, float calcIndRBNCT, float calcIndRBNCNT, float calcIndRBNCE)
        {
            // Inicio do bloco 4.10.6
            StreamWriter x;

            caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            path = caminho + @"\Bloco4.10.6.txt";
            x = File.CreateText(path);

            // Títulos
            x.WriteLine("|1 - Série|2 - Nr docto|3 - DT Emissão|4 - Participante|5 - Nr item|6 - CST PIS|" +
                "7 - Alíquota|8 - Base Calc|09 - Vlr Crédito PIS - Receita Exportação|10 - Vlr Crédito PIS - Receita Mercado interno|" +
                "11 - Vlr Crédito PIS - Receita não tributada|12 - Vlr PIS|13 - CST COFINS|14 - Alíq Cofins|15 - BC Cofins|" +
                "16 - Vlr Créd Cofins Receita Exportação|17 - Vlr Créd Cofins - Receita Mercado interno|18 - Vlr Créd Cofins Receita não tributada|" +
                "19 - Valor Cofins|20 - Dt Apropriação|Linha preenchida IN25/10 - 4.10.6|");

            i = 2;
            string calcA, calcB, calcC, calcD, calcE, calcF, calcG, calcH, calcI, calcJ, calcK, calcL, calcM, calcN, calcO, calcP, calcQ, calcR, calcS, calcT, calcU;
            int countTemp = 0, countSave = 0, usualCount = 0;
            bool savedCount = false;

            foreach (string y in A100)
            {
                calcA = "";
                calcB = "";
                calcC = "";
                calcD = "";
                calcE = "";
                calcF = "";
                calcG = "";
                calcH = "";
                calcI = "";
                calcJ = "";
                calcK = "";
                calcL = "";
                calcM = "";
                calcN = "";
                calcO = "";
                calcP = "";
                calcQ = "";
                calcR = "";
                calcS = "";
                calcT = "";
                calcU = "";

                string[] value = y.Split('|');//.Where(x => x != "");
                CarregaBlocoA(sheet, i, value);

                if (sheet.Cells[i, 1].Value.ToString() == "A170")
                {
                    calcA = calcTempB[countTemp].ToString().Substring(3, 3).Replace(',', ' ');
                    calcB = calcTempB[countTemp].Substring(6, 9).ToString();
                    calcC = calcTempB[countTemp].Substring(15, 8).ToString();
                    calcD = calcTempB[countTemp].Substring(19, 7).ToString();
                    calcE = sheet.Cells[i, 2].Value.ToString().PadLeft(3, '0');
                    calcF = sheet.Cells[i, 9].Value.ToString();
                    calcG = string.Format(@"{0:0,0000}", sheet.Cells[i, 11].Value.ToString());
                    calcH = string.Format(@"{0:0,000}", sheet.Cells[i, 10].Value.ToString());
                    calcI = string.Format(@"{0:f}", float.Parse(sheet.Cells[i, 10].Value.ToString()) * calcIndRBNCE);
                    calcJ = string.Format(@"{0:f}", float.Parse(sheet.Cells[i, 10].Value.ToString()) * calcIndRBNCT);
                    calcK = string.Format(@"{0:f}", float.Parse(sheet.Cells[i, 10].Value.ToString()) * calcIndRBNCNT);
                    calcL = string.Format(@"{0:f}", sheet.Cells[i, 12].Value.ToString());
                    calcN = string.Format(@"{0:0,0000}", sheet.Cells[i, 15].Value.ToString());
                    calcS = string.Format(@"{0:f}", sheet.Cells[i, 16].Value.ToString());
                    calcP = string.Format(@"{0:f}", float.Parse(calcS) * calcIndRBNCE);
                    calcQ = string.Format(@"{0:f}", float.Parse(calcS) * calcIndRBNCT);
                    calcR = string.Format(@"{0:f}", float.Parse(calcS) * calcIndRBNCNT);
                    calcT = calcTempB[countTemp].Substring(15, 8);
                    savedCount = false;
                }
                else
                {
                    calcA = "FALSO";
                    calcB = "FALSO";
                    calcC = "FALSO";
                    calcD = "FALSO";
                    calcE = "FALSO";
                    calcF = "FALSO";
                    calcG = "FALSO";
                    calcH = "FALSO";
                    calcI = "FALSO";
                    calcJ = "FALSO";
                    calcK = "FALSO";
                    calcL = "FALSO";
                    calcN = "FALSO";
                    calcP = "FALSO";
                    calcQ = "FALSO";
                    calcR = "FALSO";
                    calcS = "FALSO";
                    calcT = "FALSO";
                }
                countTemp = usualCount + 1;
                usualCount = countTemp;

                if (countTemp != calcTempB.Count())
                {
                    foreach (var ind in calcTempB)
                    {//estorou o numero de casas do array
                        if (calcTempB[countTemp].ToString() == "")
                        {
                            countTemp = countTemp - 1;
                        }
                        else
                        {
                            if (!savedCount)
                            {
                                countSave = countTemp;
                                savedCount = true;
                                break;
                            }
                        }
                    }
                }

                calcM = calcF;
                calcO = calcH;

                calcU = calcA.PadLeft(5, ' ') + calcB.PadLeft(9, ' ') + calcC + calcD.PadLeft(14, ' ') + calcE + calcF.PadLeft(2, ' ') + calcG.Replace(',', ' ') + calcH.Replace(',', ' ') +
                    calcI.Replace(',', ' ') + calcJ.Replace(',', ' ') + calcK.Replace(',', ' ') + calcL.Replace(',', ' ') + calcM.PadLeft(2, ' ') + calcN.Replace(',', ' ') + calcO.Replace(',', ' ') +
                    calcP.Replace(',', ' ') + calcQ.Replace(',', ' ') + calcR.Replace(',', ' ') + calcS.Replace(',', ' ') + calcT;



                x.WriteLine("|" + calcA + "|" + calcB + "|" + calcC + "|" + calcD + "|" + calcE + "|" + calcF + "|" + calcG + "|" + calcH + "|" + calcI + "|" + calcJ + "|" + calcK + "|" + calcL + "|" + calcM + "|" + calcN + "|" + calcO + "|" + calcP + "|" + calcQ + "|" + calcR + "|" + calcS + "|" + calcT + "|" + calcU + "|");
                i++;
            }
        }

        private void GerarBloco4105(ExcelWorksheet sheet, out int i, out string caminho, out string path, float calcIndRBNCT, float calcIndRBNCNT, float calcIndRBNCE, float calcPercIndRBNCT)
        {
            // Inicio bloco 4.10.5
            StreamWriter x;

            caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            path = caminho + @"\Bloco4.10.5.txt";
            x = File.CreateText(path);

            // Títulos
            x.WriteLine("|1 - Modelo docto|2 - Série|3 - Num do dcto|4 - Dt Emissão|5 - Cod Participante|" +
            "6 - Nr item|7 - CST PIS|8 - Alíquota|9 - Base Calc|10 - Vlr Crédito PIS - Receita Exportação|11 - Vlr Crédito PIS - Receita Mercado interno|" +
            "12 - Vlr Crédito PIS - Receita não tributada|13 - Vlr PIS|14 - CST COFINS|15 - Alíq Cofins|16 - BC Cofins|17 - Vlr Créd Cofins Receita Exportação|" +
            "18 - Vlr Créd Cofins - Receita Mercado interno|19 - Vlr Créd Cofins Receita não tributada|20 - Valor Cofins|21 - Dt Apropriação|Linha preenchida IN25/10 - 4.10.5|");

            i = 2;
            string calcA, calcB, calcC, calcD, calcE, calcF, calcG, calcH, calcI, calcJ, calcK, calcL, calcM, calcN, calcO, calcP, calcQ, calcR, calcS, calcT, calcU, calcV;
            int countTemp;

            foreach (string y in C100)
            {

                calcA = "";
                calcB = "";
                calcC = "";
                calcD = "";
                calcE = "";
                calcF = "";
                calcG = "";
                calcH = "";
                calcI = "";
                calcJ = "";
                calcL = "";
                calcK = "";
                calcM = "";
                calcN = "";
                calcO = "";
                calcP = "";
                calcQ = "";
                calcR = "";
                calcS = "";
                calcT = "";
                calcU = "";
                calcV = "";

                string[] value = y.Split('|');//.Where(x => x != "");

                CarregaBlocoCSheet40(sheet, i, value);

                countTemp = 0;
                if (value.Length >= 30)
                {
                    if (calcTempC[countTemp].ToString().Substring(0, 1) == "1" && sheet.Cells[i, 1].Value.ToString() == "C170")
                    {
                        calcA = calcTempC[countTemp].Substring(2, 2);
                        calcB = calcTempC[countTemp].Substring(14, 3);
                        calcC = calcTempC[countTemp].Substring(5, 9);
                        calcD = calcTempC[countTemp].Substring(14, 8);
                        calcE = calcTempC[countTemp].Substring(24, 8);
                        calcF = sheet.Cells[i, 2].Value.ToString().PadLeft(3, '0');

                        if (sheet.Cells[i, 25].Value.ToString() == "")
                        {
                            calcG = "";
                        }
                        else
                        {
                            calcG = sheet.Cells[i, 25].Value.ToString();
                        }

                        calcH = string.Format(@"{0:0,0000}", (sheet.Cells[i, 27].Value ?? "0,0000").ToString());
                        calcI = string.Format(@"{0:0,0000}", (sheet.Cells[i, 26].Value ?? "0,0000").ToString());
                        calcM = string.Format(@"{0:f}", (sheet.Cells[i, 30].Value ?? "0,00").ToString());
                        calcO = string.Format(@"{0:f}", sheet.Cells[i, 33].Value.ToString());
                        calcT = string.Format(@"{0:f}", (sheet.Cells[i, 36].Value.ToString() == "" ? "0,00" : sheet.Cells[i, 36].Value).ToString());
                        calcU = calcTempC[countTemp].Substring(25, 8);

                    }
                    else
                    {
                        calcA = "FALSO";
                        calcB = "FALSO";
                        calcC = "FALSO";
                        calcD = "FALSO";
                        calcE = "FALSO";
                        calcF = "FALSO";
                        calcG = "FALSO";
                        calcH = "FALSO";
                        calcI = "FALSO";
                        calcM = "FALSO";
                        calcO = "FALSO";
                        calcT = "FALSO";
                        calcU = "FALSO";
                    }

                    calcN = calcG;
                    calcP = calcI;

                    if (calcG == "50")
                    {
                        calcJ = "0,00";
                        calcL = "0,00";
                    }
                    else
                    {
                        if (calcTempC[countTemp].ToString().Substring(0, 1) == "1" && sheet.Cells[i, 1].Value.ToString() == "C170")
                        {
                            if (calcM != "")
                            {
                                calcJ = string.Format(@"{0:f}", float.Parse(calcM) * calcIndRBNCE);
                                calcL = string.Format(@"{0:f}", float.Parse(calcM) * calcIndRBNCNT);
                            }
                            else
                            {
                                calcJ = "0,00";
                                calcL = "0,00";
                            }
                        }
                        else
                        {
                            calcJ = "FALSO";
                            calcL = "FALSO";
                        }
                    }

                    if (calcG == "50")
                    {
                        calcK = calcM;
                    }
                    else
                    {
                        if (calcTempC[countTemp].ToString().Substring(0, 1) == "1" && sheet.Cells[i, 1].Value.ToString() == "C170")
                        {
                            if (calcM != "")
                            {
                                calcK = string.Format(@"{0:f}", float.Parse(calcM) * calcPercIndRBNCT);
                            }
                            else
                            {
                                calcK = "0,00";
                            }
                        }
                        else
                        {
                            calcK = "FALSO";
                        }
                    }

                    if (calcN == "50")
                    {
                        calcQ = "0,00";
                        calcS = "0,00";
                    }
                    else
                    {
                        if (calcTempC[countTemp].ToString().Substring(0, 1) == "1" && sheet.Cells[i, 1].Value.ToString() == "C170")
                        {
                            calcQ = string.Format(@"{0:f}", float.Parse(calcT) * calcIndRBNCE);
                            calcS = string.Format(@"{0:f}", float.Parse(calcT) * calcIndRBNCNT);
                        }
                        else
                        {
                            calcQ = "FALSO";
                            calcS = "FALSO";
                        }
                    }

                    if (calcN == "50")
                    {
                        calcR = calcT;
                    }
                    else
                    {
                        if (calcTempC[countTemp].ToString().Substring(0, 1) == "1" && sheet.Cells[i, 1].Value.ToString() == "C170")
                        {
                            calcR = string.Format(@"{0:f}", float.Parse(calcT) * calcIndRBNCT);
                        }
                        else
                        {
                            calcR = "FALSO";
                        }
                    }

                    if (sheet.Cells[i, 1].Value.ToString() == "C170")
                    {
                        countTemp++;
                    }

                    calcV = calcA + calcB.PadLeft(5, ' ') + calcC + calcD + calcE.PadLeft(14, ' ') + calcF + calcG.PadLeft(2, ' ') + calcH.Replace(',', ' ') + calcI.Replace(',', ' ') +
                        calcJ.Replace(',', ' ') + calcK.Replace(',', ' ') + calcL.Replace(',', ' ') + calcM.Replace(',', ' ') + calcN.Replace(',', ' ') + calcO.Replace(',', ' ') + calcP.Replace(',', ' ') +
                        calcQ.Replace(',', ' ') + calcR.Replace(',', ' ') + calcS.Replace(',', ' ') + calcT.Replace(',', ' ') + calcU;

                    x.WriteLine("|" + calcA + "|" + calcB + "|" + calcC + "|" + calcD + "|" + calcE + "|" + calcF + "|" + calcG + "|" + calcH + "|" + calcI + "|" + calcJ + "|" + calcK + "|" + calcL + "|" + calcM + "|" + calcN + "|" + calcO + "|" + calcP + "|" + calcQ + "|" + calcR + "|" + calcS + "|" + calcT + "|" + calcU + "|" + calcV + "|");

                    i++;
                }
            }
        }

        private void GerarBLoco4104(ExcelWorksheet sheet, out int i, out string caminho, out string path, float calcIndRBNCT, float calcIndRBNCNT, float calcIndRBNCE)
        {
            // Inicio do bloco 4.10.4
            StreamWriter x;

            caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            path = caminho + @"\Bloco4.10.4.txt";
            x = File.CreateText(path);

            // Títulos
            x.WriteLine("|1 - Modelo docto|2 - Série|3 - Num do dcto|4 - Dt Emissão|5 - Nr item|" +
            "6 - CST PIS|7 - Alíquota|8 - Base Calc|9 - Vlr Crédito PIS - Receita Exportação|10 - Vlr Crédito PIS - Receita Mercado interno|11 - Vlr Crédito PIS - Receita não tributada|" +
            "12 - Vlr PIS|13 - CST COFINS|14 - Alíq Cofins|15 - BC Cofins|16 - Vlr Créd Cofins Receita Exportação|17 - Vlr Créd Cofins - Receita Mercado interno|" +
            "18 - Vlr Créd Cofins Receita não tributada|19 - Valor Cofins|20 - Dt Apropriação|Linha preenchida IN25/10 - 4.10.4|");

            i = 2;
            string calcA, calcB, calcC, calcD, calcE, calcF, calcG, calcH, calcI, calcJ, calcK, calcL, calcM, calcN, calcO, calcP, calcQ, calcR, calcS, calcT, calcU;
            int countTemp;

            foreach (string y in C100)
            {

                calcA = "";
                calcB = "";
                calcC = "";
                calcD = "";
                calcE = "";
                calcF = "";
                calcG = "";
                calcH = "";
                calcI = "";
                calcJ = "";
                calcL = "";
                calcK = "";
                calcM = "";
                calcN = "";
                calcO = "";
                calcP = "";
                calcQ = "";
                calcR = "";
                calcS = "";
                calcT = "";
                calcU = "";

                string[] value = y.Split('|');//.Where(x => x != "");

                CarregaBlocoCSheet40(sheet, i, value);

                countTemp = 0;
                if (value.Length >= 30)
                {
                    if (calcTempC[countTemp].ToString().Substring(0, 1) == "1" && sheet.Cells[i, 1].Value.ToString() == "C170")
                    {
                        calcA = calcTempC[countTemp].Substring(2, 2);
                        calcB = calcTempC[countTemp].Substring(14, 3);
                        calcC = calcTempC[countTemp].Substring(5, 9);
                        calcD = calcTempC[countTemp].Substring(14, 8);
                        calcE = string.IsNullOrEmpty(sheet.Cells[i, 2].Value.ToString()) ? sheet.Cells[i, 2].Value.ToString().PadLeft(3, '0') : "000";
                        calcF = sheet.Cells[i, 25].Value.ToString();
                        calcG = string.Format(@"{0:0,0000}", (sheet.Cells[i, 27].Value ?? "0,0000").ToString());
                        calcH = string.Format(@"{0:0,0000}", (sheet.Cells[i, 26].Value ?? "0,0000").ToString());
                    }
                    else
                    {
                        calcA = "FALSO";
                        calcB = "FALSO";
                        calcC = "FALSO";
                        calcD = "FALSO";
                        calcE = "FALSO";
                        calcF = "FALSO";
                        calcG = "FALSO";
                        calcH = "FALSO";
                    }

                    calcM = calcF;
                    calcO = calcH;

                    if (calcTempC[countTemp].ToString().Substring(0, 1) == "0" && sheet.Cells[i, 1].Value.ToString() == "C170"
                        && calcTempC[countTemp].ToString().Substring(1, 1) == "0")
                    {
                        calcL = string.Format(@"{0,f}", sheet.Cells[i, 30].Value.ToString());
                        calcN = string.Format(@"{0,0,0000}", sheet.Cells[i, 33].Value.ToString());
                        calcS = string.Format(@"{0,f}", sheet.Cells[i, 36].Value.ToString());
                        calcT = calcTempC[countTemp].Substring(25, 8);

                    }
                    else
                    {
                        calcL = "FALSO";
                        calcN = "FALSO";
                        calcS = "FALSO";
                        calcT = "FALSO";
                    }

                    if (calcF == "50")
                    {
                        calcI = "0,00";
                        calcK = "0,00";
                    }
                    else
                    {
                        if (calcTempC[countTemp].ToString().Substring(0, 1) == "0" && sheet.Cells[i, 1].Value.ToString() == "C170"
                            && calcTempC[countTemp].ToString().Substring(1, 1) == "0")
                        {

                            calcI = string.Format(@"{0:f}", float.Parse(calcL) * calcIndRBNCE);
                        }
                        else
                        {
                            calcI = "FALSO";
                        }

                        if (calcF == "50")
                        {
                            calcK = "0,00";
                        }
                        else
                        {
                            if (calcTempC[countTemp].ToString().Substring(0, 1) == "0" && sheet.Cells[i, 1].Value.ToString() == "C170"
                            && calcTempC[countTemp].ToString().Substring(1, 1) == "0")
                            {
                                calcK = string.Format(@"{0:f}", float.Parse(calcL) * calcIndRBNCNT);
                            }
                            else
                            {
                                calcK = "FALSO";
                            }
                        }

                    }

                    if (calcF == "50")
                    {
                        calcJ = calcL;
                    }
                    else
                    {
                        if (calcTempC[countTemp].ToString().Substring(0, 1) == "0" && sheet.Cells[i, 1].Value.ToString() == "C170"
                            && calcTempC[countTemp].ToString().Substring(1, 1) == "0")
                        {
                            calcJ = string.Format(@"{0:f}", float.Parse(calcL) * calcIndRBNCT);
                        }
                        else
                        {
                            calcJ = "FALSO";
                        }
                    }

                    if (calcM == "50")
                    {
                        calcP = "0,00";
                        calcR = "0,00";
                    }
                    else
                    {
                        if (calcTempC[countTemp].ToString().Substring(0, 1) == "0" && sheet.Cells[i, 1].Value.ToString() == "C170"
                            && calcTempC[countTemp].ToString().Substring(1, 1) == "0")
                        {
                            calcP = string.Format(@"{0:f}", float.Parse(calcS) * calcIndRBNCE);
                            calcR = string.Format(@"{0:f}", float.Parse(calcS) * calcIndRBNCNT);
                        }
                        else
                        {
                            calcP = "FALSO";
                            calcR = "FALSO";
                        }
                    }

                    if (calcM == "50")
                    {
                        calcQ = calcS;
                    }
                    else
                    {
                        if (calcTempC[countTemp].ToString().Substring(0, 1) == "0" && sheet.Cells[i, 1].Value.ToString() == "C170"
                            && calcTempC[countTemp].ToString().Substring(1, 1) == "0")
                        {
                            calcQ = string.Format(@"{0:f}", float.Parse(calcS) * calcIndRBNCT);
                        }
                        else
                        {
                            calcQ = "FALSO";
                        }
                    }
                }

                if (sheet.Cells[i, 1].Value.ToString() == "C170")
                {
                    countTemp++;
                }

                calcU = calcA + calcB.PadLeft(5, ' ') + calcC + calcD + calcE + calcF.PadLeft(2, ' ') + calcG.Replace(',', ' ') + calcH.Replace(',', ' ') + calcI.Replace(',', ' ') +
                    calcJ.Replace(',', ' ') + calcK.Replace(',', ' ') + calcL.Replace(',', ' ') + calcM.PadLeft(2, ' ') + calcN.Replace(',', ' ') + calcO.Replace(',', ' ') + calcP.Replace(',', ' ') +
                    calcQ.Replace(',', ' ') + calcR.Replace(',', ' ') + calcS.Replace(',', ' ') + calcT;

                x.WriteLine("|" + calcA + "|" + calcB + "|" + calcC + "|" + calcD + "|" + calcE + "|" + calcF + "|" + calcG + "|" + calcH + "|" + calcI + "|" + calcJ + "|" + calcK + "|" + calcL + "|" + calcM + "|" + calcN + "|" + calcO + "|" + calcP + "|" + calcQ + "|" + calcR + "|" + calcS + "|" + calcT + "|" + calcU + "|");

                i++;
            }
        }

        private static void CalculoFaturamento(out float calcFatValRBNCT, out float calcFatValRBNCNT, out float calcFatValRBNCE, out float calcIndRBNCT, out float calcIndRBNCNT, out float calcIndRBNCE, out float calcFatValTotal, out float calcFacIndTotal, out float calcPercIndRBNCT, out float calcPercIndRBNCNT, out float calcPercIndRBNCE)
        {
            calcFatValRBNCT = float.Parse(Interaction.InputBox("Receita Bruta Não Cumulativa Tributada Mercado Interno", "Preencher com informações do Faturamento", "Valores R$"));
            calcFatValRBNCNT = float.Parse(Interaction.InputBox("Receita Bruta Não Cumulativa Não Tributada Mercado Interno", "Preencher com informações do Faturamento", "Valores R$"));
            calcFatValRBNCE = float.Parse(Interaction.InputBox("Receita Bruta Não Cumulativa Não Exportação", "Preencher com informações do Faturamento", "Valores R$"));

            calcFatValTotal = calcFatValRBNCE + calcFatValRBNCNT + calcFatValRBNCT;

            calcIndRBNCT = calcFatValRBNCT / calcFatValTotal;

            calcIndRBNCNT = calcFatValRBNCNT / calcFatValTotal;

            calcIndRBNCE = calcFatValRBNCE / calcFatValTotal;

            calcFacIndTotal = calcIndRBNCE + calcIndRBNCNT + calcIndRBNCT;

            calcPercIndRBNCT = calcIndRBNCT * 100;
            calcPercIndRBNCNT = calcIndRBNCNT * 100;
            calcPercIndRBNCE = calcIndRBNCE * 100;
        }

        private void GerarBloco4101(ExcelWorksheet sheet, out int i, out string caminho, out string path)
        {
            // Inicio do bloco 4.10.1
            StreamWriter x;

            caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            path = caminho + @"\Bloco4.10.1.txt";
            x = File.CreateText(path);

            x.WriteLine("|1 - Modelo docto|2 - Série|3 - Num do dcto|4 - Dt Emissão|5 - Nr item|" +
            "6 - CST PIS|7 - Alíquota|8 - Base Calc|9 - Vlr PIS|10 - CST Cofins|11 - Alíq Cofins|" +
            "12 - BC Cofins|13 - Valor Cofins|14 - Dt Apropriação|Linha preenchida IN25/10 - 4.10.1|");

            i = 2;
            string calcA, calcB, calcC, calcD, calcE, calcF, calcG, calcH, calcI, calcJ, calcK, calcL, calcM, calcN, calcO;
            int countTemp;

            foreach (string y in C100)
            {

                calcA = "";
                calcB = "";
                calcC = "";
                calcD = "";
                calcE = "";
                calcF = "";
                calcG = "";
                calcH = "";
                calcI = "";
                calcJ = "";
                calcL = "";
                calcK = "";
                calcM = "";
                calcN = "";
                calcO = "";

                string[] value = y.Split('|');//.Where(x => x != "");

                CarregaBlocoCSheet40(sheet, i, value);

                countTemp = 0;

                if (value.Length >= 30)
                {
                    if (calcTempC[countTemp].ToString().Substring(0, 1) == "1" && sheet.Cells[i, 1].Value.ToString() == "C170")
                    {
                        calcA = calcTempC[countTemp].Substring(2, 2);
                        calcB = calcTempC[countTemp].Substring(14, 3);
                        calcC = calcTempC[countTemp].Substring(5, 9);
                        calcD = calcTempC[countTemp].Substring(17, 8);
                        calcE = string.IsNullOrEmpty(sheet.Cells[i, 2].Value.ToString()) ? sheet.Cells[i, 2].Value.ToString().PadLeft(3, '0') : "000";
                        calcF = sheet.Cells[i, 25].Value.ToString();
                        calcG = string.Format(@"{0:0,0000}", (sheet.Cells[i, 27].Value ?? "0,0000").ToString());
                        calcH = string.Format(@"{0:0,0000}", (sheet.Cells[i, 26].Value ?? "0,0000").ToString());
                        calcI = string.Format(@"{0:f}", (sheet.Cells[i, 30].Value ?? "0,00").ToString());
                        calcK = string.Format(@"{0:0,0000}", (sheet.Cells[i, 33].Value ?? "0,0000").ToString());
                        calcM = string.Format(@"{0:f}", (sheet.Cells[i, 36].Value ?? "0,00").ToString());
                        calcN = calcTempC[countTemp].Substring(25, 8);

                    }
                    else
                    {
                        calcA = "FALSO";
                        calcB = "FALSO";
                        calcC = "FALSO";
                        calcD = "FALSO";
                        calcE = "FALSO";
                        calcF = "FALSO";
                        calcG = "FALSO";
                        calcH = "FALSO";
                        calcI = "FALSO";
                        calcK = "FALSO";
                        calcM = "FALSO";
                        calcN = "FALSO";
                    }

                    calcJ = calcF;
                    calcL = calcH;

                    if (sheet.Cells[i, 1].Value.ToString() == "C170")
                    {
                        countTemp++;
                    }

                    calcO = calcA + calcB.PadLeft(5, ' ') + calcC + calcD + calcE + calcF.PadLeft(2, ' ') + calcG.Replace(',', ' ') + calcH.Replace(',', ' ') + calcI.Replace(',', ' ') + calcJ.Replace(',', ' ') +
                        calcK.Replace(',', ' ') + calcL.Replace(',', ' ') + calcM.Replace(',', ' ') + calcN;

                    x.WriteLine("|" + calcA + "|" + calcB + "|" + calcC + "|" + calcD + "|" + calcE + "|" + calcF + "|" + calcG + "|" + calcH + "|" + calcI + "|" + calcJ + "|" + calcK + "|" + calcL + "|" + calcM + "|" + calcN + "|" + calcO + "|");

                    i++;

                }
            }
        }

        private void GerarBloco439(ExcelWorksheet sheet, out int i, out string caminho, out string path)
        {
            //inicio do bloco 4.3.9
            StreamWriter x;

            caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            path = caminho + @"\Bloco4.3.9.txt";
            x = File.CreateText(path);

            // Títulos
            x.WriteLine("|1 - Série|2 - Nr docto|3 - DT Emissão|4 - Participante|5 - Nr item|" +
            "6 - Código Serviço|7 - Descrição compl|8 - Valor do serviço|9 - Desconto|10 - Aliq ISS|11 - Base Calculo ISS|" +
            "12 - VL ISS|Linha preenchida IN25/10 - 4.3.9|");

            i = 2;
            string calcA, calcB, calcC, calcD, calcE, calcF, calcG, calcH, calcI, calcJ, calcK, calcL, calcM;
            int countTemp = 0, countSave = 0, usualCount = 0;
            float calc = 0;
            bool savedCount = false;

            foreach (string y in A100)
            {
                calcA = "";
                calcB = "";
                calcC = "";
                calcD = "";
                calcE = "";
                calcF = "";
                calcG = "";
                calcH = "";
                calcI = "";
                calcJ = "";
                calcK = "";
                calcL = "";
                calcM = "";

                string[] value = y.Split('|');//.Where(x => x != "");
                CarregaBlocoA(sheet, i, value);

                if (sheet.Cells[i, 1].Value.ToString() == "A170")
                {
                    calcA = calcTempB[countTemp].Substring(3, 3);
                    calcB = calcTempB[countTemp].Substring(6, 9);
                    calcC = calcTempB[countTemp].Substring(15, 8);
                    calcD = calcTempB[countTemp].Substring(19, 7);
                    calcE = sheet.Cells[i, 2].Value.ToString().PadLeft(3, '0');
                    calcF = sheet.Cells[i, 3].Value.ToString();
                    calcH = string.Format(@"{0:f}", sheet.Cells[i, 5].Value.ToString());
                    calcI = string.Format(@"{0:f}", sheet.Cells[i, 6].Value.ToString());
                    calcL = string.Format(@"{0:f}", sheet.Cells[i, 21].Value.ToString());
                    savedCount = false;
                }
                else
                {
                    calcA = "FALSO";
                    calcB = "FALSO";
                    calcC = "FALSO";
                    calcD = "FALSO";
                    calcE = "FALSO";
                    calcF = "FALSO";
                    calcH = "FALSO";
                    calcI = "FALSO";
                    calcL = "FALSO";
                }
                countTemp = usualCount + 1;
                usualCount = countTemp;

                if (countTemp != calcTempB.Count())
                {
                    foreach (var ind in calcTempB)
                    {//estorou o numero de casas do array
                        if (calcTempB[countTemp].ToString() == "")
                        {
                            countTemp = countTemp - 1;
                        }
                        else
                        {
                            if (!savedCount)
                            {
                                countSave = countTemp;
                                savedCount = true;
                                break;
                            }
                        }
                    }
                }

                calcK = calcH;

                if (!string.IsNullOrEmpty(calcK) && !string.IsNullOrEmpty(calcL) &&
                    calcK != "FALSO" && calcL != "FALSO")
                {
                    calc = (float.Parse(calcL) / float.Parse(calcK)) * 100;
                    calcJ = string.Format(@"{0:f}", calc);
                }
                else
                {
                    calcJ = "FALSO";
                }

                calcM = calcA.PadLeft(5, ' ') + calcB.PadLeft(9, ' ') + calcC + calcD.PadLeft(14, ' ') + calcE + calcF.PadLeft(20, ' ') + calcG.PadLeft(45, ' ') +
                    calcH.Replace(',', ' ') + calcI.Replace(',', ' ') + calcJ.Replace(',', ' ') + calcK.Replace(',', ' ') + calcL.Replace(',', ' ');

                x.WriteLine("|" + calcA + "|" + calcB + "|" + calcC + "|" + calcD + "|" + calcE + "|" + calcF + "|" + calcG + "|" + calcH + "|" + calcI + "|" + calcJ + "|" + calcK + "|" + calcL + "|" + calcM + "|");
                i++;
            }
        }

        private void GerarBloco438(ExcelWorksheet sheet, out int i, out string caminho, out string path)
        {

            //Inicio do bloco 4.3.8
            StreamWriter x;

            caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            path = caminho + @"\Bloco4.3.8.txt";
            x = File.CreateText(path);

            // Títulos
            x.WriteLine("|1 - Série|2 - Nr docto|3 - DT Emissão|4 - Participante|5 - Valor do serviço|" +
            "6 - Desconto|7 - Aliq IRR|8 - Base Calculo IRRF|9 - VL IRRF|Linha preenchida IN25/10 - 4.3.8|");

            i = 2;
            string calcA, calcB, calcC, calcD, calcE, calcF, calcG, calcH, calcI, calcJ;


            foreach (string y in A100)
            {
                calcA = "";
                calcB = "";
                calcC = "";
                calcD = "";
                calcE = "";
                calcF = "";
                calcG = "";
                calcH = "";
                calcI = "";
                calcJ = "";

                string[] value = y.Split('|');//.Where(x => x != "");
                CarregaBlocoA(sheet, i, value);

                if (sheet.Cells[i, 1].Value.ToString() == "A100")
                {
                    calcA = sheet.Cells[i, 6].Value.ToString().Replace(',', ' ');
                    calcB = sheet.Cells[i, 8].Value.ToString().PadLeft(9, '0');
                    calcC = sheet.Cells[i, 10].Value.ToString();
                    calcD = sheet.Cells[i, 4].Value.ToString();
                    calcE = string.Format(@"{0:f}", sheet.Cells[i, 12].Value.ToString());
                    calcF = string.Format(@"{0:f}", sheet.Cells[i, 14].Value.ToString());
                }
                else
                {
                    calcA = "FALSO";
                    calcB = "FALSO";
                    calcC = "FALSO";
                    calcD = "FALSO";
                    calcE = "FALSO";
                    calcF = "FALSO";
                }
                calcG = "";
                calcH = "";
                calcI = "";

                calcJ = calcA.PadLeft(5, ' ') + calcB + calcC + calcD.PadLeft(14, ' ') + calcE.Replace(',', ' ') + calcF.Replace(',', ' ') +
                    calcG.Replace(',', ' ') + calcH.Replace(',', ' ') + calcI.Replace(',', ' ');

                x.WriteLine("|" + calcA + "|" + calcB + "|" + calcC + "|" + calcD + "|" + calcE + "|" + calcF + "|" + calcG + "|" + calcH + "|" + calcI + "|" + calcJ + "|");

                i++;
            }
        }

        private static void CarregaBlocoA(ExcelWorksheet sheet, int i, string[] value)
        {
            for (int j = 1; j < value.Count(); j++)
            {
                if (j == 1)
                {
                    sheet.Cells[i, 1].Value = value[j];
                }
                if (j == 3)
                {
                    sheet.Cells[i, 3].Value = value[j];
                }
                if (j == 4)//substituindo a mesma linha, corrigir o valor de i
                {
                    sheet.Cells[i, 4].Value = value[j];
                }
                if (j == 5)
                {
                    sheet.Cells[i, 5].Value = value[j];
                }
                if (j == 7)
                {
                    sheet.Cells[i, 7].Value = value[j];
                }
                if (j == 8)
                {
                    sheet.Cells[i, 8].Value = value[j];
                }
                if (j == 10)
                {
                    sheet.Cells[i, 10].Value = value[j];
                }
                if (j == 11)
                {
                    sheet.Cells[i, 11].Value = value[j];
                }
                if (j == 16)
                {
                    sheet.Cells[i, 16].Value = value[j];
                }
                if (j == 14)
                {
                    sheet.Cells[1, 14].Value = value[j];
                }
                if (j == 18)
                {
                    sheet.Cells[i, 18].Value = value[j];
                }
                if (j == 19)
                {
                    sheet.Cells[i, 19].Value = value[j];
                }
                if (j == 20)
                {
                    sheet.Cells[i, 20].Value = value[j];
                }
                if (j == 12)
                {
                    sheet.Cells[i, 12].Value = value[j];
                }
                if (j == 17)
                {
                    sheet.Cells[i, 17].Value = value[j];
                }
                if (j == 6)
                {
                    sheet.Cells[i, 6].Value = value[j];
                }
                if (j == 13)
                {
                    sheet.Cells[i, 13].Value = value[j];
                }
                if (j == 21)
                {
                    sheet.Cells[i, 21].Value = value[j];
                }

            }
        }

        private void GerarBloco434(ExcelWorksheet sheet, out int i, out string caminho, out string path)
        {
            //Inicio do bloco 4.3.4
            StreamWriter x;


            caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            path = caminho + @"\Bloco4.3.4.txt";
            x = File.CreateText(path);

            // Títulos
            x.WriteLine("|1 - Modelo|2 - Série|3 - Nr docto|4 - DT Emissão|5 - Participante|" +
            "6 - Nr item|7 - Cód Merc/Serv|8 - Descrição compl|9 - CFOP|10 - Cod Nat|11 - Clas Fisc Merc|12 - Qtdade|13 - unid|" +
            "14 - Vlr Unit|15 -Vlr Tot Item|16 - Desconto|17 - Ind Trib IPI|18 - Aliq IPI|19 - BC IPI|20 - Vlr IPI|21 - CST ICMS|" +
            "22 - Ind ICMS|23 - Aliq ICMS|24 - BC ICMS|25 - Vlr ICMS Pr|26 - BC ICMS ST|27 -Vlr ICMS ST|28 - Ind Mov|29 - CST IPI|" +
            "Linha preenchida IN86 - 4.3.4|");

            float calc = 0;

            i = 2;
            string calcA, calcB, calcC, calcD, calcE, calcF, calcG, calcH, calcI, calcJ, calcK, calcL, calcM, calcN, calcO, calcP, calcQ, calcR, calcS, calcT, calcU, calcV, calcW, calcX, calcY, calcZ, calcAA, calcAB, calcAC, calcAD;
            int countTemp, count;

            foreach (string y in C100)
            {

                calcA = "";
                calcB = "";
                calcC = "";
                calcD = "";
                calcE = "";
                calcF = "";
                calcG = "";
                calcH = "";
                calcI = "";
                calcJ = "";
                calcL = "";
                calcK = "";
                calcM = "";
                calcN = "";
                calcO = "";
                calcP = "";
                calcQ = "";
                calcR = "";
                calcS = "";
                calcT = "";
                calcU = "";
                calcV = "";
                calcW = "";
                calcX = "";
                calcY = "";
                calcZ = "";
                calcAA = "";
                calcAB = "";
                calcAC = "";
                calcAD = "";

                string[] value = y.Split('|');//.Where(x => x != "");

                CarregaBlocoCSheet40(sheet, i, value);

                countTemp = 0;
                count = 0;

                if (value.Length >= 30)
                {
                    if (calcTempC[countTemp].ToString().Substring(0, 1) == "1" && sheet.Cells[i, 1].Value.ToString() == "C170")
                    {
                        calcA = calcTempC[countTemp].Substring(2, 2);
                        calcB = calcTempC[countTemp].Substring(14, 3);
                        calcC = calcTempC[countTemp].Substring(5, 9);
                        calcD = calcTempC[countTemp].Substring(17, 8);
                        calcE = calcTempC[countTemp].Substring(14, 8);
                        calcF = sheet.Cells[i, 2].Value.ToString().PadLeft(3, '0');
                        calcG = sheet.Cells[i, 3].Value.ToString();

                        if (sheet.Cells[i, 4].Value.ToString() == "")
                        {
                            calcH = "";
                        }
                        else
                        {
                            calcH = sheet.Cells[i, 4].Value.ToString();
                        }

                        calcI = sheet.Cells[i, 11].Value.ToString();
                        calcJ = sheet.Cells[i, 12].Value.ToString();
                        calcL = sheet.Cells[i, 5].Value.ToString();
                        calcM = sheet.Cells[i, 6].Value.ToString();

                        if (!string.IsNullOrEmpty(calcL))
                        {
                            calc = float.Parse(sheet.Cells[i, 7].Value.ToString());
                            calcN = string.Format(@"{0:f}", calc / float.Parse(calcL));
                        }
                        else
                        {
                            calcN = "0,00";
                        }

                        calcO = string.Format(@"{0:f}", sheet.Cells[i, 7].Value.ToString());
                        calcP = string.Format(@"{0:f}", sheet.Cells[i, 8].Value.ToString());
                        calcR = string.Format(@"{0:f}", sheet.Cells[i, 23].Value.ToString());
                        calcS = string.Format(@"{0:f}", sheet.Cells[i, 22].Value.ToString());
                        calcT = string.Format(@"{0:f}", sheet.Cells[i, 24].Value.ToString());
                        calcU = sheet.Cells[i, 10].Value.ToString().PadLeft(3, '0');
                        calcW = string.Format(@"{0:f}", sheet.Cells[i, 14].Value.ToString());
                        calcX = sheet.Cells[i, 13].Value.ToString();
                        calcY = string.Format(@"{0:f}", sheet.Cells[i, 15].Value.ToString());
                        calcZ = string.Format(@"{0:f}", sheet.Cells[i, 16].Value.ToString());
                        calcAA = string.Format(@"{0:f}", sheet.Cells[i, 18].Value.ToString());

                        if (sheet.Cells[i, 9].Value.ToString() == "0")
                        {
                            calcAB = "S";
                        }
                        else
                        {
                            calcAB = "N";
                        }


                    }
                    else
                    {
                        calcA = "FALSO";
                        calcB = "FALSO";
                        calcC = "FALSO";
                        calcD = "FALSO";
                        calcE = "FALSO";
                        calcF = "FALSO";
                        calcG = "FALSO";
                        calcH = "FALSO";
                        calcI = "FALSO";
                        calcJ = "FALSO";
                        calcK = "FALSO";
                        calcL = "FALSO";
                        calcM = "FALSO";
                        calcN = "FALSO";
                        calcO = "FALSO";
                        calcP = "FALSO";
                        calcT = "FALSO";
                        calcR = "FALSO";
                        calcS = "FALSO";
                        calcU = "FALSO";
                        calcW = "0,00";
                        calcX = "FALSO";
                        calcY = "FALSO";
                        calcZ = "FALSO";
                        calcAA = "FALSO";
                        calcAB = "FALSO";

                        if (calcQ == "1")
                        {
                            calcAC = "00";
                        }
                        else
                        {
                            calcAC = "02";
                        }
                    }
                    countTemp++;

                    if (calcG.ToString() != null && calcG.ToString() != "FALSO")
                    {
                        foreach (string g in calcR200)
                        {
                            if (g.Equals(calcG.ToString()))
                            {
                                calcK = g;
                            }

                            if (calcK != "")
                                break;

                            count++;
                        }
                    }
                    else
                    {
                        calcK = "FALSO";
                    }

                    if (calcT == "0,00")
                    {
                        calcQ = "2";
                    }
                    else
                    {
                        calcQ = "1";
                    }

                    if (calcU.Substring(2, 1) == "2" || calcU.Substring(2, 1) == "1" || calcU.Substring(2, 1) == "0")
                    {
                        calcV = "1";
                    }
                    else if (calcU.Substring(2, 1) == "9")
                    {
                        calcV = "3";
                    }
                    else if (calcU.Substring(2, 1) == "7")
                    {
                        calcV = "1";
                    }
                    else
                    {
                        calcV = "2";
                    }

                    if (calcQ == "1")
                    {
                        calcAC = "00";
                    }
                    else
                    {
                        calcAC = "02";
                    }

                    calcAD = calcA + calcB.PadLeft(5, ' ') + calcC + calcD + calcE.PadLeft(14, ' ') + calcF.PadLeft(3, ' ') + calcG.PadLeft(20, ' ') +
                        calcH.PadLeft(45, ' ') + calcI + calcJ.PadLeft(6, ' ') + calcK.PadLeft(8, ' ') + calcL.Replace(',', ' ') + calcM.PadLeft(3, ' ') +
                        calcN.Replace(',', ' ') + calcO.Replace(',', ' ') + calcP.Replace(',', ' ') + calcQ + calcR.Replace(',', ' ') + calcS.Replace(',', ' ') +
                        calcT.Replace(',', ' ') + calcU + calcV + calcW.Replace(',', ' ') + calcX.Replace(',', ' ') + calcY.Replace(',', ' ') + calcZ.Replace(',', ' ') +
                        calcAA.Replace(',', ' ') + calcAB + calcAC.PadLeft(2, ' ');

                    x.WriteLine("|" + calcA + "|" + calcB + "|" + calcC + "|" + calcD + "|" + calcE + "|" + calcF + "|" + calcG + "|" + calcH + "|" + calcI + "|" + calcJ + "|" + calcK + "|" + calcL + "|" + calcM + "|" + calcN + "|" + calcO + "|" + calcP + "|" + calcQ + "|" + calcR + "|" + calcS + "|" + calcT + "|" + calcU + "|" + calcV + "|" + calcW + "|" + calcX + "|" + calcY + "|" + calcZ + "|" + calcAA + "|" + calcAB + "|" + calcAC + "|" + calcAD + "|");
                }

                i++;
            }
        }

        private void GerarBloco433(ExcelWorksheet sheet, out int i, out string caminho, out string path)
        {
            //Inicio do bloco 4.3.3
            StreamWriter x;


            caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            path = caminho + @"\Bloco4.3.3.txt";
            x = File.CreateText(path);

            // Títulos
            x.WriteLine("|1 - Modelo|2 - Série|3 - Nr docto|4 - DT Emissão|5 - Participante|" +
            "6 - DT Entrada|7 - VL Mercadorias|8 - Desc|9 - Vlr Frete|10 - Vlr Seguro|11 - Vlr Out Despesas|12 - Vlr IPI|13 - Vlr ICMS ST|" +
            "14 -Vlr T NF|15 - IE Sub|16 - Tipo Fat|17 - Observ|18 - Ato Declaratorio|19 - Mod Doc Ref|20 - Ser/Sub Doc Ref|21 - Num Doc Ref|" +
            "22 - Data Em Doc Ref|23 -Part Doc Ref|Linha preenchida IN86 - 4.3.3|");


            i = 2;
            string calcA, calcB, calcC, calcD, calcE, calcF, calcG, calcH, calcI, calcJ, calcK, calcL, calcM, calcN, calcO, calcP, calcQ, calcR, calcS, calcT, calcU, calcV, calcW, calcX, calcY;

            foreach (string y in C100)
            {

                calcA = "";
                calcB = "";
                calcC = "";
                calcD = "";
                calcE = "";
                calcF = "";
                calcG = "";
                calcH = "";
                calcI = "";
                calcJ = "";
                calcL = "";
                calcK = "";
                calcM = "";
                calcN = "";
                calcO = "";
                calcP = "";
                calcQ = "";
                calcR = "";
                calcS = "";
                calcT = "";
                calcU = "";
                calcV = "";
                calcW = "";
                calcX = "";
                calcY = "";

                string[] value = y.Split('|');//.Where(x => x != "");

                CarregaBlocoCSheet40(sheet, i, value);

                if (value.Length >= 30)
                {
                    if (sheet.Cells[i, 1].Value.ToString() == "C100" && sheet.Cells[i, 3].Value.ToString() == "1")
                    {
                        calcA = sheet.Cells[i, 5].Value.ToString();
                        calcB = sheet.Cells[i, 7].Value.ToString().Replace("*", "").PadLeft(3, '0');
                        calcC = sheet.Cells[i, 8].Value.ToString().PadLeft(9, '0');

                        if (sheet.Cells[i, 9].Value.ToString() == "")
                        {
                            calcD = "";
                        }
                        else
                        {
                            calcD = sheet.Cells[i, 9].Value.ToString();
                        }

                        if (sheet.Cells[i, 4].Value.ToString() == "")
                        {
                            calcE = "";
                        }
                        else
                        {
                            calcE = sheet.Cells[i, 4].Value.ToString();
                        }

                        if (sheet.Cells[i, 10].Value.ToString() == "")
                        {
                            calcF = "";
                        }
                        else
                        {
                            calcF = sheet.Cells[i, 11].Value.ToString();
                        }

                        calcG = string.Format(@"{0:f}", sheet.Cells[i, 16].Value.ToString());
                        calcH = string.Format(@"{0:f}", sheet.Cells[i, 14].Value.ToString());
                        calcI = string.Format(@"{0:f}", sheet.Cells[i, 18].Value.ToString());
                        calcJ = string.Format(@"{0:f}", sheet.Cells[i, 19].Value.ToString());
                        calcK = string.Format(@"{0:f}", sheet.Cells[i, 20].Value.ToString());
                        calcL = string.Format(@"{0:f}", sheet.Cells[i, 25].Value.ToString());
                        calcM = string.Format(@"{0:f}", sheet.Cells[i, 24].Value.ToString());
                        calcN = string.Format(@"{0:f}", sheet.Cells[i, 12].Value.ToString());
                        calcO = "";

                        if (sheet.Cells[i, 13].Value.ToString() == "0")
                        {
                            calcP = "1";
                        }
                        else
                        {
                            if (sheet.Cells[i, 13].Value.ToString() == "1")
                            {
                                calcP = "2";
                            }
                            else
                            {
                                calcP = "";
                            }
                        }

                        calcQ = "";
                        calcR = "";

                        if (sheet.Cells[i, 38].Value.ToString() == "")
                        {
                            calcS = "";
                        }
                        else
                        {
                            calcS = sheet.Cells[i, 38].Value.ToString();
                        }

                        if (sheet.Cells[i, 39].Value.ToString() == "")
                        {
                            calcT = "";
                        }
                        else
                        {
                            calcT = sheet.Cells[i, 39].Value.ToString();
                        }

                        if (sheet.Cells[i, 40].Value.ToString() == "")
                        {
                            calcU = "";
                        }
                        else
                        {
                            calcU = sheet.Cells[i, 40].Value.ToString();
                        }

                        if (sheet.Cells[i, 41].Value.ToString() == "")
                        {
                            calcV = "";
                        }
                        else
                        {
                            calcV = sheet.Cells[i, 41].Value.ToString();
                        }

                        if (sheet.Cells[i, 42].Value.ToString() == "")
                        {
                            calcW = "";
                        }
                        else
                        {
                            calcW = sheet.Cells[i, 42].Value.ToString();
                        }

                        calcX = "";

                    }
                    else
                    {
                        calcA = "FALSO";
                        calcB = "FALSO";
                        calcC = "FALSO";
                        calcD = "FALSO";
                        calcE = "FALSO";
                        calcF = "FALSO";
                        calcG = "FALSO";
                        calcH = "FALSO";
                        calcI = "FALSO";
                        calcJ = "FALSO";
                        calcK = "FALSO";
                        calcL = "FALSO";
                        calcM = "FALSO";
                        calcN = "FALSO";
                        calcO = "";
                        calcP = "FALSO";
                        calcQ = "";
                        calcR = "";
                        calcS = "FALSO";
                        calcT = "FALSO";
                        calcU = "FALSO";
                        calcV = "FALSO";
                        calcW = "FALSO";
                        calcX = "";
                    }

                    calcY = calcA + calcB.PadLeft(5, ' ').Substring(5) + calcC + calcD.PadLeft(8, ' ').Substring(8) + calcE.PadLeft(14, ' ').Substring(14) +
                        calcF.PadLeft(8, ' ').Substring(8) + calcG.Replace(',', ' ') + calcH.Replace(',', ' ') + calcI.Replace(',', ' ') + calcJ.Replace(',', ' ') +
                        calcK.Replace(',', ' ') + calcL.Replace(',', ' ') + calcM.Replace(',', ' ') + calcN.Replace(',', ' ') + calcO.PadLeft(14, ' ').Substring(14) +
                        calcP.PadLeft(1, ' ').Substring(1) + calcQ.PadLeft(45, ' ').Substring(45) + calcR.PadLeft(50, ' ').Substring(50) + calcS.PadLeft(2, ' ').Substring(2) +
                        calcT.PadLeft(5, ' ').Substring(5) + calcU.PadLeft(9, ' ').Substring(9) + calcV.PadLeft(8, ' ').Substring(8) + calcW.PadLeft(14, ' ').Substring(14);

                    x.WriteLine("|" + calcA + "|" + calcB + "|" + calcC + "|" + calcD + "|" + calcE + "|" + calcF + "|" + calcG + "|" + calcH + "|" + calcI + "|" + calcJ + "|" + calcK + "|" + calcL + "|" + calcM + "|" + calcN + "|" + calcO + "|" + calcP + "|" + calcQ + "|" + calcR + "|" + calcS + "|" + calcT + "|" + calcU + "|" + calcV + "|" + calcW + "|" + calcX + "|" + calcY + "|");
                }

                i++;
            }
        }

        private void GerarBloco432(ExcelWorksheet sheet, out int i, out string caminho, out string path)
        {
            //Inicio do bloco 4.3.2
            StreamWriter x;


            caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            path = caminho + @"\Bloco4.3.2.txt";
            x = File.CreateText(path);

            // Títulos
            x.WriteLine("|1 - Ind Movto|2 - Modelo|3 - Série|4 - Nr docto|5 - DT Emissão" +
            "|6 - Nr item|7 - Cód Merc/Serv|8 - Descrição compl|9 - CFOP|10 - Cod Nat|11 - Clas Fisc Merc|12 - Qtdade|13 - unid" +
            "|14 - Vlr Unit|15 - Vlr Tot Item|16 - Desconto|17 - Ind Trib IPI|18 - Aliq IPI|19 - BC IPI|20 - Vlr IPI|21 - CST ICMS" +
            "|22 - Ind ICMS|23 - Aliq ICMS|24 - BC ICMS|25 - Vlr ICMS Pr|26 - BC ICMS ST|27 - Vlr ICMS ST|28 - Ind Mov|29 - CST IPI" +
            "|Linha Preenchida IN86 - 4.3.2|");

            i = 2;
            string calcA, calcB, calcC, calcD, calcE, calcF, calcG, calcH, calcI, calcJ, calcK, calcL, calcM, calcN, calcO, calcP, calcQ, calcR, calcS, calcT, calcU, calcV, calcW, calcX, calcY, calcZ, calcAA, calcAB, calcAC, calcAD;
            float nCalc;
            int count = 1, countTemp;

            countTemp = 0;

            foreach (string y in C100)
            {
                calcA = "";
                calcB = "";
                calcC = "";
                calcD = "";
                calcE = "";
                calcF = "";
                calcG = "";
                calcH = "";
                calcI = "";
                calcJ = "";
                calcL = "";
                calcK = "";
                calcM = "";
                calcN = "";
                calcO = "";
                calcP = "";
                calcQ = "";
                calcR = "";
                calcS = "";
                calcT = "";
                calcU = "";
                calcV = "";
                calcW = "";
                calcX = "";
                calcY = "";
                calcZ = "";
                calcAA = "";
                calcAB = "";
                calcAC = "";
                calcAD = "";

                string[] value = y.Split('|');//.Where(x => x != "");

                CarregarSheetBlocoC(sheet, i, value);

                if (value.Length >= 39)
                {
                    if (calcTempC[countTemp].ToString() != "" && sheet.Cells[i, 1].Value != null)
                    {
                        if (calcTempC[countTemp].Substring(0, 1).ToString() == "0" && sheet.Cells[i, 1].Value.ToString() == "C170")
                        {
                            calcA = "E";

                            calcB = calcTempC[countTemp].Substring(2, 2);
                            calcC = calcTempC[countTemp].Substring(14, 3);
                            calcD = calcTempC[countTemp].Substring(5, 9);
                            calcE = calcTempC[countTemp].Substring(17, 8);

                            calcC = calcC.Replace('*', ' ');
                        }
                        else
                        {//errado a verificação e os valores do calculo
                            calcA = "S";

                            calcB = calcTempC[countTemp].Substring(2, 2);
                            calcC = calcTempC[countTemp].Substring(14, 3);
                            calcD = calcTempC[countTemp].Substring(5, 9);
                            calcE = calcTempC[countTemp].Substring(17, 8);
                            calcC = calcC.Replace('*', ' ');
                        }
                        if (!string.IsNullOrEmpty(sheet.Cells[i, 2].Value.ToString()))
                        {
                            calcF = sheet.Cells[i, 2].Value.ToString().PadLeft(3, '0');
                        }
                        else
                        {
                            calcF = "FALSO";
                        }
                        if (!string.IsNullOrEmpty(sheet.Cells[i, 3].Value.ToString()))
                        {
                            calcG = sheet.Cells[i, 3].Value.ToString();
                        }
                        else
                        {
                            calcG = "FALSO";
                        }
                        if (!string.IsNullOrEmpty(sheet.Cells[i, 4].Value.ToString()))
                        {
                            calcH = sheet.Cells[i, 4].Value.ToString();
                        }
                        else
                        {
                            calcH = "FALSO";
                        }
                        if (!string.IsNullOrEmpty(sheet.Cells[i, 11].Value.ToString()))
                        {
                            calcI = sheet.Cells[i, 11].Value.ToString();
                        }
                        else
                        {
                            calcI = "FALSO";
                        }
                        if (!string.IsNullOrEmpty(sheet.Cells[i, 12].Value.ToString()))
                        {
                            calcJ = sheet.Cells[i, 12].Value.ToString();
                        }
                        else
                        {
                            calcJ = "FALSO";
                        }
                        if (calcG.ToString() != null && calcG.ToString() != "FALSO")
                        {
                            foreach (string g in calcR200)
                            {
                                if (g.Equals(calcG.ToString()))
                                {
                                    calcK = g;
                                }

                                if (calcK != "")
                                    break;

                                count++;
                            }
                        }
                        else
                        {
                            calcK = "FALSO";
                        }
                        if (!string.IsNullOrEmpty(sheet.Cells[i, 5].Value.ToString()))
                        {
                            calcL = string.Format(@"{0:f}", sheet.Cells[i, 5].Value.ToString());
                        }
                        else
                        {
                            calcL = "FALSO";
                        }
                        if (!string.IsNullOrEmpty(sheet.Cells[i, 6].Value.ToString()))
                        {
                            calcM = sheet.Cells[i, 6].Value.ToString();
                        }
                        else
                        {
                            calcM = "FALSO";
                        }
                        if (!string.IsNullOrEmpty(sheet.Cells[i, 7].Value.ToString()) && calcL != null && calcL != "FALSO")
                        {
                            nCalc = float.Parse(sheet.Cells[i, 7].Value.ToString());
                            calcN = string.Format(@"{0:f}", nCalc / float.Parse(calcL));
                        }
                        else
                        {
                            calcN = "FALSO";
                        }
                        if (!string.IsNullOrEmpty(sheet.Cells[i, 7].Value.ToString()))
                        {
                            calcO = string.Format(@"{0:f}", sheet.Cells[i, 7].Value.ToString());
                        }
                        else
                        {
                            calcO = "FALSO";
                        }
                        if (!string.IsNullOrEmpty(sheet.Cells[i, 8].Value.ToString()))
                        {
                            calcP = string.Format(@"{0:f}", sheet.Cells[i, 8].Value.ToString());
                        }
                        else
                        {
                            calcP = "FALSO";
                        }
                        if (!string.IsNullOrEmpty(sheet.Cells[i, 24].Value.ToString()))
                        {
                            calcT = sheet.Cells[i, 24].Value.ToString();
                        }
                        else
                        {
                            calcT = "FALSO";
                        }
                        if (!string.IsNullOrEmpty(sheet.Cells[i, 23].Value.ToString()))
                        {
                            calcR = string.Format(@"{0:f}", sheet.Cells[i, 23].Value.ToString());
                        }
                        else
                        {
                            calcR = "FALSO";
                        }
                        if (!string.IsNullOrEmpty(sheet.Cells[i, 22].Value.ToString()))
                        {
                            calcS = string.Format(@"{0:f}", sheet.Cells[i, 22].Value.ToString());
                        }
                        else
                        {
                            calcS = "FALSO";
                        }
                        if (!string.IsNullOrEmpty(sheet.Cells[i, 10].Value.ToString()))
                        {
                            calcU = string.Format(@"{0:f}", sheet.Cells[i, 10].Value.ToString());
                        }
                        else
                        {
                            calcU = "FALSO";
                        }
                        if (string.IsNullOrEmpty(sheet.Cells[i, 14].Value.ToString()))
                        {
                            calcW = "0,00";
                        }
                        else
                        {
                            calcW = string.Format(@"{0:f}", sheet.Cells[i, 14].Value.ToString());
                        }
                        if (!string.IsNullOrEmpty(sheet.Cells[i, 13].Value.ToString()))
                        {
                            calcX = sheet.Cells[i, 13].Value.ToString();
                        }
                        else
                        {
                            calcX = "FALSO";
                        }
                        if (!string.IsNullOrEmpty(sheet.Cells[i, 15].Value.ToString()))
                        {
                            calcY = string.Format(@"{0:f}", sheet.Cells[i, 15].Value.ToString());
                        }
                        else
                        {
                            calcY = "FALSO";
                        }
                        if (!string.IsNullOrEmpty(sheet.Cells[i, 16].Value.ToString()))
                        {
                            calcZ = string.Format(@"{0:f}", sheet.Cells[i, 16].Value.ToString());
                        }
                        else
                        {
                            calcZ = "FALSO";
                        }
                        if (!string.IsNullOrEmpty(sheet.Cells[i, 18].Value.ToString()))
                        {
                            calcAA = string.Format(@"{0:f}", sheet.Cells[i, 18].Value.ToString());
                        }
                        else
                        {
                            calcAA = "FALSO";
                        }
                        if (sheet.Cells[i, 9].Value.ToString() == "0")
                        {
                            calcAB = "S";
                        }
                        else
                        {
                            calcAB = "N";
                        }
                    }
                    else
                    {
                        calcA = "FALSO";
                        calcB = "FALSO";
                        calcC = "FALSO";
                        calcD = "FALSO";
                        calcE = "FALSO";
                        calcF = "FALSO";
                        calcG = "FALSO";
                        calcH = "FALSO";
                        calcI = "FALSO";
                        calcJ = "FALSO";
                        calcK = "FALSO";
                        calcL = "FALSO";
                        calcM = "FALSO";
                        calcN = "FALSO";
                        calcO = "FALSO";
                        calcP = "FALSO";
                        calcT = "FALSO";
                        calcR = "FALSO";
                        calcS = "FALSO";
                        calcU = "FALSO";
                        calcW = "0,00";
                        calcX = "FALSO";
                        calcY = "FALSO";
                        calcZ = "FALSO";
                        calcAA = "FALSO";
                        calcAB = "FALSO";
                    }
                    countTemp++;
                    if (calcT.Equals("0,00"))
                    {
                        calcQ = "2";
                    }
                    else
                    {
                        calcQ = "1";
                    }

                    if (!string.IsNullOrEmpty(calcU) && calcU != "FALSO")
                    {
                        if (int.Parse(calcU.Substring(2, 1)) < 3)
                        {
                            calcV = "1";
                        }
                        else if (int.Parse(calcU.Substring(2, 1)) == 9)
                        {
                            calcV = "3";
                        }
                        else if (int.Parse(calcU.Substring(2, 1)) == 7)
                        {
                            calcV = "1";
                        }
                        else
                        {
                            calcV = "2";
                        }
                    }
                    else
                    {
                        calcV = "2";
                    }

                    if (sheet.Cells[i, 15].Value.ToString() == "1" && calcA == "S")
                    {
                        calcAC = "50";
                    }
                    else if (sheet.Cells[i, 15].Value.ToString() == "2" && calcA == "S")
                    {
                        calcAC = "52";
                    }
                    else if (sheet.Cells[i, 15].Value.ToString() == "1" && calcA == "E")
                    {
                        calcAC = "00";
                    }
                    else
                    {
                        calcAC = "02";
                    }

                    calcAD = calcA + calcB + calcC.PadRight(5, ' ') + calcD + calcE + calcF + calcG.PadLeft(20, ' ') + calcH.PadRight(45, ' ') + calcI.PadRight(6, ' ') +
                        calcJ.PadRight(6, ' ') + calcK.Replace(',', ' ').PadLeft(8, '0') + calcL.Replace(',', ' ').PadLeft(17, '0') + calcM.PadLeft(3, ' ').PadLeft(3, '0') +
                        calcN.Replace(',', ' ').PadLeft(17, '0') + calcO.Replace(',', ' ').PadLeft(17, '0') + calcP.Replace(',', ' ').PadLeft(17, '0') + calcQ +
                        calcR.Replace(',', ' ').PadLeft(5, '0') + calcS.Replace(',', ' ').PadLeft(17, '0') + calcT.Replace(',', ' ').PadLeft(17, '0') + calcU + calcV +
                        calcW.Replace(',', ' ').PadLeft(5, '0') + calcX.Replace(',', ' ').PadLeft(17, '0') + calcY.Replace(',', ' ').PadLeft(17, '0') + calcZ.Replace(',', ' ').PadLeft(17, '0') +
                        calcAA.Replace(',', ' ').PadLeft(17, '0') + calcAB + calcAC.PadRight(2, ' ');

                    x.WriteLine("|" + calcA + "|" + calcB + "|" + calcC + "|" + calcD + "|" + calcE + "|" + calcF + "|" + calcG + "|" + calcH + "|" + calcI + "|" + calcJ + "|" + calcK + "|" + calcL + "|" + calcM + "|" + calcN + "|" + calcO + "|" + calcP + "|" + calcQ + "|" + calcR + "|" + calcS + "|" + calcT + "|" + calcU + "|" + calcV + "|" + calcW + "|" + calcX + "|" + calcY + "|" + calcZ + "|" + calcAA + "|" + calcAB + "|" + calcAC + "|" + calcAD + "|");

                }
                i++;

            }
        }

        private static void CarregarSheetBlocoC(ExcelWorksheet sheet, int i, string[] value)
        {
            for (int j = 1; j < value.Count(); j++)
            {
                if (value.Length >= 30)
                {
                    if (j == 1)
                    {
                        sheet.Cells[i, 1].Value = value[j];
                    }
                    if (j == 2)
                    {
                        sheet.Cells[i, 2].Value = value[j];
                    }
                    if (j == 3)
                    {
                        sheet.Cells[i, 3].Value = value[j];
                    }
                    if (j == 4)
                    {
                        sheet.Cells[i, 4].Value = value[j];
                    }
                    if (j == 5)
                    {
                        sheet.Cells[i, 5].Value = value[j];
                    }
                    if (j == 6)
                    {
                        sheet.Cells[i, 6].Value = value[j];
                    }
                    if (j == 7)
                    {
                        sheet.Cells[i, 7].Value = value[j];
                    }
                    if (j == 8)
                    {
                        sheet.Cells[i, 8].Value = value[j];
                    }
                    if (j == 9)
                    {
                        sheet.Cells[i, 9].Value = value[j];
                    }
                    if (j == 10)
                    {
                        sheet.Cells[i, 10].Value = value[j];
                    }
                    if (j == 11)
                    {
                        sheet.Cells[i, 11].Value = value[j];
                    }
                    if (j == 12)
                    {
                        sheet.Cells[i, 12].Value = value[j];
                    }
                    if (j == 13)
                    {
                        sheet.Cells[i, 13].Value = value[j];
                    }
                    if (j == 14)
                    {
                        sheet.Cells[i, 14].Value = value[j];
                    }
                    if (j == 15)
                    {
                        sheet.Cells[i, 15].Value = value[j];
                    }
                    if (j == 16)
                    {
                        sheet.Cells[i, 16].Value = value[j];
                    }
                    if (j == 17)
                    {
                        sheet.Cells[i, 17].Value = value[j];
                    }
                    if (j == 18)
                    {
                        sheet.Cells[i, 18].Value = value[j];
                    }
                    if (j == 19)
                    {
                        sheet.Cells[i, 19].Value = value[j];
                    }
                    if (j == 20)
                    {
                        sheet.Cells[i, 20].Value = value[j];
                    }
                    if (j == 21)
                    {
                        sheet.Cells[i, 21].Value = value[j];
                    }
                    if (j == 22)
                    {
                        sheet.Cells[i, 22].Value = value[j];
                    }
                    if (j == 23)
                    {
                        sheet.Cells[i, 23].Value = value[j];
                    }
                    if (j == 24)
                    {
                        sheet.Cells[i, 24].Value = value[j];

                    }
                    if (j == 25)
                    {
                        sheet.Cells[i, 25].Value = value[j];
                    }
                    if (j == 26)
                    {
                        sheet.Cells[i, 26].Value = value[j];
                    }
                    if (j == 27)
                    {
                        sheet.Cells[i, 27].Value = value[j];
                    }
                    if (j == 28)
                    {
                        sheet.Cells[i, 28].Value = value[j];
                    }
                    if (j == 29)
                    {
                        sheet.Cells[i, 29].Value = value[j];
                    }
                }
            }
        }

        private void GerarBloco431(ExcelWorksheet sheet, out int i, int num, out string caminho, out string path)
        {
            // Inicio do bloco 4.3.1 
            StreamWriter x;


            caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            path = caminho + @"\Bloco4.3.1.txt";
            x = File.CreateText(path);

            // Títulos
            x.WriteLine("|1 - Ind Movto|2 - Modelo|3 - Série|4 - Nr docto|5 - DT Emissão|" +
                "6 - Participante|7 - DT Entrada|8 - Vl Mercadorias|9 - Desc|10 - Vlr Frete|11 - Vlr Seguro|12 - Vlr Out Despesas|13 - Vlr IPI|" +
                "14 - Vlr ICMS ST|15 - VlrNF|16 - IE Sub|17 - Via Transp|18 - Código Transp|19 - Qt Vol|20 - Esp Volume|21 - Peso Bruto|" +
                "22 - Peso Liq|23 - Mod Frete|24 - Ident Veic|25 - Ind Canc|26 - Tipo Fat|27 - Observ|28 - ADE|29 - Mod doc Ref|" +
                "|30 - Ser Sub|31 - Nr doc ref|32 - DT Emis Ref|33 - Cod Part Ref|Linha preenchida IN86 - 4.3.1|");

            i = 2;
            string calcA, calcB, calcC, calcD, calcE, calcF, calcG, calcH, calcI, calcJ, calcK, calcL, calcM, calcN, calcO, calcW, calcY, calcZ, calcAC, calcAD, calcAE, calcAF, calcAG, calcAH;

            foreach (string y in C100)
            {
                calcA = "";
                calcB = "";
                calcC = "";
                calcD = "";
                calcE = "";
                calcF = "";
                calcG = "";
                calcH = "";
                calcI = "";
                calcJ = "";
                calcK = "";
                calcL = "";
                calcM = "";
                calcN = "";
                calcO = "";
                calcW = "";
                calcY = "";
                calcZ = "";
                calcAC = "";
                calcAD = "";
                calcAE = "";
                calcAF = "";
                calcAG = "";
                calcAH = "";
                string[] value = y.Split('|');//.Where(x => x != "");
                CarregaBlocoCSheet40(sheet, i, value);
                if (sheet.Cells[i, 1].Value != null && sheet.Cells[i, 3].Value != null && sheet.Cells[i, 4].Value != null)
                {
                    if (sheet.Cells[i, 1].Value.ToString() == "C100" && sheet.Cells[i, 3].Value.ToString() == "0")
                    {
                        if (sheet.Cells[i, 4].Value.ToString() == "0")
                        {
                            calcA = "E";
                        }
                        else
                        {
                            calcA = "S";
                        }

                        if (sheet.Cells[i, 5].Value != null)
                        {
                            calcB = sheet.Cells[i, 5].Value.ToString();
                        }
                        else
                        {
                            calcB = "FALSO";
                        }

                        if (sheet.Cells[i, 7].Value != null)
                        {
                            calcC = sheet.Cells[i, 7].Value.ToString().PadLeft(3, '0');
                        }
                        else
                        {
                            calcC = "FALSO";
                        }

                        if (sheet.Cells[i, 8].Value != null)
                        {
                            calcD = sheet.Cells[i, 8].Value.ToString().PadLeft(9, '0');
                        }
                        else
                        {
                            calcD = "FALSO";
                        }

                        if (sheet.Cells[i, 10].Value.ToString() == "")
                        {
                            calcE = "";
                        }
                        else
                        {
                            calcE = sheet.Cells[i, 10].Value.ToString();
                        }

                        if (sheet.Cells[i, 4].Value.ToString() == "")
                        {
                            calcF = "";
                        }
                        else
                        {
                            calcF = sheet.Cells[i, 4].Value.ToString().PadRight(14, ' ');
                        }

                        if (sheet.Cells[i, 11].Value.ToString() == "")
                        {
                            calcG = "";
                        }
                        else
                        {
                            calcG = sheet.Cells[i, 11].Value.ToString();
                        }

                        if (sheet.Cells[i, 16].Value.ToString() != null)
                        {
                            calcH = string.Format(@"{0:f}", sheet.Cells[i, 16].Value.ToString());
                        }
                        else
                        {
                            calcH = "FALSO";
                        }

                        if (sheet.Cells[i, 14].Value.ToString() != null)
                        {
                            calcI = string.Format(@"{0:f}", sheet.Cells[i, 14].Value.ToString());
                        }
                        else
                        {
                            calcI = "FALSO";
                        }
                        if (sheet.Cells[i, 18].Value.ToString() != "")
                        {
                            calcJ = string.Format(@"{0:f}", sheet.Cells[i, 18].Value.ToString());
                        }
                        else
                        {
                            calcJ = "FALSO";
                        }
                        if (sheet.Cells[i, 19].Value.ToString() != "")
                        {
                            calcK = string.Format(@"{0:f}", sheet.Cells[i, 19].Value.ToString());
                        }
                        else
                        {
                            calcK = "FALSO";
                        }
                        if (sheet.Cells[i, 20].Value.ToString() != "")
                        {
                            calcL = string.Format(@"{0:f}", sheet.Cells[i, 20].Value.ToString());
                        }
                        else
                        {
                            calcL = "FALSO";
                        }
                        if (sheet.Cells[i, 25].Value.ToString() != "")
                        {
                            calcM = string.Format(@"{0:f}", sheet.Cells[i, 25].Value.ToString());
                        }
                        else
                        {
                            calcM = "FALSO";
                        }
                        if (sheet.Cells[i, 24].Value.ToString() != "")
                        {
                            calcN = string.Format(@"{0:f}", sheet.Cells[i, 24].Value.ToString());
                        }
                        else
                        {
                            calcN = "FALSO";
                        }
                        if (sheet.Cells[i, 12].Value.ToString() != "")
                        {
                            calcO = string.Format(@"{0:f}", sheet.Cells[i, 12].Value.ToString());
                        }
                        else
                        {
                            calcO = "FALSO";
                        }
                        if (sheet.Cells[i, 17].Value != null || sheet.Cells[i, 17].Value.ToString() != "")
                        {
                            if (sheet.Cells[i, 17].Value.ToString() == "1" || sheet.Cells[i, 17].Value.ToString() == "2" || sheet.Cells[i, 17].Value.ToString() == "9")
                            {
                                calcW = "FOB";
                            }
                            else
                            {
                                calcW = "CIF";
                            }
                        }
                        else
                        {
                            calcW = "FALSO";
                        }
                        if (sheet.Cells[i, 6].Value.ToString() != "" || sheet.Cells[i, 6].Value != null)
                        {
                            if (sheet.Cells[i, 6].Value.ToString() == "02" || sheet.Cells[i, 6].Value.ToString() == "03"
                                || sheet.Cells[i, 6].Value.ToString() == "04" || sheet.Cells[i, 6].Value.ToString() == "05")
                            {
                                calcY = "S";
                            }
                            else
                            {
                                calcY = "N";
                            }
                        }
                        else
                        {
                            calcY = "FALSO";
                        }
                        if (sheet.Cells[i, 13].Value.ToString() != "" || sheet.Cells[i, 13].Value != null)
                        {
                            if (sheet.Cells[i, 13].Value.ToString() == "0")
                            {
                                calcZ = "1";
                            }
                            else
                            {
                                if (sheet.Cells[i, 13].Value.ToString() == "1")
                                {
                                    calcZ = "2";
                                }
                                else
                                {
                                    calcZ = "";
                                }
                            }
                        }
                        else
                        {
                            calcZ = "FALSO";
                        }
                        if (sheet.Cells[i, 38].Value == null)
                        {
                            calcAC = "";
                        }
                        else
                        {
                            calcAC = sheet.Cells[i, 38].Value.ToString();
                        }
                        if (sheet.Cells[i, 39].Value == null)
                        {
                            calcAD = "";
                        }
                        else
                        {
                            calcAD = sheet.Cells[i, 39].Value.ToString();
                        }
                        if (sheet.Cells[i, 40].Value == null)
                        {
                            calcAE = "";
                        }
                        else
                        {
                            calcAE = sheet.Cells[i, 40].Value.ToString();
                        }
                        if (sheet.Cells[i, 41].Value == null)
                        {
                            calcAF = "";
                        }
                        else
                        {
                            calcAF = sheet.Cells[i, 41].Value.ToString();
                        }
                        if (sheet.Cells[i, 42].Value == null)
                        {
                            calcAG = "";
                        }
                        else
                        {
                            calcAG = sheet.Cells[i, 42].Value.ToString();
                        }
                    }
                    else
                    {
                        calcA = "FALSO";
                        calcB = "FALSO";
                        calcC = "FALSO";
                        calcD = "FALSO";
                        calcE = "FALSO";
                        calcE = "FALSO";
                        calcF = "FALSO";
                        calcG = "FALSO";
                        calcH = "FALSO";
                        calcI = "FALSO";
                        calcJ = "FALSO";
                        calcK = "FALSO";
                        calcL = "FALSO";
                        calcM = "FALSO";
                        calcN = "FALSO";
                        calcO = "FALSO";
                        calcW = "FALSO";
                        calcY = "FALSO";
                        calcAC = "FALSO";
                        calcAD = "FALSO";
                        calcAE = "FALSO";
                        calcAF = "FALSO";
                        calcAG = "FALSO";
                        calcAH = "FALSO";
                    }
                }
                else
                {
                    calcA = "FALSO";
                    calcB = "FALSO";
                    calcC = "FALSO";
                    calcD = "FALSO";
                    calcE = "FALSO";
                    calcE = "FALSO";
                    calcF = "FALSO";
                    calcG = "FALSO";
                    calcH = "FALSO";
                    calcI = "FALSO";
                    calcJ = "FALSO";
                    calcK = "FALSO";
                    calcL = "FALSO";
                    calcM = "FALSO";
                    calcN = "FALSO";
                    calcO = "FALSO";
                    calcW = "FALSO";
                    calcY = "FALSO";
                    calcAC = "FALSO";
                    calcAD = "FALSO";
                    calcAE = "FALSO";
                    calcAF = "FALSO";
                    calcAG = "FALSO";
                    calcAH = "FALSO";
                }

                if (calcAH != "FALSO")
                {
                    calcAH = calcA + calcB + calcC.PadLeft(5, '0') + calcD + calcE.PadLeft(8, '0') + calcF.PadLeft(14, '0')
                        + calcG.PadLeft(8, '0') + calcH.Length + calcH.Replace(',', ' ') + calcI.Length + calcI.Replace(',', ' ') + calcJ.Length + calcJ.Replace(',', ' ')
                        + calcK.Length + calcK.Replace(',', ' ') + calcL.Length + calcL.Replace(',', ' ') + calcM.Length + calcM.Replace(',', ' ') + calcN.Length + calcN.Replace(',', ' ')
                        + calcO.Length + calcO.Replace(',', ' ') + "" + "" + "" + "" + "" + "" + "" + calcW.PadLeft(3, '0') + "" + calcY + calcZ.PadLeft(1, '0')
                        + "" + "" + calcAC.PadLeft(2, '0') + calcAD.PadLeft(5, '0') + calcAE.Length + calcAE.Replace(',', ' ') + calcAF.Length + calcAF.Replace(',', ' ')
                        + calcAG.PadLeft(14, '0');
                }

                x.WriteLine("|" + calcA + "|" + calcB + "|" + calcC + "|" + calcD + "|" + calcE + "|" + calcF + "|" + calcG + "|" + calcH + "|" + calcI + "|" + calcJ + "|" + calcK + "|" + calcL + "|" + calcM + "|" + calcN + "|" + calcO + "|" + calcW + "|" + calcY + "|" + calcZ + "|" + calcAC + "|" + calcAD + "|" + calcAD + "|" + calcAE + "|" + calcAF + "|" + calcAG + "|" + calcAH + "|");
                i++;
            }

        }

        private static void CarregaBlocoCSheet40(ExcelWorksheet sheet, int i, string[] value)
        {
            for (int j = 1; j < value.Count(); j++)
            {
                if (value.Length > 31)
                {
                    if (j == 1)
                    {
                        sheet.Cells[i, 1].Value = value[j];
                    }
                    if (j == 3)
                    {
                        sheet.Cells[i, 3].Value = value[j];
                    }
                    if (j == 4)//substituindo a mesma linha, corrigir o valor de i
                    {
                        sheet.Cells[i, 4].Value = value[j];
                    }
                    if (j == 5)
                    {
                        sheet.Cells[i, 5].Value = value[j];
                    }
                    if (j == 7)
                    {
                        sheet.Cells[i, 7].Value = value[j];
                    }
                    if (j == 8)
                    {
                        sheet.Cells[i, 8].Value = value[j];
                    }
                    if (j == 10)
                    {
                        sheet.Cells[i, 10].Value = value[j];
                    }
                    if (j == 11)
                    {
                        sheet.Cells[i, 11].Value = value[j];
                    }
                    if (j == 16)
                    {
                        sheet.Cells[i, 16].Value = value[j];
                    }
                    if (j == 14)
                    {
                        sheet.Cells[1, 14].Value = value[j];
                    }
                    if (j == 18)
                    {
                        sheet.Cells[i, 18].Value = value[j];
                    }
                    if (j == 19)
                    {
                        sheet.Cells[i, 19].Value = value[j];
                    }
                    if (j == 20)
                    {
                        sheet.Cells[i, 20].Value = value[j];
                    }
                    if (j == 25)
                    {
                        sheet.Cells[i, 25].Value = value[j];
                    }
                    if (j == 24)
                    {
                        sheet.Cells[i, 24].Value = value[j];
                    }
                    if (j == 12)
                    {
                        sheet.Cells[i, 12].Value = value[j];
                    }
                    if (j == 17)
                    {
                        sheet.Cells[i, 17].Value = value[j];
                    }
                    if (j == 6)
                    {
                        sheet.Cells[i, 6].Value = value[j];
                    }
                    if (j == 13)
                    {
                        sheet.Cells[i, 13].Value = value[j];
                    }
                    if (j == 38)
                    {
                        sheet.Cells[i, 38].Value = value[j];
                    }
                    if (j == 39)
                    {
                        sheet.Cells[i, 39].Value = value[j];
                    }
                    if (j == 40)
                    {
                        sheet.Cells[i, 40].Value = value[j];
                    }
                    if (j == 41)
                    {
                        sheet.Cells[i, 41].Value = value[j];
                    }
                    if (j == 42)
                    {
                        sheet.Cells[i, 42].Value = value[j];
                    }
                }
                else
                {
                    if (j == 1)
                    {
                        sheet.Cells[i, 1].Value = value[j];
                    }
                    if (j == 2)
                    {
                        sheet.Cells[i, 2].Value = value[j];
                    }
                    if (j == 3)
                    {
                        sheet.Cells[i, 3].Value = value[j];
                    }
                    if (j == 4)//substituindo a mesma linha, corrigir o valor de i
                    {
                        sheet.Cells[i, 4].Value = value[j];
                    }
                    if (j == 5)
                    {
                        sheet.Cells[i, 5].Value = value[j];
                    }
                    if (j == 7)
                    {
                        sheet.Cells[i, 7].Value = value[j];
                    }
                    if (j == 8)
                    {
                        sheet.Cells[i, 8].Value = value[j];
                    }
                    if (j == 9)
                    {
                        sheet.Cells[i, 9].Value = value[j];
                    }
                    if (j == 10)
                    {
                        sheet.Cells[i, 10].Value = value[j];
                    }
                    if (j == 11)
                    {
                        sheet.Cells[i, 11].Value = value[j];
                    }
                    if (j == 16)
                    {
                        sheet.Cells[i, 16].Value = value[j];
                    }
                    if (j == 14)
                    {
                        sheet.Cells[1, 14].Value = value[j];
                    }
                    if (j == 15)
                    {
                        sheet.Cells[i, 15].Value = value[j];
                    }
                    if (j == 18)
                    {
                        sheet.Cells[i, 18].Value = value[j];
                    }
                    if (j == 19)
                    {
                        sheet.Cells[i, 19].Value = value[j];
                    }
                    if (j == 20)
                    {
                        sheet.Cells[i, 20].Value = value[j];
                    }
                    if (j == 21)
                    {
                        sheet.Cells[i, 21].Value = value[j];
                    }
                    if (j == 22)
                    {
                        sheet.Cells[i, 22].Value = value[j];
                    }
                    if (j == 23)
                    {
                        sheet.Cells[i, 23].Value = value[j];
                    }
                    if (j == 4)

                        if (j == 25)
                        {
                            sheet.Cells[i, 25].Value = value[j];
                        }
                    if (j == 24)
                    {
                        sheet.Cells[i, 24].Value = value[j];
                    }
                    if (j == 12)
                    {
                        sheet.Cells[i, 12].Value = value[j];
                    }
                    if (j == 17)
                    {
                        sheet.Cells[i, 17].Value = value[j];
                    }
                    if (j == 6)
                    {
                        sheet.Cells[i, 6].Value = value[j];
                    }
                    if (j == 13)
                    {
                        sheet.Cells[i, 13].Value = value[j];
                    }
                    if (j == 26)
                    {
                        sheet.Cells[i, 26].Value = value[j];
                    }
                    if (j == 27)
                    {
                        sheet.Cells[i, 27].Value = value[j];
                    }
                    if (j == 28)
                    {
                        sheet.Cells[i, 28].Value = value[j];
                    }
                    if (j == 29)
                    {
                        sheet.Cells[i, 29].Value = value[j];
                    }
                    if (j == 30)
                    {
                        sheet.Cells[i, 30].Value = value[j];
                    }
                    if (j == 31)
                    {
                        sheet.Cells[i, 31].Value = value[j];
                    }
                    sheet.Cells[i, 32].Value = "";
                    sheet.Cells[i, 33].Value = "";
                    sheet.Cells[i, 34].Value = "";
                    sheet.Cells[i, 35].Value = "";
                    sheet.Cells[i, 36].Value = "";
                    sheet.Cells[i, 37].Value = "";
                    sheet.Cells[i, 38].Value = "";
                    sheet.Cells[i, 39].Value = "";
                    sheet.Cells[i, 40].Value = "";
                    sheet.Cells[i, 41].Value = "";
                    sheet.Cells[i, 42].Value = "";

                }
            }
        }

        private void GerarBlocoR0150(out int i, out string caminho, out string path)
        {
            // Inicio bloco 0150
            StreamWriter x;

            caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            path = caminho + @"\Bloco0150.txt";
            x = File.CreateText(path);

            // Títulos
            x.WriteLine("|Registro|Código|Razão Social|Código Pais|CNPJ|" +
                "CPF|IE|Municipio|SUFRAMA|ENDEREÇO|Numero|Complemento|Bairro|");

            i = 2;
            //int arrayTotal;
            foreach (string y in R150)
            {
                x.WriteLine(y);
                i++;
            }
        }

        private void GerarBlocoR0200(ExcelWorksheet sheet, out int i, int num, out string caminho, out string path)
        {
            // Inicio bloco 0200
            StreamWriter x;

            caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            path = caminho + @"\Bloco0200.txt";
            x = File.CreateText(path);

            // Títulos
            i = 1;
            x.WriteLine("|Código||||||NCM|||||Reg|");

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
                            sheet.Cells[i, 12].Value = value[num];
                        }

                        if ((value[num].Length == 1 || value[num].Length == 2 ||
                            value[num].Length == 3) && !Regex.IsMatch(value[num], @"^[0-9]+$"))
                        {
                            sheet.Cells[i, 5].Value = value[num];
                        }

                        if (value[num].Length == 2 && Regex.IsMatch(value[num], @"^[0-9]+$"))
                        {
                            sheet.Cells[i, 5].Value = value[num];
                        }
                        if (value[num].Length >= 7)
                        {
                            if (value[num].Length == 8 && Regex.IsMatch(value[num], @"^[0-9]+$"))
                            {
                                sheet.Cells[i, 6].Value = value[num];
                            }
                            else
                            {
                                string[] count = value[num].Split(' ');
                                if (count.Count() == 1)
                                {
                                    sheet.Cells[i, 1].Value = value[num];
                                    calcR200.Add(value[num]);
                                }
                                else
                                {
                                    sheet.Cells[i, 2].Value = value[num];
                                }
                            }
                        }
                    }

                    if (sheet.Cells[i, 1].Value != null)
                    {
                        x.WriteLine("|" + sheet.Cells[i, 1].Value + "|" + sheet.Cells[i, 2].Value + "|||" + sheet.Cells[i, 5].Value + "|" + sheet.Cells[i, 6].Value + "|" + sheet.Cells[i, 6].Value + "||||||" + sheet.Cells[i, 12].Value + "|");
                    }
                    else
                    {
                        x.WriteLine("||" + sheet.Cells[i, 2].Value + "|||" + sheet.Cells[i, 5].Value + "|" + sheet.Cells[i, 6].Value + "|" + sheet.Cells[i, 6].Value + "||||||" + sheet.Cells[i, 12].Value + "|");
                    }
                    num++;
                }
                i++;
            }

        }

        private void GerarBLocoA(ExcelWorksheet sheet, out int i, out int num, string caminho, string path)
        {
            // Inicio bloco A
            StreamWriter x;


            caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            path = caminho + @"\BlocoA.txt";
            x = File.CreateText(path);

            x.WriteLine("|Código|Código Geral|REG|IND_OPER|IND_EMIT|" +
                "COd_PART|COD_SIT|SER|SUB|NUM_DOC|CHV_NFSE|" +
                "DT_DOC|DT_EXE_SERV|VL_DOC|IND_PGTO|VL_DESC|VL_BC_PIS|" +
                "VL_PIS|VL_BC_CONFINS|VL_PIS_RET|VL_CONFINS_RE|VL_ISS|");

            // Valores
            i = 2;
            num = 0;
            string calcA;
            string calcB;

            foreach (string y in A100)
            {
                num = 0;
                string[] value = y.Split('|');//.Where(x => x != "");
                for (int j = 2; j <= value.Count(); j++)
                {
                    sheet.Cells[i, j].Value = value[num];

                    num++;
                }

                calcA = null;
                calcB = null;

                if (y.Contains("A100"))
                {
                    if (sheet.Cells[i, 5].Value != null && sheet.Cells[i, 4].Value != null &&
                        sheet.Cells[i, 10].Value != null && sheet.Cells[i, 12].Value != null
                        && sheet.Cells[i, 6].Value != null)
                    {
                        if (sheet.Cells[i, 8].Value != null)
                            calcA = "" + sheet.Cells[i, 5].Value + "" + sheet.Cells[i, 4].Value + "" + sheet.Cells[i, 8].Value + "" + sheet.Cells[i, 10].Value.ToString().PadLeft(9, '0') + "" + sheet.Cells[i, 12].Value + "" + sheet.Cells[i, 6].Value;
                        else
                            calcA = "" + sheet.Cells[i, 5].Value + "" + sheet.Cells[i, 4].Value + " " + sheet.Cells[i, 10].Value.ToString().PadLeft(9, '0') + "" + sheet.Cells[i, 12].Value + "" + sheet.Cells[i, 6].Value;
                    }
                    else
                    {
                        calcA = "";
                    }
                }
                else
                {
                    calcA = "";
                }

                if (calcA != "")
                {
                    calcB = calcA;
                }
                else
                {
                    calcB = "";
                }

                calcTempB.Add(calcB);
                x.WriteLine("|" + calcA + "|" + calcB + y);

                i++;
            }
        }

        private void GerarBlocoC(ExcelWorksheet sheet, out int i, out string caminho, out string path)
        {
            // inicio para criar o bloco C
            StreamWriter x;

            caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            path = caminho + @"\BlocoC.txt";
            x = File.CreateText(path);

            x.WriteLine("Código|Código|1 - Registro|2 - OPERAÇÃO|3 - TIPO EMITENTE|" +
                "4 - PARTICIPANTE|5 - MODELO NFE|6 - SITUAÇÃO TRIBUTARIA|7 - S**ERIE|8 - NUMERO|9 - CHAVE NFE|" +
                "10 - DT.EMISSÃO|11 - DT.SAIDA|12 - VL.TOTAL|13 - TP.PAGTO|14 - VL.DESCONTO|15 - ABATIMENTOS ZFM|" +
                "16 - VL.MERCADORIA|17 - FRETE|18 - VL.FRETE|19 - VL.SEGURO|20 - VL.DESP.ACESS|21 - VL.BASE ICMS|" +
                "22 - VL ICMS|23 - VL.BASE.ICMS.ST|24 - VL.ICMS.ST|25 - VALOR IPI|26 - VALOR PIS|27 - VALOR COFINS|" +
                "28 - VL.PIS.ST|29 - VL.COFINS.ST|30|31|32|33|34|35|36|37|38|39|40|41|42");


            // Valores
            i = 2;
            string calcA;
            string calcB;
            int count = 0;
            foreach (string y in C100)
            {
                calcA = null;
                calcB = null;

                string[] value = y.Split('|');//.Where(x => x != "");
                for (int j = 2; j < value.Count(); j++)
                {
                    if (value.Length > 15)
                        sheet.Cells[i, j].Value = value[j];
                    else
                        break;
                }

                if (value.Length > 15)
                {
                    if (sheet.Cells[i, 3].Value != null && sheet.Cells[i, 2].Value != null &&
                       sheet.Cells[i, 5].Value != null && sheet.Cells[i, 8].Value != null &&
                       sheet.Cells[i, 7].Value != null && sheet.Cells[i, 10].Value != null &&
                       sheet.Cells[i, 11].Value != null && sheet.Cells[i, 4].Value != null)
                    {
                        if (y.Contains("C100"))
                        {
                            if (!string.IsNullOrEmpty(sheet.Cells[i, 7].Value.ToString()))
                            {
                                calcA = "" + sheet.Cells[i, 3].Value + "" + sheet.Cells[i, 2].Value + "" +
                                    sheet.Cells[i, 5].Value + "" + sheet.Cells[i, 8].Value.ToString().PadLeft(9, '0') + "" +
                                    sheet.Cells[i, 7].Value.ToString().PadLeft(3, '0') + sheet.Cells[i, 10].Value +
                                    sheet.Cells[i, 11].Value + sheet.Cells[i, 4].Value;
                            }
                            else
                            {
                                calcA = "" + sheet.Cells[i, 3].Value + "" + sheet.Cells[i, 2].Value + "" +
                                    sheet.Cells[i, 5].Value + "" + sheet.Cells[i, 8].Value.ToString().PadLeft(9, '0') + "   " +
                                    sheet.Cells[i, 10].Value + sheet.Cells[i, 11].Value + sheet.Cells[i, 4].Value;

                            }
                        }
                        else
                        {
                            calcA = "";
                        }
                    }
                    else
                    {
                        calcA = "";
                    }

                    if (calcA != "")
                    {
                        calcB = calcA;
                        count = 0;
                    }
                    else
                    {
                        if (y.Contains("C100"))
                        {
                            count = 0;
                        }
                        else
                        {
                            count += 1;
                        }
                        if (count != 0)
                        {
                            if (!string.IsNullOrEmpty(sheet.Cells[i - count, 7].Value.ToString()))
                            {
                                calcB = "" + sheet.Cells[i - count, 3].Value + "" + sheet.Cells[i - count, 2].Value + "" +
                                    sheet.Cells[i - count, 5].Value + "" + sheet.Cells[i - count, 8].Value.ToString().PadLeft(9, '0') + "" +
                                    sheet.Cells[i - count, 7].Value.ToString().PadLeft(3, '0') + sheet.Cells[i - count, 10].Value +
                                    sheet.Cells[i - count, 11].Value + sheet.Cells[i - count, 4].Value;
                            }
                            else
                            {
                                calcB = "" + sheet.Cells[i - count, 3].Value + "" + sheet.Cells[i - count, 2].Value + "" +
                                    sheet.Cells[i - count, 5].Value + "" + sheet.Cells[i - count, 8].Value.ToString().PadLeft(9, '0') + "   " +
                                    sheet.Cells[i - count, 10].Value + sheet.Cells[i - count, 11].Value + sheet.Cells[i - count, 4].Value;

                            }
                            calcTempC.Add(calcB);
                        }
                    }
                    x.WriteLine("|" + calcA + "|" + calcB + y);
                }
                else
                {
                    count += 1;
                }

                i++;
            }
        }

        private string CarregaListaDados(ref int counter)
        {
            string line;
            // Read the file and display it line by line.  
            System.IO.StreamReader file =
                new System.IO.StreamReader(openFileDialog1.FileName);

            int count;
            while ((line = file.ReadLine()) != null)
            {
                string[] palavra = line.Split('|');
                count = 0;
                foreach (var X in palavra)
                {
                    count++;
                    if (line.Contains("|0000|") && count == 8)
                    {
                        R0.Add(X);
                    }
                }

                if (line.Contains("|C100|") || line.Contains("|C170|") || line.Contains("|C113|") || line.Contains("|C120|"))
                {
                    C100.Add(line);
                }

                if (line.Contains("|0150|"))
                {
                    R150.Add(line);
                }

                if (line.Contains("|A100|") || line.Contains("|A170|") || line.Contains("|A120|"))
                {
                    A100.Add(line);
                }

                if (line.Contains("|0200|"))
                {
                    R200.Add(line);
                }

                if (line.Contains("|C113|"))
                {
                    C113.Add(line);
                }

                if (line.Contains("|1100|") || (line.Contains("|1105|")))
                {
                    R1100.Add(line);
                }

                counter++;
            }
            file.Close();

            MessageBox.Show("Carregamento finalizado, favor carregar o arquivo do PIS Confins!", "Carregamento Concluido", MessageBoxButtons.OK);

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

        private void btnCarregar_Click(object sender, EventArgs e)
        {
            int counter = 0;
            string line;

            line = CarregaListaDados(ref counter);

            txtArquivo.Enabled = true;
            btnSearch1.Enabled = true;
            carregado = true;
        }

        private void btnLoadArch_Click(object sender, EventArgs e)
        {
            string line;
            // Read the file and display it line by line.  
            System.IO.StreamReader file =
                new System.IO.StreamReader(openFileDialog2.FileName);

            bool range = false;

            for (int i = 0; i < R0.Count; i++)
            {
                while ((line = file.ReadLine()) != null)
                {

                    if (range)
                    {
                        if (line.Contains("|C100|") || line.Contains("|C170|") || line.Contains("|C113|") || line.Contains("|C120|"))
                        {
                            C100.Add(line);
                        }

                        if (line.Contains("|0150|"))
                        {
                            R150.Add(line);
                        }

                        if (line.Contains("|A100|") || line.Contains("|A170|") || line.Contains("|A120|"))
                        {
                            A100.Add(line);
                        }

                        if (line.Contains("|0200|"))
                        {
                            R200.Add(line);
                        }

                        if (line.Contains("|C113|"))
                        {
                            C113.Add(line);
                        }

                        if (line.Contains("|1100|") || (line.Contains("|1105|")))
                        {
                            R1100.Add(line);
                        }

                        if (line.Contains("|A010|") || line.Contains("|C010|"))
                        {
                            range = false;
                            continue;
                        }

                    }

                    if (line.Contains("|A010|" + R0[i] + "|") || line.Contains("|C010|" + R0[i] + "|"))
                    {
                        range = true;
                    }
                }
            }

            file.Close();
            
            MessageBox.Show("Carregamento finalizado, favor Realizar a conversão dos dados!", "Carregamento Concluido", MessageBoxButtons.OK);
            carregadoPis = true;
        }

    }
}
