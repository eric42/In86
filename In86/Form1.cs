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
using System.Diagnostics;

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
        List<string> C170Temp = new List<string>();
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
            RemoveDuplicada();

            if (carregado && carregadoPis)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var excelPackage = new ExcelPackage())
                {
                    excelPackage.Workbook.Properties.Title = "IN86";
                    var sheet = excelPackage.Workbook.Worksheets.Add("Bloco A");

                    int i, num;
                    //GeraC170NrDocumento();
                    string caminho, path;

                    float calcFatValRBNCT, calcFatValRBNCNT, calcFatValRBNCE, calcIndRBNCT, calcIndRBNCNT, calcIndRBNCE, calcFatValTotal, calcFacIndTotal, calcPercIndRBNCT, calcPercIndRBNCNT, calcPercIndRBNCE;

                    CalculoFaturamento(out calcFatValRBNCT, out calcFatValRBNCNT, out calcFatValRBNCE, out calcIndRBNCT, out calcIndRBNCNT, out calcIndRBNCE, out calcFatValTotal, out calcFacIndTotal, out calcPercIndRBNCT, out calcPercIndRBNCNT, out calcPercIndRBNCE);

                    GerarBlocoC(sheet, out i, out caminho, out path);

                    GerarBLocoA(sheet, out i, out num, caminho, path);

                    GerarBlocoR0200(sheet, out i, num, out caminho, out path);

                    GerarBlocoR0150(out i, out caminho, out path);

                    //revisar os campos
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


                    calcF = calcA + calcB.PadRight(5, ' ') + calcC + calcD + calcE.Replace(",", "");

                    x.WriteLine(calcF);

                }
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
                    calcG = calcA + calcB.PadRight(5, ' ') + calcC.PadRight(9, ' ') + calcD + calcE + calcF.Replace(",", "");

                    x.WriteLine(calcG);
                    x.Flush();
                }
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
                    calcA = calcTempB[countTemp].Substring(2, 3);
                    calcB = calcTempB[countTemp].Substring(7, 9);
                    calcC = calcTempB[countTemp].Substring(14, 8);
                    calcD = calcTempB[countTemp].Substring(22, 7);
                    calcE = sheet.Cells[i, 3].Value.ToString();
                    calcE = calcE.Substring(0, 3);
                    calcF = sheet.Cells[i, 9].Value.ToString();
                    calcG = string.Format(@"{0:0,0000}", sheet.Cells[i, 11].Value.ToString());
                    calcG = calcG.Replace(",", "");
                    calcH = string.Format(@"{0:0,000}", sheet.Cells[i, 10].Value.ToString());
                    calcH = calcH.Replace(",", "");
                    calcI = string.Format(@"{0:f}", float.Parse(sheet.Cells[i, 10].Value.ToString()) * calcIndRBNCE);
                    calcI = calcI.Replace(",", "");
                    calcJ = string.Format(@"{0:f}", float.Parse(sheet.Cells[i, 10].Value.ToString()) * calcIndRBNCT);
                    calcJ = calcJ.Replace(",", "");
                    calcK = string.Format(@"{0:f}", float.Parse(sheet.Cells[i, 10].Value.ToString()) * calcIndRBNCNT);
                    calcK = calcK.Replace(",", "");
                    calcL = string.Format(@"{0:f}", sheet.Cells[i, 12].Value.ToString());
                    calcL = calcL.Replace(",", "");
                    calcN = string.Format(@"{0:0,0000}", sheet.Cells[i, 14].Value.ToString());
                    calcN = calcN.Replace(",", "");
                    calcS = string.Format(@"{0:f}", sheet.Cells[i, 16].Value.ToString());
                    calcS = calcS.Replace(",", "");
                    calcP = string.Format(@"{0:f}", float.Parse(calcS) * calcIndRBNCE);
                    calcP = calcP.Replace(",", "");
                    calcQ = string.Format(@"{0:f}", float.Parse(calcS) * calcIndRBNCT);
                    calcQ = calcQ.Replace(",", "");
                    calcR = string.Format(@"{0:f}", float.Parse(calcS) * calcIndRBNCNT);
                    calcR = calcR.Replace(",", "");
                    calcT = calcTempB[countTemp].Substring(15, 8);
                    savedCount = false;

                    countTemp = usualCount + 1;
                    usualCount = countTemp;

                    if (countTemp != calcTempB.Count())
                    {
                        foreach (var ind in calcTempB)
                        {
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
                    calcO = calcO.Replace(",", "");

                    calcU = calcA.PadRight(5, ' ') + calcB.PadLeft(9, '0') + calcC.PadLeft(8, '0') + calcD.PadRight(14, ' ') + calcE.PadLeft(3, '0') + calcF.PadRight(2, ' ') + calcG.PadLeft(8, '0') + calcH.PadLeft(17, '0') +
                        calcI.PadLeft(17, '0') + calcJ.PadLeft(17, '0') + calcK.PadLeft(17, '0') + calcL.PadLeft(17, '0') + calcM.PadRight(2, ' ') + calcN.PadLeft(8, '0') + calcO.PadLeft(17, '0') +
                        calcP.PadLeft(17, '0') + calcQ.PadLeft(17, '0') + calcR.PadLeft(17, '0') + calcS.PadLeft(17, '0') + calcT.PadLeft(8, '0');



                    x.WriteLine(calcU);
                    x.Flush();

                }
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

            i = 2;
            string calcA, calcB, calcC, calcD, calcE, calcF, calcG, calcH, calcI, calcJ, calcK, calcL, calcM, calcN, calcO, calcP, calcQ, calcR, calcS, calcT, calcU, calcV;
            int countTemp = 0;

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

                if (value.Length >= 30)
                {
                    if (sheet.Cells[i, 1].Value.ToString() == "C170")
                    {
                        if (calcTempC[countTemp].ToString().Substring(0, 1) == "1")
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
                            calcH = calcH.Replace(",", "");
                            calcI = string.Format(@"{0:0,0000}", (sheet.Cells[i, 26].Value ?? "0,0000").ToString());
                            calcI = calcI.Replace(",", "");
                            calcM = string.Format(@"{0:f}", (sheet.Cells[i, 30].Value ?? "0,00").ToString());
                            calcM = calcM.Replace(",", "");
                            calcO = string.Format(@"{0:f}", sheet.Cells[i, 33].Value.ToString());
                            calcO = calcO.Replace(",", "");
                            calcT = string.Format(@"{0:f}", (sheet.Cells[i, 36].Value.ToString() == "" ? "0,00" : sheet.Cells[i, 36].Value).ToString());
                            calcT = calcT.Replace(",", "");
                            calcU = calcTempC[countTemp].Substring(25, 8);

                            countTemp++;

                        }

                        calcN = calcG;
                        calcN = calcN.Replace(",", "");
                        calcP = calcI;
                        calcP = calcP.Replace(",", "");

                        if (calcG == "50")
                        {
                            calcJ = "0,00";
                            calcJ = calcJ.Replace(",", "");
                            calcL = "0,00";
                            calcL = calcL.Replace(",", "");
                        }
                        else
                        {
                            if (calcTempC[countTemp].ToString().Substring(0, 1) == "1" && sheet.Cells[i, 1].Value.ToString() == "C170")
                            {
                                if (calcM != "")
                                {
                                    calcJ = string.Format(@"{0:f}", float.Parse(calcM) * calcIndRBNCE);
                                    calcJ = calcJ.Replace(",", "");
                                    calcL = string.Format(@"{0:f}", float.Parse(calcM) * calcIndRBNCNT);
                                    calcL = calcL.Replace(",", "");
                                }
                                else
                                {
                                    calcJ = "0,00";
                                    calcJ = calcJ.Replace(",", "");
                                    calcL = "0,00";
                                    calcL = calcL.Replace(",", "");
                                }
                            }
                        }

                        if (calcG == "50")
                        {
                            calcK = calcM;
                            calcK = calcK.Replace(",", "");
                        }
                        else
                        {
                            if (calcTempC[countTemp].ToString().Substring(0, 1) == "1" && sheet.Cells[i, 1].Value.ToString() == "C170")
                            {
                                if (calcM != "")
                                {
                                    calcK = string.Format(@"{0:f}", float.Parse(calcM) * calcPercIndRBNCT);
                                    calcK = calcK.Replace(",", "");
                                }
                                else
                                {
                                    calcK = "0,00";
                                    calcK = calcK.Replace(",", "");
                                }
                            }
                        }

                        if (calcN == "50")
                        {
                            calcQ = "0,00";
                            calcQ = calcQ.Replace(",", "");
                            calcS = "0,00";
                            calcS = calcS.Replace(",", "");
                        }
                        else
                        {
                            if (calcTempC[countTemp].ToString().Substring(0, 1) == "1" && sheet.Cells[i, 1].Value.ToString() == "C170")
                            {
                                calcQ = string.Format(@"{0:f}", float.Parse(calcT) * calcIndRBNCE);
                                calcQ = calcQ.Replace(",", "");
                                calcS = string.Format(@"{0:f}", float.Parse(calcT) * calcIndRBNCNT);
                                calcS = calcS.Replace(",", "");
                            }
                        }

                        if (calcN == "50")
                        {
                            calcR = calcT;
                            calcR = calcR.Replace(",", "");
                        }
                        else
                        {
                            if (calcTempC[countTemp].ToString().Substring(0, 1) == "1" && sheet.Cells[i, 1].Value.ToString() == "C170")
                            {
                                calcR = string.Format(@"{0:f}", float.Parse(calcT) * calcIndRBNCT);
                                calcR = calcR.Replace(",", "");
                            }
                        }


                        calcV = calcA.PadRight(2, ' ') + calcB.PadRight(5, ' ') + calcC.PadLeft(9, '0') + calcD.PadLeft(8, '0') + calcE.PadRight(14, ' ') + calcF.PadLeft(3, '0') + calcG.PadRight(2, ' ') + calcH.PadLeft(8, '0') + calcI.PadLeft(17, '0') +
                            calcJ.PadLeft(17, '0') + calcK.PadLeft(17, '0') + calcL.PadLeft(17, '0') + calcM.PadLeft(17, '0') + calcN.PadRight(2, ' ') + calcO.PadLeft(8, '0') + calcP.PadLeft(17, '0') +
                            calcQ.PadLeft(17, '0') + calcR.PadLeft(17, '0') + calcS.PadLeft(17, '0') + calcT.PadLeft(17, '0') + calcU.PadLeft(8, '0');

                        x.WriteLine(calcV);
                        x.Flush();

                    }
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

            i = 2;
            string calcA, calcB, calcC, calcD, calcE, calcF, calcG, calcH, calcI, calcJ, calcK, calcL, calcM, calcN, calcO, calcP, calcQ, calcR, calcS, calcT, calcU;
            int countTemp = 0;

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

                if (value.Length >= 30)
                {

                    if (sheet.Cells[i, 1].Value.ToString() == "C170")
                    {
                        if (calcTempC[countTemp].ToString().Substring(0, 1) == "1")
                        {
                            calcA = calcTempC[countTemp].Substring(2, 2);
                            calcB = calcTempC[countTemp].Substring(14, 3);
                            calcC = calcTempC[countTemp].Substring(5, 9);
                            calcD = calcTempC[countTemp].Substring(14, 8);
                            calcE = string.IsNullOrEmpty(sheet.Cells[i, 2].Value.ToString()) ? sheet.Cells[i, 2].Value.ToString().PadLeft(3, '0') : "000";
                            calcF = sheet.Cells[i, 25].Value.ToString();
                            calcG = string.Format(@"{0:0,0000}", (sheet.Cells[i, 27].Value ?? "0,0000").ToString());
                            calcG = calcG.Replace(",", "");
                            calcH = string.Format(@"{0:0,0000}", (sheet.Cells[i, 26].Value ?? "0,0000").ToString());
                            calcH = calcH.Replace(",", "");

                            countTemp++;

                        }

                        calcM = calcF;
                        calcO = calcH;
                        calcO = calcO.Replace(",", "");

                        if (calcTempC[countTemp].ToString().Substring(0, 1) == "0" && sheet.Cells[i, 1].Value.ToString() == "C170"
                            && calcTempC[countTemp].ToString().Substring(1, 1) == "0")
                        {
                            calcL = sheet.Cells[i, 30].Value.ToString();
                            calcL = calcL.Replace(",", "");
                            calcN = sheet.Cells[i, 33].Value.ToString();
                            calcN = calcN.Replace(",", "");
                            calcS = sheet.Cells[i, 36].Value.ToString();
                            calcS = calcS.Replace(",", "");
                            calcT = calcTempC[countTemp].Substring(25, 8);

                        }

                        if (calcF == "50")
                        {
                            calcI = "0,00";
                            calcI = calcI.Replace(",", "");
                            calcK = "0,00";
                            calcK = calcK.Replace(",", "");
                        }
                        else
                        {
                            if (calcTempC[countTemp].ToString().Substring(0, 1) == "0" && sheet.Cells[i, 1].Value.ToString() == "C170"
                                && calcTempC[countTemp].ToString().Substring(1, 1) == "0")
                            {
                                if (calcL == "")
                                {
                                    calcL = "0,00";
                                    calcL = calcL.Replace(",", "");
                                }
                                calcI = string.Format(@"{0:f}", float.Parse(calcL) * calcIndRBNCE);
                                calcI = calcI.Replace(",", "");
                            }

                            if (calcF == "50")
                            {
                                calcK = "0,00";
                                calcK = calcK.Replace(",", "");
                            }
                            else
                            {
                                if (calcTempC[countTemp].ToString().Substring(0, 1) == "0" && sheet.Cells[i, 1].Value.ToString() == "C170"
                                && calcTempC[countTemp].ToString().Substring(1, 1) == "0")
                                {
                                    calcK = string.Format(@"{0:f}", float.Parse(calcL) * calcIndRBNCNT);
                                    calcK = calcK.Replace(",", "");
                                }
                            }

                        }

                        if (calcF == "50")
                        {
                            calcJ = calcL;
                            calcJ = calcJ.Replace(",", "");
                        }
                        else
                        {
                            if (calcTempC[countTemp].ToString().Substring(0, 1) == "0" && sheet.Cells[i, 1].Value.ToString() == "C170"
                                && calcTempC[countTemp].ToString().Substring(1, 1) == "0")
                            {
                                calcJ = string.Format(@"{0:f}", float.Parse(calcL) * calcIndRBNCT);
                                calcJ = calcJ.Replace(",", "");
                            }
                        }

                        if (calcM == "50")
                        {
                            calcP = "0,00";
                            calcP = calcP.Replace(",", "");
                            calcR = "0,00";
                            calcR = calcR.Replace(",", "");
                        }
                        else
                        {
                            if (calcTempC[countTemp].ToString().Substring(0, 1) == "0" && sheet.Cells[i, 1].Value.ToString() == "C170"
                                && calcTempC[countTemp].ToString().Substring(1, 1) == "0")
                            {

                                if (calcS == "")
                                {
                                    calcS = "0,00";
                                    calcS = calcS.Replace(",", "");
                                }
                                calcP = string.Format(@"{0:f}", float.Parse(calcS) * calcIndRBNCE);
                                calcP = calcP.Replace(",", "");
                                calcR = string.Format(@"{0:f}", float.Parse(calcS) * calcIndRBNCNT);
                                calcR = calcR.Replace(",", "");
                            }
                        }

                        if (calcM == "50")
                        {
                            calcQ = calcS;
                            calcQ = calcQ.Replace(",", "");
                        }
                        else
                        {
                            if (calcTempC[countTemp].ToString().Substring(0, 1) == "0" && sheet.Cells[i, 1].Value.ToString() == "C170"
                                && calcTempC[countTemp].ToString().Substring(1, 1) == "0")
                            {
                                calcQ = string.Format(@"{0:f}", float.Parse(calcS) * calcIndRBNCT);
                                calcQ = calcQ.Replace(",", "");
                            }
                        }

                        calcU = calcA.PadRight(2, ' ') + calcB.PadRight(5, ' ') + calcC.PadLeft(9, '0') + calcD.PadLeft(8, '0') + calcE.PadLeft(3, '0') + calcF.PadRight(2, ' ') + calcG.PadLeft(8, '0') + calcH.PadLeft(17, '0') + calcI.PadLeft(17, '0') +
                            calcJ.PadLeft(17, '0') + calcK.PadLeft(17, '0') + calcL.PadLeft(17, '0') + calcM.PadRight(2, ' ') + calcN.PadLeft(8, '0') + calcO.PadLeft(17, '0') + calcP.PadLeft(17, '0') +
                            calcQ.PadLeft(17, '0') + calcR.PadLeft(17, '0') + calcS.PadLeft(17, '0') + calcT.PadLeft(8, '0');

                        x.WriteLine(calcU);
                        x.Flush();

                    }
                }

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

            i = 2;
            string calcA, calcB, calcC, calcD, calcE, calcF, calcG, calcH, calcI, calcJ, calcK, calcL, calcM, calcN, calcO;
            int countTemp = 0;

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


                if (value.Length >= 30)
                {
                    if (sheet.Cells[i, 1].Value.ToString() == "C170")
                    {
                        if (calcTempC[countTemp].ToString().Substring(0, 1) == "1")
                        {
                            calcA = calcTempC[countTemp].Substring(2, 2);
                            calcB = calcTempC[countTemp].Substring(14, 3);
                            calcD = calcTempC[countTemp].Substring(17, 8);
                            calcE = string.IsNullOrEmpty(sheet.Cells[i, 2].Value.ToString()) ? sheet.Cells[i, 2].Value.ToString().PadLeft(3, '0') : "000";
                            calcF = sheet.Cells[i, 25].Value.ToString();
                            calcG = string.Format(@"{0:0,0000}", (sheet.Cells[i, 27].Value ?? "0,0000").ToString());
                            calcG = calcG.Replace(",", "");
                            calcH = string.Format(@"{0:0,0000}", (sheet.Cells[i, 26].Value ?? "0,0000").ToString());
                            calcH = calcH.Replace(",", "");
                            calcI = string.Format(@"{0:f}", (sheet.Cells[i, 30].Value ?? "0,00").ToString());
                            calcI = calcI.Replace(",", "");
                            calcK = string.Format(@"{0:0,0000}", (sheet.Cells[i, 33].Value ?? "0,0000").ToString());
                            calcK = calcK.Replace(",", "");
                            calcM = string.Format(@"{0:f}", (sheet.Cells[i, 36].Value ?? "0,00").ToString());
                            calcM = calcM.Replace(",", "");
                            calcN = calcTempC[countTemp].Substring(25, 8);
                            countTemp++;

                            calcJ = calcF;
                            calcJ = calcJ.Replace(",", "");
                            calcL = calcH;


                            calcO = calcA.PadRight(2, ' ') + calcB.PadRight(5, ' ') + calcC.PadLeft(9, '0') + calcD.PadLeft(8, '0') + calcE.PadLeft(3, '0') + calcF.PadRight(2, ' ') + calcG.PadLeft(8, '0') + calcH.PadLeft(17, '0') + calcI.PadLeft(17, '0') + calcJ.PadRight(2, ' ') +
                                calcK.PadLeft(8, '0') + calcL.PadLeft(17, '0') + calcM.PadLeft(17, '0') + calcN.PadLeft(8, '0');

                            x.WriteLine(calcO);
                            x.Flush();

                        }
                    }

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

            i = 2;
            string calcA, calcB, calcC, calcD, calcE, calcF, calcG, calcH, calcI, calcJ, calcK, calcL, calcM;
            int countTemp = 0, countSave = 0, usualCount = 0, size = 0;
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
                    calcA = calcTempB[countTemp].Substring(2, 3);
                    calcB = calcTempB[countTemp].Substring(7, 9);
                    calcC = calcTempB[countTemp].Substring(14, 8);
                    calcD = calcTempB[countTemp].Substring(22, 7);
                    calcE = sheet.Cells[i, 2].Value.ToString().PadLeft(3, '0');
                    calcF = sheet.Cells[i, 3].Value.ToString();
                    calcH = string.Format(@"{0:f}", sheet.Cells[i, 5].Value.ToString());
                    calcH = calcH.Replace(",", "");
                    calcI = string.Format(@"{0:f}", sheet.Cells[i, 6].Value.ToString());
                    calcI = calcI.Replace(",", "");
                    calcL = string.Format(@"{0:f}", sheet.Cells[i, 21].Value.ToString());
                    calcL = calcL.Replace(",", "");
                    savedCount = false;

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
                    calcK = calcK.Replace(",", "");

                    if (!string.IsNullOrEmpty(calcK) && !string.IsNullOrEmpty(calcL) &&
                        calcK != "FALSO" && calcL != "FALSO")
                    {
                        calc = (float.Parse(calcL) / float.Parse(calcK)) * 100;
                        calcJ = string.Format(@"{0:f}", calc);
                        calcJ = calcJ.Replace(",", "");
                    }

                    size = 0;
                    if (calcJ.Count() > 5)
                    {
                        size = calcJ.Count() - 5;
                    }


                    calcM = calcA.PadRight(5, ' ') + calcB.PadLeft(9, '0') + calcC.PadLeft(8, '0') + calcD.PadRight(14, ' ') + calcE.PadLeft(3, '0') + calcF.PadRight(20, ' ') + calcG.PadLeft(45, ' ') +
                        calcH.PadLeft(17, '0') + calcI.PadLeft(17, '0') + calcJ.PadLeft(5, '0') + calcK.PadLeft(17 - size, '0') + calcL.PadLeft(17, '0');

                    x.WriteLine(calcM);
                    x.Flush();

                }
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

            i = 2;
            string calcA, calcB, calcC, calcD, calcE, calcF, calcG, calcH, calcI, calcJ;
            int size = 0;

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
                    calcA = sheet.Cells[i, 6].Value.ToString();
                    calcB = sheet.Cells[i, 8].Value.ToString();
                    calcC = sheet.Cells[i, 10].Value.ToString();
                    calcD = sheet.Cells[i, 4].Value.ToString();
                    calcE = string.Format(@"{0:f}", sheet.Cells[i, 12].Value.ToString());
                    calcE = calcE.Replace(",", "");
                    calcF = string.Format(@"{0:f}", sheet.Cells[i, 14].Value.ToString());
                    calcF = calcF.Replace(",", "");

                    calcG = "";
                    calcH = "";
                    calcI = "";

                    size = 0;
                    if (calcE.Count() > 5)
                    {
                        size = calcE.Count() - 5;
                    }

                    calcJ = calcA.PadRight(5, ' ') + calcB.PadLeft(9, '0') + calcC.PadLeft(8, '0') + calcD.PadRight(14, ' ') + calcE.PadLeft(5, '0') + calcF.PadLeft(17 - size, '0') +
                        calcG.PadLeft(17, '0') + calcH.PadLeft(5, '0') + calcI.PadLeft(17, '0');

                    x.WriteLine(calcJ);
                    x.Flush();

                }

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
                if (j == 15)
                {
                    sheet.Cells[i, 15].Value = value[j];
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

            float calc = 0;

            i = 2;
            string calcA, calcB, calcC, calcD, calcE, calcF, calcG, calcH, calcI, calcJ, calcK, calcL, calcM, calcN, calcO, calcP, calcQ, calcR, calcS, calcT, calcU, calcV, calcW, calcX, calcY, calcZ, calcAA, calcAB, calcAC, calcAD;
            int countTemp, count;
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

                CarregaBlocoCSheet40(sheet, i, value);

                count = 0;

                if (value.Length >= 30)
                {
                    if (sheet.Cells[i, 1].Value.ToString() == "C170")
                    {
                        if (calcTempC[countTemp].Substring(0, 1).Equals("1"))
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
                            calcL = calcL.Replace(",", "");
                            calcM = sheet.Cells[i, 6].Value.ToString();

                            if (calcJ.Contains(","))
                            {
                                calcJ = calcJ.Replace(",", "");
                            }

                            if (!string.IsNullOrEmpty(calcL))
                            {
                                calc = float.Parse(sheet.Cells[i, 7].Value.ToString() != "" ? sheet.Cells[i, 7].Value.ToString() : "0");
                                calcN = string.Format(@"{0:f}", calc / float.Parse(calcL));
                            }
                            else
                            {
                                calcN = "0,00";
                            }

                            calcN = calcN.Replace(",", "");
                            calcO = string.Format(@"{0:f}", sheet.Cells[i, 7].Value.ToString());
                            calcO = calcO.Replace(",", "");
                            calcP = string.Format(@"{0:f}", sheet.Cells[i, 8].Value.ToString());
                            calcP = calcP.Replace(",", "");
                            calcR = string.Format(@"{0:f}", sheet.Cells[i, 23].Value.ToString());
                            calcR = calcR.Replace(",", "");
                            calcS = string.Format(@"{0:f}", sheet.Cells[i, 22].Value.ToString());
                            calcS = calcS.Replace(",", "");
                            calcT = string.Format(@"{0:f}", sheet.Cells[i, 24].Value.ToString());
                            calcT = calcT.Replace(",", "");

                            if (sheet.Cells[i, 10].Value.ToString() != "")
                            {
                                calcU = sheet.Cells[i, 10].Value.ToString().PadLeft(3, '0');
                            }
                            else
                            {
                                calcU = "000";
                            }

                            calcW = string.Format(@"{0:f}", sheet.Cells[i, 14].Value.ToString());
                            calcW = calcW.Replace(",", "");
                            calcX = sheet.Cells[i, 13].Value.ToString();
                            calcX = calcX.Replace(",", "");
                            calcY = string.Format(@"{0:f}", sheet.Cells[i, 15].Value.ToString());
                            calcY = calcY.Replace(",", "");
                            calcZ = string.Format(@"{0:f}", sheet.Cells[i, 16].Value.ToString());
                            calcZ = calcZ.Replace(",", "");
                            calcAA = string.Format(@"{0:f}", sheet.Cells[i, 18].Value.ToString());
                            calcAA = calcAA.Replace(",", "");

                            if (sheet.Cells[i, 9].Value.ToString() == "0")
                            {
                                calcAB = "S";
                            }
                            else
                            {
                                calcAB = "N";
                            }

                            if (calcQ == "1")
                            {
                                calcAC = "00";
                            }
                            else
                            {
                                calcAC = "02";
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

                            calcAD = calcA + calcB.PadRight(5, ' ') + calcC.PadLeft(9, '0') + calcD.PadLeft(8, '0') + calcE.PadRight(14, ' ') + calcF.PadLeft(3, ' ') + calcG.PadRight(20, ' ') +
                                calcH.PadRight(45, ' ') + calcI.PadRight(4, ' ') + calcJ.PadRight(6, ' ') + calcK.PadRight(8, ' ') + calcL.PadLeft(17, '0') + calcM.PadRight(3, ' ') +
                                calcN.PadLeft(17, '0') + calcO.PadLeft(17, '0') + calcP.PadLeft(17, '0') + calcQ.PadRight(1, ' ') + calcR.PadLeft(5, '0') + calcS.PadLeft(17, '0') +
                                calcT.PadLeft(17, '0') + calcU.PadRight(3, ' ') + calcV.PadRight(1, ' ') + calcW.PadLeft(5, '0') + calcX.PadLeft(17, '0') + calcY.PadLeft(17, '0') + calcZ.PadLeft(17, '0') +
                                calcAA.PadLeft(17, '0') + calcAB.PadRight(1, ' ') + calcAC.PadRight(2, ' ');

                            x.WriteLine(calcAD);
                            x.Flush();

                        }

                    }
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
                    if (sheet.Cells[i, 1].Value.ToString() == "C100" && sheet.Cells[i, 2].Value.ToString() == "0")
                    {
                        calcA = sheet.Cells[i, 5].Value.ToString();
                        calcB = sheet.Cells[i, 7].Value.ToString().Replace("*", "").PadLeft(3, '0');
                        calcC = sheet.Cells[i, 8].Value.ToString().PadLeft(9, '0');

                        if (sheet.Cells[i, 10].Value.ToString() == "")
                        {
                            calcD = "";
                        }
                        else
                        {
                            calcD = sheet.Cells[i, 10].Value.ToString();
                        }

                        if (sheet.Cells[i, 4].Value.ToString() == "")
                        {
                            calcE = "";
                        }
                        else
                        {
                            calcE = sheet.Cells[i, 4].Value.ToString();
                        }

                        if (calcD == "")
                        {
                            calcF = "";
                        }
                        else
                        {
                            calcF = sheet.Cells[i, 11].Value.ToString();
                        }

                        calcG = string.Format(@"{0:f}", sheet.Cells[i, 16].Value.ToString());
                        calcG = calcG.Replace(",", "");
                        calcH = string.Format(@"{0:f}", sheet.Cells[i, 14].Value.ToString());
                        calcH = calcH.Replace(",", "");
                        calcI = string.Format(@"{0:f}", sheet.Cells[i, 18].Value.ToString());
                        calcI = calcI.Replace(",", "");
                        calcJ = string.Format(@"{0:f}", sheet.Cells[i, 19].Value.ToString());
                        calcJ = calcJ.Replace(",", "");
                        calcK = string.Format(@"{0:f}", sheet.Cells[i, 20].Value.ToString());
                        calcK = calcK.Replace(",", "");
                        calcL = string.Format(@"{0:f}", sheet.Cells[i, 25].Value.ToString());
                        calcL = calcL.Replace(",", "");
                        calcM = string.Format(@"{0:f}", sheet.Cells[i, 24].Value.ToString());
                        calcM = calcM.Replace(",", "");
                        calcN = string.Format(@"{0:f}", sheet.Cells[i, 12].Value.ToString());
                        calcN = calcN.Replace(",", "");
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


                        calcY = calcA + calcB.PadRight(5, ' ') + calcC + calcD.PadLeft(8, ' ').Substring(0, 8) + calcE.PadRight(14, ' ') +
                            calcF.PadLeft(8, '0').Substring(0, 8) + calcG.PadLeft(17, '0')
                            + calcH.PadLeft(17, '0') + calcI.PadLeft(17, '0') + calcJ.PadLeft(17, '0') +
                            calcK.PadLeft(17, '0') + calcL.PadLeft(17, '0') + calcM.PadLeft(17, '0') + calcN.PadLeft(17, '0') + calcO.PadRight(14, ' ') +
                            calcP.PadLeft(1, ' ') + calcQ.PadRight(45, ' ') + calcR.PadRight(50, ' ') + calcS.PadRight(2, ' ') + calcT.PadRight(5, ' ')
                            + calcU.PadLeft(9, '0') + calcV.PadLeft(8, '0') + calcW.PadRight(14, ' ');

                        x.WriteLine(calcY);
                        x.Flush();

                    }
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

            i = 2;
            string calcA, calcB, calcC, calcD, calcE, calcF, calcG, calcH, calcI, calcJ, calcK, calcL, calcM, calcN, calcO, calcP, calcQ, calcR, calcS, calcT, calcU, calcV, calcW, calcX, calcY, calcZ, calcAA, calcAB, calcAC, calcAD;
            float nCalc;
            int count = 1, countTemp, j = 0;

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
                    if (sheet.Cells[i, 1].Value.ToString() == "C170")
                    {
                        if (calcTempC[countTemp - j].Substring(0, 1).Equals("0"))
                        {
                            calcA = "E";

                            calcB = calcTempC[countTemp - j].Substring(2, 2);
                            if ((calcTempC[countTemp - j].Count()) > 20)
                            {
                                calcC = calcTempC[countTemp - j].Substring(15, 3);
                                calcE = calcTempC[countTemp - j].Substring(16, 8);
                            }
                            else
                            {
                                calcC = calcC.PadRight(5, ' ');
                                calcE = calcE.PadLeft(8, '0');
                            }

                            calcD = calcTempC[countTemp - j].Substring(4, 9);

                            calcC = calcC.Replace('*', ' ');
                        }
                        else
                        {
                            calcA = "S";

                            calcB = calcTempC[countTemp - j].Substring(2, 2);
                            calcD = calcTempC[countTemp - j].Substring(4, 9);

                            if ((calcTempC[countTemp - j].Count()) >= 16)
                            {
                                calcC = calcTempC[countTemp - j].Substring(15, 3);
                                calcE = calcTempC[countTemp - j].Substring(16, 8);
                            }
                            else
                            {
                                calcC = calcC.PadRight(5, ' ');
                                calcE = calcE.PadLeft(8, '0');
                            }

                            calcC = calcC.Replace("*", "");
                        }

                        if (!string.IsNullOrEmpty(sheet.Cells[i, 2].Value.ToString()))
                        {
                            calcF = sheet.Cells[i, 2].Value.ToString().PadLeft(3, '0');
                        }

                        if (!string.IsNullOrEmpty(sheet.Cells[i, 3].Value.ToString()))
                        {
                            calcG = sheet.Cells[i, 3].Value.ToString();
                        }

                        if (!string.IsNullOrEmpty(sheet.Cells[i, 4].Value.ToString()))
                        {
                            calcH = sheet.Cells[i, 4].Value.ToString();
                        }

                        if (!string.IsNullOrEmpty(sheet.Cells[i, 11].Value.ToString()))
                        {
                            calcI = sheet.Cells[i, 11].Value.ToString();
                        }

                        if (!string.IsNullOrEmpty(sheet.Cells[i, 12].Value.ToString()))
                        {
                            calcJ = sheet.Cells[i, 12].Value.ToString();
                        }

                        if (calcG.ToString() != null && calcG.ToString() != "FALSO")
                        {
                            foreach (string g in calcR200)
                            {
                                if (g.Equals(calcG.ToString()))
                                {
                                    calcK = g.Replace(",", "");
                                }

                                if (calcK != "")
                                    break;

                                count++;
                            }
                        }

                        if (!string.IsNullOrEmpty(sheet.Cells[i, 5].Value.ToString()))
                        {
                            calcL = string.Format(@"{0:f}", sheet.Cells[i, 5].Value.ToString());
                            calcL = calcL.Replace(",", "");
                        }
                        else
                        {
                            calcL = "000";
                        }

                        if (!string.IsNullOrEmpty(sheet.Cells[i, 6].Value.ToString()))
                        {
                            calcM = sheet.Cells[i, 6].Value.ToString();
                        }

                        if (!string.IsNullOrEmpty(sheet.Cells[i, 7].Value.ToString()) && calcL != null && calcL != "FALSO")
                        {
                            nCalc = float.Parse(sheet.Cells[i, 7].Value.ToString());
                            calcN = string.Format(@"{0:f}", nCalc / float.Parse(calcL));
                            calcN = calcN.Replace(",", "");
                        }

                        if (!string.IsNullOrEmpty(sheet.Cells[i, 7].Value.ToString()))
                        {
                            calcO = string.Format(@"{0:f}", sheet.Cells[i, 7].Value.ToString());
                            calcO = calcO.Replace(",", "");
                        }

                        if (!string.IsNullOrEmpty(sheet.Cells[i, 8].Value.ToString()))
                        {
                            calcP = string.Format(@"{0:f}", sheet.Cells[i, 8].Value.ToString());
                            calcP = calcP.Replace(",", "");
                        }

                        if (!string.IsNullOrEmpty(sheet.Cells[i, 24].Value.ToString()))
                        {
                            calcT = sheet.Cells[i, 24].Value.ToString();
                            calcT = calcT.Replace(",", "");
                        }

                        if (!string.IsNullOrEmpty(sheet.Cells[i, 23].Value.ToString()))
                        {
                            calcR = string.Format(@"{0:f}", sheet.Cells[i, 23].Value.ToString());
                            calcR = calcR.Replace(",", "");
                        }

                        if (!string.IsNullOrEmpty(sheet.Cells[i, 22].Value.ToString()))
                        {
                            calcS = string.Format(@"{0:f}", sheet.Cells[i, 22].Value.ToString());
                            calcS = calcS.Replace(",", "");
                        }

                        if (!string.IsNullOrEmpty(sheet.Cells[i, 10].Value.ToString()))
                        {
                            calcU = string.Format(@"{0:f}", sheet.Cells[i, 10].Value.ToString());
                        }

                        if (string.IsNullOrEmpty(sheet.Cells[i, 14].Value.ToString()))
                        {
                            calcW = "000";
                        }
                        else
                        {
                            calcW = string.Format(@"{0:f}", sheet.Cells[i, 14].Value.ToString());
                            calcW = calcW.Replace(",", "");
                        }

                        if (!string.IsNullOrEmpty(sheet.Cells[i, 13].Value.ToString()))
                        {
                            calcX = sheet.Cells[i, 13].Value.ToString();
                            calcX = calcX.Replace(",", "");
                        }

                        if (!string.IsNullOrEmpty(sheet.Cells[i, 15].Value.ToString()))
                        {
                            calcY = string.Format(@"{0:f}", sheet.Cells[i, 15].Value.ToString());
                            calcY = calcY.Replace(",", "");
                        }

                        if (!string.IsNullOrEmpty(sheet.Cells[i, 16].Value.ToString()))
                        {
                            calcZ = string.Format(@"{0:f}", sheet.Cells[i, 16].Value.ToString());
                            calcZ = calcZ.Replace(",", "");
                        }

                        if (!string.IsNullOrEmpty(sheet.Cells[i, 18].Value.ToString()))
                        {
                            calcAA = string.Format(@"{0:f}", sheet.Cells[i, 18].Value.ToString());
                            calcAA = calcAA.Replace(",", "");
                        }

                        if (sheet.Cells[i, 9].Value.ToString() == "0")
                        {
                            calcAB = "S";
                        }
                        else
                        {
                            calcAB = "N";
                        }


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

                        calcAD = calcA + calcB + calcC.PadRight(5, ' ') + calcD + calcE + calcF + calcG.PadRight(20, ' ') + calcH.PadRight(45, ' ') + calcI +
                            calcJ.PadRight(6, ' ') + calcK.PadLeft(17, '0') + calcL.PadLeft(17, '0') + calcM.PadRight(3, ' ').PadLeft(3, '0') +
                            calcN.PadLeft(17, '0') + calcO.PadLeft(17, '0') + calcP.PadLeft(17, '0') + calcQ +
                            calcR.PadLeft(5, '0') + calcS.PadLeft(17, '0') + calcT.PadLeft(17, '0') + calcU + calcV +
                            calcW.PadLeft(5, '0') + calcX.PadLeft(17, '0') + calcY.PadLeft(17, '0') + calcZ.PadLeft(17, '0') +
                            calcAA.PadLeft(17, '0') + calcAB + calcAC.PadRight(2, ' ');

                        x.WriteLine(calcAD);
                        x.Flush();
                    }
                }
                countTemp++;
                i++;

                if (countTemp >= calcTempC.Count() && calcTempC.Count != i)
                {
                    j++;
                }

            }
        }

        private static void CarregarSheetBlocoC(ExcelWorksheet sheet, int i, string[] value)
        {
            int count = 0;

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

            int a2 = 0;
            i = 2;
            string calcA, calcB, calcC, calcD, calcE, calcF, calcG, calcH, calcI, calcJ, calcK, calcL, calcM, calcN, calcO, calcP, calcQ, calcR, calcS, calcT, calcU, calcV, calcW, calcX, calcY, calcZ, calcAA, calcAB, calcAC, calcAD, calcAE, calcAF, calcAG;
            string final;

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
                calcAE = "";
                calcAF = "";
                calcAG = "";
                final = "";

                string[] value = y.Split('|');//.Where(x => x != "");
                CarregaBlocoCSheet40(sheet, i, value);

                if (sheet.Cells[i, 1].Value.ToString() == "C100" && sheet.Cells[i, 2].Value.ToString() == "0")
                {
                    /*if (sheet.Cells[i, 1].Value != null && sheet.Cells[i, 3].Value != null && sheet.Cells[i, 4].Value != null)
                    {*/
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

                    if (sheet.Cells[i, 7].Value != null)
                    {
                        calcC = sheet.Cells[i, 7].Value.ToString().PadLeft(3, '0').PadRight(5, ' ');
                    }

                    if (sheet.Cells[i, 8].Value != null)
                    {
                        calcD = sheet.Cells[i, 8].Value.ToString().PadLeft(9, '0');
                    }

                    if (sheet.Cells[i, 10].Value.ToString() == "")
                    {
                        calcE = "00000000";
                    }
                    else
                    {
                        calcE = sheet.Cells[i, 10].Value.ToString().PadLeft(8, '0');
                    }

                    if (sheet.Cells[i, 4].Value.ToString() == "")
                    {
                        calcF = "              ";
                    }
                    else
                    {
                        calcF = sheet.Cells[i, 4].Value.ToString().PadRight(14, ' ');
                    }

                    if (sheet.Cells[i, 11].Value.ToString() == "")
                    {
                        calcG = "00000000";
                    }
                    else
                    {
                        calcG = sheet.Cells[i, 11].Value.ToString().PadLeft(8, '0');
                    }

                    if (sheet.Cells[i, 16].Value.ToString() != null)
                    {
                        calcH = sheet.Cells[i, 16].Value.ToString().Replace(",", "");
                        calcH = calcH.PadLeft(17, '0');
                    }

                    if (sheet.Cells[i, 14].Value.ToString() != null)
                    {
                        calcI = sheet.Cells[i, 14].Value.ToString().Replace(",", "");
                        calcI = calcI.PadLeft(17, '0');
                    }

                    if (sheet.Cells[i, 18].Value.ToString() != "")
                    {
                        calcJ = sheet.Cells[i, 18].Value.ToString().Replace(",", "");
                        calcJ = calcJ.PadLeft(17, '0');
                    }

                    if (sheet.Cells[i, 19].Value.ToString() != "")
                    {
                        calcK = sheet.Cells[i, 19].Value.ToString().Replace(",", "");
                        calcK = calcK.PadLeft(17, '0');
                    }

                    if (sheet.Cells[i, 20].Value.ToString() != "")
                    {
                        calcL = sheet.Cells[i, 20].Value.ToString().Replace(",", "");
                        calcL = calcL.PadLeft(17, '0');
                    }

                    if (sheet.Cells[i, 25].Value.ToString() != "")
                    {
                        calcM = sheet.Cells[i, 25].Value.ToString().Replace(",", "");
                        calcM = calcM.PadLeft(17, '0');
                    }

                    if (sheet.Cells[i, 24].Value.ToString() != "")
                    {
                        calcN = sheet.Cells[i, 24].Value.ToString().Replace(",", "");
                        calcN = calcN.PadLeft(17, '0');
                    }

                    if (sheet.Cells[i, 12].Value.ToString() != "")
                    {
                        calcO = sheet.Cells[i, 12].Value.ToString().Replace(",", "");
                        calcO = calcO.PadLeft(17, '0');
                    }

                    if (sheet.Cells[i, 17].Value != null || sheet.Cells[i, 17].Value.ToString() != "")
                    {
                        if (sheet.Cells[i, 17].Value.ToString() == "1")
                        {
                            calcW = "CIF";
                        }
                        else if (sheet.Cells[i, 17].Value.ToString() == "2")
                        {
                            calcW = "FOB";
                        }
                        else
                        {
                            calcW = "   ";
                        }
                    }
                    else
                    {
                        calcW = "   ";
                    }

                    if (sheet.Cells[i, 6].Value.ToString() != "" || sheet.Cells[i, 6].Value != null)
                    {
                        if (sheet.Cells[i, 6].Value.ToString() == "02")
                        {
                            calcY = "S";
                        }
                        else if (sheet.Cells[i, 6].Value.ToString() == "01")
                        {
                            calcY = "N";
                        }
                        else
                        {
                            calcY = " ";
                        }
                    }

                    if (sheet.Cells[i, 13].Value.ToString() != "" || sheet.Cells[i, 13].Value != null)
                    {
                        if (sheet.Cells[i, 13].Value.ToString() == "0")
                        {
                            calcZ = "1";
                        }
                        else if (sheet.Cells[i, 13].Value.ToString() == "1" || sheet.Cells[i, 13].Value.ToString() == "9")
                        {
                            calcZ = "2";
                        }
                        else
                        {
                            calcZ = " ";
                        }
                    }

                    if (sheet.Cells[i, 38].Value.ToString() == "")
                    {
                        calcAC = "  ";
                    }
                    else
                    {
                        calcAC = sheet.Cells[i, 38].Value.ToString().PadRight(2, ' ');
                    }

                    if (sheet.Cells[i, 39].Value.ToString() == "")
                    {
                        calcAD = "     ";
                    }
                    else
                    {
                        calcAD = sheet.Cells[i, 39].Value.ToString().PadRight(5, ' ');
                    }

                    if (sheet.Cells[i, 40].Value.ToString() == "")
                    {
                        calcAE = "000000000";
                    }
                    else
                    {
                        calcAE = sheet.Cells[i, 40].Value.ToString().PadLeft(9, '0');
                    }

                    if (sheet.Cells[i, 41].Value.ToString() == "")
                    {
                        calcAF = "00000000";
                    }
                    else
                    {
                        calcAF = sheet.Cells[i, 41].Value.ToString().PadLeft(8, '0');
                    }

                    if (sheet.Cells[i, 42].Value.ToString() == "")
                    {
                        calcAG = "              ";
                    }
                    else
                    {
                        calcAG = sheet.Cells[i, 42].Value.ToString().PadRight(14, ' ');
                    }


                    calcP = "              ";
                    calcQ = "               ";
                    calcR = "              ";
                    calcS = "                 ";
                    calcT = "          ";
                    calcU = "                 ";
                    calcV = "                 ";
                    calcX = "               ";
                    calcAA = "                                             ";
                    calcAB = "                                                  ";

                    final = calcA + calcB + calcC + calcD + calcE + calcF + calcG + calcH + calcI + calcJ + calcK + calcL + calcM + calcN + calcO + calcP + calcQ + calcR + calcS + calcT + calcU + calcV + calcW + calcY + calcZ + calcAA + calcAB + calcAC + calcAD + calcAE + calcAF + calcAG;
                    /* final = new string((from c in final
                                         where char.IsWhiteSpace(c) || char.IsLetterOrDigit(c)
                                         select c
                    ).ToArray());*/

                    x.WriteLine(final);
                    x.Flush();
                    a2++;
                    //}
                }

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
                    x.Flush();

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
                        if (sheet.Cells[i, 8].Value != null && !string.IsNullOrEmpty(sheet.Cells[i, 8].Value.ToString()))
                            calcA = sheet.Cells[i, 5].Value + "" + sheet.Cells[i, 4].Value + "" + sheet.Cells[i, 8].Value + "" + sheet.Cells[i, 10].Value.ToString().PadLeft(9, '0') + "" + sheet.Cells[i, 12].Value + "" + sheet.Cells[i, 6].Value;
                        else
                            calcA = sheet.Cells[i, 5].Value + "" + sheet.Cells[i, 4].Value + "   " + sheet.Cells[i, 10].Value.ToString().PadLeft(9, '0') + "" + sheet.Cells[i, 12].Value + "" + sheet.Cells[i, 6].Value;
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
                x.Flush();

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

                            if (sheet.Cells[i, 7].Value.ToString() != "")
                            {
                                calcA = sheet.Cells[i, 3].Value.ToString() + sheet.Cells[i, 2].Value.ToString() +
                                    sheet.Cells[i, 5].Value.ToString() + sheet.Cells[i, 8].Value.ToString().PadLeft(9, '0') +
                                    sheet.Cells[i, 7].Value.ToString().PadLeft(3, '0') + sheet.Cells[i, 10].Value.ToString() +
                                    sheet.Cells[i, 11].Value.ToString() + sheet.Cells[i, 4].Value.ToString();
                            }
                            else
                            {
                                calcA = sheet.Cells[i, 3].Value.ToString() + sheet.Cells[i, 2].Value.ToString() +
                                    sheet.Cells[i, 5].Value.ToString() + sheet.Cells[i, 8].Value.ToString().PadLeft(9, '0') +
                                    "   " + sheet.Cells[i, 10].Value.ToString() +
                                    sheet.Cells[i, 11].Value.ToString() + sheet.Cells[i, 4].Value.ToString();

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

                        if (!string.IsNullOrEmpty(sheet.Cells[i - count, 7].Value.ToString()))
                        {
                            calcB = sheet.Cells[i - count, 3].Value.ToString() + sheet.Cells[i - count, 2].Value.ToString() +
                                sheet.Cells[i - count, 5].Value.ToString() + sheet.Cells[i - count, 8].Value.ToString().PadLeft(9, '0') +
                                sheet.Cells[i - count, 7].Value.ToString().PadLeft(3, '0') + sheet.Cells[i - count, 10].Value.ToString() +
                                sheet.Cells[i - count, 11].Value.ToString() + sheet.Cells[i - count, 4].Value.ToString();
                        }
                        else
                        {
                            calcB = sheet.Cells[i - count, 3].Value.ToString() + sheet.Cells[i - count, 2].Value.ToString() +
                                sheet.Cells[i - count, 5].Value.ToString() + sheet.Cells[i - count, 8].Value.ToString().PadLeft(9, '0') +
                                "   " + sheet.Cells[i - count, 10].Value.ToString() +
                                sheet.Cells[i - count, 11].Value.ToString() + sheet.Cells[i - count, 4].Value.ToString();
                        }

                    }
                    calcTempC.Add(calcB);
                    x.WriteLine("|" + calcA + "|" + calcB + y);
                    x.Flush();
                }
                else
                {
                    count += 1;
                }

                i++;
            }
        }

        private void GeraC170NrDocumento()
        {
            foreach (string y in C100)
            {
                if (y.Contains("C170"))
                {
                    string[] value = y.Split('|');

                    //calcTempB[0].ToString();

                    // if()
                }
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

        public void RemoveDuplicada()
        {
            var temp = new List<string>();

            if (!C100.Any())
            {
                temp = C100.Distinct().ToList();

                C100.Clear();

                foreach (var y in temp)
                {
                    C100.Add(y);
                }

                temp.Clear();
            }

            if (!A100.Any())
            {
                temp = A100.Distinct().ToList();

                A100.Clear();

                foreach (var y in temp)
                {
                    A100.Add(y);
                }

                temp.Clear();
            }

            if (!C113.Any())
            {
                temp = C113.Distinct().ToList();

                C113.Clear();

                foreach (var y in temp)
                {
                    C113.Add(y);
                }

                temp.Clear();
            }

            if (!R200.Any())
            {
                temp = R200.Distinct().ToList();

                R200.Clear();

                foreach (var y in temp)
                {
                    R200.Add(y);
                }

                temp.Clear();
            }

            if (!R150.Any())
            {
                temp = R150.Distinct().ToList();

                R150.Clear();

                foreach (var y in temp)
                {
                    R150.Add(y);
                }

                temp.Clear();
            }

            if (!R1100.Any())
            {
                temp = R1100.Distinct().ToList();

                R1100.Clear();

                foreach (var y in temp)
                {
                    R1100.Add(y);
                }

                temp.Clear();
            }
        }

    }
}
