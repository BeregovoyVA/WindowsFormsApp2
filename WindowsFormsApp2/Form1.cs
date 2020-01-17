using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();            
        }

        string TNCascade;


        private void button1_Click(object sender, EventArgs e)
        {

            //Открываем файл Экселя
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //Создаём приложение.
                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //Открываем книгу.                                                                                                                                                        
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

                //Очищаем от старого текста окно вывода.
                richTextBox1.Clear();

                for (int i = 1; i < 101; i++)
                {
                    //Выбираем область таблицы. (в нашем случае просто ячейку)
                    Microsoft.Office.Interop.Excel.Range range = ObjWorkSheet.get_Range(textBox1.Text + i.ToString(), textBox1.Text + i.ToString());
                    //Добавляем полученный из ячейки текст.
                    richTextBox1.Text = richTextBox1.Text + range.Text.ToString() + "\n";
                    //это чтобы форма прорисовывалась (не подвисала)...
                    Application.DoEvents();
                }

                //Удаляем приложение (выходим из экселя) - ато будет висеть в процессах!
                ObjExcel.Quit();
            }
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            ////Band $LTE $GSM $WCDMA
            //richTextBox1.Clear();
            //richTextBox1.Text = richTextBox1.Text + "//Band" + "\n";
            ////LTE
            //int LTE = 0;
            //string[] Band_LTE = new string[] { "L18", "DL18", "GL18", "L26" };
            //foreach (string element in Band_LTE)
            //{
            //    LTE = textBox4.Text.IndexOf(element);
            //    if (LTE > 0) break;
            //}
            //if (LTE == -1)
            //{ richTextBox1.Text = richTextBox1.Text + "$LTE := 0" + "\n"; }
            //else
            //{ richTextBox1.Text = richTextBox1.Text + "$LTE := 1" + "\n"; }
            
            ////GSM
            //int GSM = 0;
            //string[] Band_GSM = new string[] { "G9", "G(", "G18", "D18", "DL18", "GL18" };
            //foreach (string element in Band_GSM)
            //{
            //    GSM = textBox4.Text.IndexOf(element);
            //    if (GSM > 0) break;
            //}
            //if (GSM == -1)
            //{ richTextBox1.Text = richTextBox1.Text + "$GSM := 0" + "\n"; }
            //else
            //{ richTextBox1.Text = richTextBox1.Text + "$GSM := 1" + "\n"; }

            ////WCDMA
            //int WCDMA = 0;
            //string[] Band_WCDMA = new string[] { "W(", "W21", "W2100"};
            //foreach (string element in Band_WCDMA)
            //{
            //    WCDMA = textBox4.Text.IndexOf(element);
            //    if (WCDMA > 0) break;
            //}
            //if (WCDMA == -1)
            //{ richTextBox1.Text = richTextBox1.Text + "$WCDMA := 0" + "\n"; }
            //else
            //{ richTextBox1.Text = richTextBox1.Text + "$WCDMA := 1" + "\n"; }

            ////HW R503
            //int R503 = 0;
            //string[] HW_R503 = new string[] { "R503"};
            //foreach (string element in HW_R503)
            //{
            //    R503 = textBox4.Text.IndexOf(element);
            //    if (R503 > 0) break;
            //}
            //if (R503 == -1)
            //{ richTextBox1.Text = richTextBox1.Text + "$R503 := 0" + "\n"; }
            //else
            //{ richTextBox1.Text = richTextBox1.Text + "$R503 := 1" + "\n"; }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "x:\\3G_LTE\\Center";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = openFileDialog1.FileName;
                richTextBox2.LoadFile(openFileDialog1.FileName, RichTextBoxStreamType.PlainText);

                //string[] template = File.ReadAllLines(openFileDialog1.FileName);        //считываем в массив текст из файла шаблона
                //string[,] template_work; // = new string[1000, 2];
                //for (int i = 0; i<template.Length; i++)
                //{
                //    template_work[i, 0] = template[i];
                //    template_work[i, 1] = "0";
                //}

                //for (int i = 0; i < template.Length; i++)
                //{
                //    richTextBox2.AppendText(template_work[i, 0]);
                //}



                //richTextBox2.LoadFile(openFileDialog1.FileName, RichTextBoxStreamType.PlainText);
                //richTextBox2.SelectAll();
                //richTextBox2.SelectionColor = Color.Red;
                //richTextBox1.Lines = template_result;
                //richTextBox2.Lines = template_work;
            }

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
     
       
        







     //Tracker
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            //richTextBox3.Clear();
            string text = textBox3.Text;
            if (text == "") return;
            string[] words = text.Split(new char[] { '\t' });
            if ((words.Length < 14) || (textBox3.Text==""))
            {
                MessageBox.Show("Не достаточно исходных данных. Нужно 14столбцов из Tracker (Region<---->Reh3G)!");
                textBox3.Clear();
                return;
            }
            //Site, HW, Region
            richTextBox4.Text = richTextBox4.Text + "#Site := " + words[1] + "\n";
            string RBSNAME = words[1].Substring(words[1].Length - 4);       //RBSNAME= 4 последние цыфры

            richTextBox4.Text = richTextBox4.Text + "#Region := " + words[0] + "\n";
            saveFileDialog1.FileName = words[1]+"_"+"Before_BB_";  //44665_Before_BB_

            //Maping Region<-->Region number
            string[] Region_array = new string[] { "KEM",   "KRN", "ORB", "SAR", "PNZ", "ULN", "YOL","UFA", "BAR", "OMS", "KZN", "EKB", "TOM", "NVK", "CHB", "KRG", "SRN", "NSK", "ABN" };
            string[] Regnum_array = new string[] { "42",    "24", "56", "64", "58", "73", "12", "2", "22", "55", "16", "66", "70", "42", "21", "45", "13", "54", "19" };
            if (words[0] != "")
            {
                int count = 0;
                foreach (string element in Region_array)
                {
                    if (words[0].Substring(4) == element)
                    {
                        RBSNAME = Regnum_array[count] + RBSNAME;
                        break;
                    }
                    else { count++; }
                }
            }
            
            
            richTextBox4.Text = richTextBox4.Text + "#HW := " + words[2] + "\n";
            richTextBox4.Text = richTextBox4.Text + "#Band := " + words[3] + "\n";
            richTextBox4.Text = richTextBox4.Text + "#MixedModeRadio := " + words[4] + "\n";
            richTextBox4.Text = richTextBox4.Text + "#ARS := " + words[5] + "\n";
            richTextBox4.Text = richTextBox4.Text + "#SAU := " + words[6] + "\n";
            richTextBox4.Text = richTextBox4.Text + "#Climate := " + words[7] + "\n";
            richTextBox4.Text = richTextBox4.Text + "#HubPosition := " + words[8] + "\n";
            richTextBox4.Text = richTextBox4.Text + "#TNCascade := " + words[10] + "\n";
            TNCascade = words[10];
            richTextBox4.Text = richTextBox4.Text + "#KGET := " + words[11] + "\n";
            richTextBox4.Text = richTextBox4.Text + "#Reh2G := " + words[12] + "\n";

            //ПЕРЕМЕННЫЕ
            //$supportSystemControl            
            if ((words[7] == "MASTER")||(words[7] == "")) { richTextBox3.Text = richTextBox3.Text + "$supportSystemControl := 1" + "\n"; }
            if (words[7] == "SLAVE") { richTextBox3.Text = richTextBox3.Text + "$supportSystemControl := 0" + "\n"; }
            //$SAU
            if (words[6] == "YES") { richTextBox3.Text = richTextBox3.Text + "$SAU := 1" + "\n"; }
            else { richTextBox3.Text = richTextBox3.Text + "$SAU := 0" + "\n"; }
            //$ExternalNode
            if ((words[7] == "MASTER")||(words[7] == "SLAVE")) { richTextBox3.Text = richTextBox3.Text + "$ExternalNode := 1" + "\n"; }
            else { richTextBox3.Text = richTextBox3.Text + "$ExternalNode := 0" + "\n"; }
            //$Bridge
            if (words[10] != "")
            {
                richTextBox3.Text = richTextBox3.Text + "$Bridge := 1" + "\n";
                //$TNPORT
                richTextBox3.Text = richTextBox3.Text + "$TNPORT := "+ words[10].Substring(0, 4) + "\n";
            }
            else { richTextBox3.Text = richTextBox3.Text + "$Bridge := 0" + "\n"; }
            //$kgetpath
            if (words[11] != "") richTextBox3.Text = richTextBox3.Text + "$kgetpath := "+ words[11] + "\n";
            //$BSCNAME
            if (words[12] != "") richTextBox3.Text = richTextBox3.Text + "$BSCNAME := " + words[12].Substring(words[12].Length-7) + "\n";
            //$SHARED
            if (words[5] == "YES") { richTextBox3.Text = richTextBox3.Text + "$SHARED := 1" + "\n"; }
            else { richTextBox3.Text = richTextBox3.Text + "$SHARED := 0" + "\n"; }



            if (words[12].Length >= 7) {richTextBox4.Text = richTextBox4.Text + "#Rehoming2G := " + words[12].Substring(words[12].Length - 7) +"\n";}
            else {richTextBox3.Text = richTextBox3.Text + "#Rehoming2G := " + words[12] +"\n";}

            richTextBox4.Text = richTextBox4.Text + "#Rehoming3G := " + words[13] + "\n";
            //Band $LTE $GSM $WCDMA
            //richTextBox3.Text = richTextBox3.Text + "//Band" + "\n";
            
            //LTE
            int LTE_18 = 0;
            int LTE_26 = 0;
            string[] Band_LTE_18 = new string[] { "L18(", "DL18(", "GL18(" };
            foreach (string element in Band_LTE_18)
            {
                LTE_18 = textBox3.Text.IndexOf(element);
                if (LTE_18 > 0) break;
            }
            string[] Band_LTE_26 = new string[] { "L26(" };
            foreach (string element in Band_LTE_26)
            {
                LTE_26 = textBox3.Text.IndexOf(element);
                if (LTE_26 > 0) break;
            }
            if ((LTE_18 == -1) & (LTE_26 == -1))
            { richTextBox3.Text = richTextBox3.Text + "$LTE := 0" + "\n"; }
            else
            {
                richTextBox3.Text = richTextBox3.Text + "$LTE := 1" + "\n";
                RBSNAME = "L" + RBSNAME;
                saveFileDialog1.FileName = saveFileDialog1.FileName + "L";  //44665_Before_BB_L
                //Ищем порты
                string port_LTE_26 = "";
                string port_LTE_18 = "";
                if (LTE_18 != -1)
                {
                    port_LTE_18 = textBox3.Text.Substring(LTE_18);
                    port_LTE_18 = port_LTE_18.Substring(port_LTE_18.IndexOf("(") + 1, port_LTE_18.IndexOf(")") - port_LTE_18.IndexOf("(") - 1);
                }
                if (LTE_26 != -1)
                {
                    port_LTE_26 = textBox3.Text.Substring(LTE_26);
                    port_LTE_26 = port_LTE_26.Substring(port_LTE_26.IndexOf("(") + 1, port_LTE_26.IndexOf(")") - port_LTE_26.IndexOf("(") - 1);
                }

                if ((LTE_26 != -1) & (LTE_18 == -1)) //есть только L26
                {
                    richTextBox3.Text = richTextBox3.Text + "$numbandsLTE := 1" + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$LTEband_b1 := 2600" + "\n" + "$LTEband_b2 := 1800" + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$numsectorsLTE_b1 := " + port_LTE_26.Length.ToString() + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$numsectorsLTE_b2 := " + "\n";
                    for (int i = 1; i <= port_LTE_26.Length; i++)
                    {                        
                        richTextBox3.Text = richTextBox3.Text + "$RiPortLTE[" + i.ToString() + "] := " + port_LTE_26.Substring(i - 1, 1) + "\n";
                    }
                }
                if ((LTE_26 == -1) & (LTE_18 != -1)) //есть только L18
                {
                    richTextBox3.Text = richTextBox3.Text + "$numbandsLTE := 1" + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$LTEband_b1 := 1800" + "\n" + "$LTEband_b2 := 2600" + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$numsectorsLTE_b1 := " + port_LTE_18.Length.ToString() + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$numsectorsLTE_b2 := " + "\n";
                    for (int i = 1; i <= port_LTE_18.Length; i++)
                    {
                        richTextBox3.Text = richTextBox3.Text + "$RiPortLTE[" + i.ToString() + "] := " + port_LTE_18.Substring(i - 1, 1) + "\n";
                    }
                }
                if ((LTE_26 != -1) & (LTE_18 != -1)) //есть L26+L18
                {
                    string port_LTE26_LTE18 = port_LTE_26 + port_LTE_18;
                    richTextBox3.Text = richTextBox3.Text + "$numbandsLTE := 2" + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$LTEband_b1 := 2600" + "\n" + "$LTEband_b2 := 1800" + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$numsectorsLTE_b1 := " + port_LTE_26.Length.ToString() + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$numsectorsLTE_b2 := " + port_LTE_18.Length.ToString() + "\n";
                    for (int i = 1; i <= port_LTE26_LTE18.Length; i++)
                    {
                        richTextBox3.Text = richTextBox3.Text + "$RiPortLTE[" + i.ToString() + "] := " + port_LTE26_LTE18.Substring(i - 1, 1) + "\n";
                    }                    
                }                
                //$mixedmode
                if (words[4] == "YES") { richTextBox3.Text = richTextBox3.Text + "$mixedmode := 1" + "\n"; }
                else { richTextBox3.Text = richTextBox3.Text + "$mixedmode := 0" + "\n"; }
            }
            
            //WCDMA
            int WCDMA = 0;
            string[] Band_WCDMA = new string[] { "W(", "W21", "W2100" };
            foreach (string element in Band_WCDMA)
            {
                WCDMA = textBox3.Text.IndexOf(element);
                if (WCDMA > 0) break;
            }
            if (WCDMA == -1)
            { richTextBox3.Text = richTextBox3.Text + "$WCDMA := 0" + "\n"; }
            else
            {
                richTextBox3.Text = richTextBox3.Text + "$WCDMA := 1" + "\n";
                richTextBox3.Text = richTextBox3.Text + "$RBSID := " + words[1] + "\n";
                RBSNAME = "W" + RBSNAME;
                saveFileDialog1.FileName = saveFileDialog1.FileName + "W";
                //Ищем порты
                string port_WCDMA = textBox3.Text.Substring(WCDMA);
                port_WCDMA = port_WCDMA.Substring(port_WCDMA.IndexOf("(") + 1, port_WCDMA.IndexOf(")") - port_WCDMA.IndexOf("(")-1);
                richTextBox3.Text = richTextBox3.Text + "$numsectorsWCDMA := " + port_WCDMA.Length.ToString() + "\n";
                for (int i = 1; i<= port_WCDMA.Length; i++)
                {
                    //$RiPortWCDMA
                    richTextBox3.Text = richTextBox3.Text + "$RiPortWCDMA[" + i.ToString() + "] := "+ port_WCDMA.Substring(i-1,1)+ "\n";
                }


            }

            //GSM
            int GSM_1800 = 0;
            int GSM_900 = 0;
            string[] Band_GSM_900 = new string[] { "G9(", "G(" };
            string[] Band_GSM_1800 = new string[] { "G18(", "D18(", "DL18(", "GL18(" };
            foreach (string element in Band_GSM_1800)
            {
                GSM_1800 = textBox3.Text.IndexOf(element);
                if (GSM_1800 > 0) break;
            }
            foreach (string element in Band_GSM_900)
            {
                GSM_900 = textBox3.Text.IndexOf(element);
                if (GSM_900 > 0) break;
            }
            if ((GSM_1800 == -1)& (GSM_900 == -1))
            { richTextBox3.Text = richTextBox3.Text + "$GSM := 0" + "\n"; }
            else
            {
                richTextBox3.Text = richTextBox3.Text + "$GSM := 1" + "\n";
                RBSNAME = "G" + RBSNAME;
                saveFileDialog1.FileName = saveFileDialog1.FileName + "G";
                //Ищем порты
                string port_GSM_1800 = "";
                string port_GSM_900 = "";
                if (GSM_1800 != -1)
                {
                    port_GSM_1800 = textBox3.Text.Substring(GSM_1800);
                    port_GSM_1800 = port_GSM_1800.Substring(port_GSM_1800.IndexOf("(") + 1, port_GSM_1800.IndexOf(")") - port_GSM_1800.IndexOf("(") - 1);
                }
                if (GSM_900 != -1)
                {
                    port_GSM_900 = textBox3.Text.Substring(GSM_900);
                    port_GSM_900 = port_GSM_900.Substring(port_GSM_900.IndexOf("(") + 1, port_GSM_900.IndexOf(")") - port_GSM_900.IndexOf("(") - 1);
                }

                if ((GSM_1800 != -1) & (GSM_900 == -1)) //есть только GSM_1800
                {
                    richTextBox3.Text = richTextBox3.Text + "$numsectorsGSM_b1 := " + port_GSM_1800.Length.ToString() + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$numbandsGSM := 1"+ "\n";
                    richTextBox3.Text = richTextBox3.Text + "$GSMband_b1 := 1800" + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$GSMband_b2 := 900" + "\n";
                    for (int i = 1; i <= port_GSM_1800.Length; i++) richTextBox3.Text = richTextBox3.Text + "$GsmSector_b1[" + i.ToString() + "] := " + words[1] + i.ToString() + "\n";
                    if (LTE_18 == -1)   //если нет L18, то ищем порты. Если есть, порты искать не нужно, используются порты L18
                    {
                        for (int i = 1; i <= port_GSM_1800.Length; i++)
                        {
                            richTextBox3.Text = richTextBox3.Text + "$RiPortGSM[" + i.ToString() + "] := " + port_GSM_1800.Substring(i - 1, 1) + "\n";
                        }
                    }
                }

                if ((GSM_1800 == -1) & (GSM_900 != -1)) //есть только GSM_900
                {
                    richTextBox3.Text = richTextBox3.Text + "$numsectorsGSM_b1 := " + port_GSM_900.Length.ToString() + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$numbandsGSM := 1" + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$GSMband_b1 := 900" + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$GSMband_b2 := 1800" + "\n";
                    for (int i = 1; i <= port_GSM_900.Length; i++) richTextBox3.Text = richTextBox3.Text + "$GsmSector_b1[" + i.ToString() + "] := " + words[1] + (i + 4).ToString() + "\n";
                    for (int i = 1; i <= port_GSM_900.Length; i++)
                    {
                        richTextBox3.Text = richTextBox3.Text + "$RiPortGSM[" + i.ToString() + "] := " + port_GSM_900.Substring(i - 1, 1) + "\n";
                    }                    
                }

                if ((GSM_1800 != -1) & (GSM_900 != -1)) //есть GSM_1800 + GSM_900
                {
                    richTextBox3.Text = richTextBox3.Text + "$numsectorsGSM_b1 := " + port_GSM_1800.Length.ToString() + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$numsectorsGSM_b2 := " + port_GSM_900.Length.ToString() + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$numbandsGSM := 2" + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$GSMband_b1 := 1800" + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$GSMband_b2 := 900" + "\n";
                    for (int i = 1; i <= port_GSM_1800.Length; i++) richTextBox3.Text = richTextBox3.Text + "$GsmSector_b1[" + i.ToString() + "] := " + words[1] + i.ToString() + "\n";
                    for (int i = 1; i <= port_GSM_900.Length; i++) richTextBox3.Text = richTextBox3.Text + "$GsmSector_b2[" + i.ToString() + "] := " + words[1] + (i+4).ToString() + "\n";
                    if (LTE_18 == -1)   //если нет L18, то ищем все порты. Если есть, то ищем порты GSM_900, для GSM_1800 используются порты L18
                    {
                        string port_GSM = port_GSM_1800 + port_GSM_900;
                        for (int i = 1; i <= port_GSM.Length; i++)
                        {
                            richTextBox3.Text = richTextBox3.Text + "$RiPortGSM[" + i.ToString() + "] := " + port_GSM.Substring(i - 1, 1) + "\n";
                        }
                    }
                    else   //есть L18 диапазон
                    {
                        for (int i = 1; i <= port_GSM_900.Length; i++)
                        {
                            richTextBox3.Text = richTextBox3.Text + "$RiPortGSM[" + i.ToString() + "] := " + port_GSM_900.Substring(i - 1, 1) + "\n";
                        }
                    }
                }
                //$mixedmodeGSM
                if (words[4] == "YES")
                {
                    if ((textBox3.Text.IndexOf("G18") != -1) || (textBox3.Text.IndexOf("D18") != -1) || (textBox3.Text.IndexOf("DL18") != -1) || (textBox3.Text.IndexOf("GL18") != -1)) { richTextBox3.Text = richTextBox3.Text + "$mixedmodeGSM := 1" + "\n"; }
                    else { richTextBox3.Text = richTextBox3.Text + "$mixedmodeGSM := 0" + "\n"; }
                }
            }


            //$RBSNAME    
            richTextBox3.Text = richTextBox3.Text + "$RBSNAME := "+ RBSNAME + "\n";            

            //HW R503
            int R503 = 0;
            string[] HW_R503 = new string[] { "R503" };
            foreach (string element in HW_R503)
            {
                R503 = textBox3.Text.IndexOf(element);
                if (R503 > 0) break;
            }
            if (R503 == -1)
            { richTextBox3.Text = richTextBox3.Text + "$R503 := 0" + "\n"; }
            else
            { richTextBox3.Text = richTextBox3.Text + "$R503 := 1" + "\n"; }


            //GPS пока всегда 0
            richTextBox3.Text = richTextBox3.Text + "$GPS := 0" + "\n";















        }
        //IP plan
        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            string text = textBox5.Text;
            string[] str = text.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries); //str[0] - строка заголовока, str[1] - строка данных
            if (str.Length != 2)
            {
                MessageBox.Show("Нужно 2 строки (заголовок и данные)!");
                textBox5.Clear();
                return;
            }
            //Анализ
            string[] str_0 = str[0].Split(new char[] { '\t' }); //str_0[] - поля заголовока
            string[] str_1 = str[1].Split(new char[] { '\t' }); //str_1[] - поля данных
            for (int i = 0; i < str_0.Length; i++)
            {
                //ABIS
                if (str_0[i] == "ABIS")
                {
                    //провека на корректность ip (содержит 4 точки)                    
                    int count = str_1[i].ToCharArray().Where(j => j == '.').Count();
                    if (count != 3) continue;
                    string abis_mask = str_1[i + 1].Replace(" ", "");
                    switch (abis_mask)
                    {
                        case "255.255.255.224":
                            abis_mask = "/27";
                            break;
                        case "255.255.255.240":
                            abis_mask = "/28";
                            break;
                        case "255.255.255.252":
                            abis_mask = "/30";
                            break;
                        default:
                            abis_mask = "/???????????";
                            break;
                    }
                    richTextBox3.Text = richTextBox3.Text + "$GSM_IP := " + str_1[i].Replace(" ", "") + abis_mask + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$GSM_DG := " + str_1[i + 2].Replace(" ", "") + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$GSM_Vlan := " + str_1[i + 3].Replace(" ", "") + "\n";
                    //Bridge                    
                    if (richTextBox3.Text.IndexOf("$Bridge := 1") != -1)
                    {
                        if (TNCascade.Substring(TNCascade.Length - 2) == "2G")
                        {
                            richTextBox3.Text = richTextBox3.Text + "$OAMVLAN := " + str_1[i+7].Replace(" ", "") + "\n";
                            richTextBox3.Text = richTextBox3.Text + "$TRAFFICVLAN := " + str_1[i+3].Replace(" ", "") + "\n";
                        }
                    }

                }
                //WCDMA
                if (str_0[i] == "3G DATA")
                {
                    int count = str_1[i].ToCharArray().Where(j => j == '.').Count();
                    if (count != 3) continue;
                    string WCDMA_mask = str_1[i + 1].Replace(" ", "");
                    switch (WCDMA_mask)
                    {
                        case "255.255.255.224":
                            WCDMA_mask = "/27";
                            break;
                        case "255.255.255.252":
                            WCDMA_mask = "/30";
                            break;
                        default:
                            WCDMA_mask = "/???????????";
                            break;
                    }
                    richTextBox3.Text = richTextBox3.Text + "$WCDMA_IP := " + str_1[i].Replace(" ", "") + WCDMA_mask + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$WCDMA_DG := " + str_1[i + 2].Replace(" ", "") + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$WCDMA_Vlan := " + str_1[i + 3].Replace(" ", "") + "\n";
                    //Bridge                    
                    if (richTextBox3.Text.IndexOf("$Bridge := 1") != -1)
                    {
                        if (TNCascade.Substring(TNCascade.Length - 2) == "3G")
                        {
                            richTextBox3.Text = richTextBox3.Text + "$OAMVLAN := " + str_1[i + 7].Replace(" ", "") + "\n";
                            richTextBox3.Text = richTextBox3.Text + "$TRAFFICVLAN := " + str_1[i + 3].Replace(" ", "") + "\n";
                        }
                    }

                }
                //LTE
                if (str_0[i] == "4G DATA")
                {
                    int count = str_1[i].ToCharArray().Where(j => j == '.').Count();
                    if (count != 3) continue;
                    string LTE_mask = str_1[i + 1].Replace(" ", "");
                    switch (LTE_mask)
                    {
                        case "255.255.255.224":
                            LTE_mask = "/27";
                            break;
                        case "255.255.255.252":
                            LTE_mask = "/30";
                            break;
                        default:
                            LTE_mask = "/???????????";
                            break;
                    }
                    richTextBox3.Text = richTextBox3.Text + "$LTE_IP := " + str_1[i].Replace(" ", "") + LTE_mask + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$LTE_DG := " + str_1[i + 2].Replace(" ", "") + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$LTE_Vlan := " + str_1[i + 3].Replace(" ", "") + "\n";
                    //Bridge                    
                    if (richTextBox3.Text.IndexOf("$Bridge := 1") != -1)
                    {
                        if (TNCascade.Substring(TNCascade.Length - 2) == "4G")
                        {
                            richTextBox3.Text = richTextBox3.Text + "$OAMVLAN := " + str_1[i + 7].Replace(" ", "") + "\n";
                            richTextBox3.Text = richTextBox3.Text + "$TRAFFICVLAN := " + str_1[i + 3].Replace(" ", "") + "\n";
                        }
                    }
                }
                //ARS
                if (str_0[i] == "IP S1(МТС)")
                {
                    int count = str_1[i].ToCharArray().Where(j => j == '.').Count();
                    if (count != 3) continue;
                    string ARS_mask = str_1[i + 1].Replace(" ", "");
                    switch (ARS_mask)
                    {
                        case "255.255.255.224":
                            ARS_mask = "/27";
                            break;
                        case "255.255.255.252":
                            ARS_mask = "/30";
                            break;
                        default:
                            ARS_mask = "/???????????";
                            break;
                    }
                    richTextBox3.Text = richTextBox3.Text + "$LTE_Vlan_SHARED := " + str_1[i + 3].Replace(" ", "") + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$TNPORT_SHARED := " + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$LTE_IP_SHARED_OUTER := " + str_1[i].Replace(" ", "") + ARS_mask + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$LTE_DG_SHARED_OUTER := " + str_1[i + 2].Replace(" ", "") + "\n";
                    richTextBox3.Text = richTextBox3.Text + "$LTE_IP_SHARED_INNER := " + str_1[i + 4].Replace(" ", "") + "/32\n";
                    richTextBox3.Text = richTextBox3.Text + "$LTE_IP_SHARED_SECGW := " + str_1[i + 5].Replace(" ", "") + "\n";
                }                             

            }

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox3_TextChanged(object sender, EventArgs e)
        {
            string template = richTextBox2.Text;
            string[] str_template = template.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            string variable = richTextBox3.Text;
            string[] str_variable = variable.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string element in str_variable)
            {
                //проверка строки на комманду (наяинается с $)
                if (element != "")
                {
                    if (element.Substring(0, 1) == "$") //значит это переменная
                    {
                        string var_serch = "";
                        var_serch = element.Substring(0, element.IndexOf(' '));
                        int count = 0;
                        foreach (string element2 in str_template)
                        {
                            if (element2.IndexOf(var_serch) != -1)
                            {
                                str_template[count] = element;
                                break;
                            }
                            count++;
                        }
                    }
                    else { continue; }

                }
                else { continue; }
            }
            richTextBox2.Clear();
            foreach (string element3 in str_template)
            {
                richTextBox2.Text += element3 + "\r\n";
            }
        }

    
        
        
        
     //Refresh
        private void button5_Click(object sender, EventArgs e)
        {
            string template = richTextBox2.Text;
            string[] str_template = template.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            string variable = richTextBox3.Text;
            string[] str_variable = variable.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string element in str_variable)
            {
                //проверка строки на комманду (наяинается с $)
                if (element != "")
                {
                    if (element.Substring(0, 1) == "$") //значит это переменная
                    {
                        string var_serch = "";
                        var_serch = element.Substring(0, element.IndexOf(' '));
                        int count = 0;
                        foreach (string element2 in str_template)
                        {
                            if (element2.IndexOf(var_serch) != -1)
                            {
                                str_template[count] = element;
                                break;
                            }
                            count++;
                        }
                    }
                    else { continue; }

                }
                else { continue; }
            }
            richTextBox2.Clear();
            foreach (string element3 in str_template)
            {
                richTextBox2.Text += element3 + "\r\n";
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //saveFileDialog1.FileName = "";  //words[0] EAS-BAR
            saveFileDialog1.InitialDirectory = "C:\\Users\\eerlbav\\OneDrive - Ericsson AB\\Desktop\\Integration";
            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            // получаем выбранный файл

            string filename = saveFileDialog1.FileName;
            // сохраняем текст в файл
            System.IO.File.WriteAllText(filename, richTextBox2.Text);
            MessageBox.Show("Файл сохранен");
        }

        private void button7_Click(object sender, EventArgs e)
        {

        }

        private void button7_reload_template(object sender, EventArgs e)
        {
            if (textBox2.Text != "")
            {
                richTextBox2.Text = "";
                textBox2.Text = openFileDialog1.FileName;
                richTextBox2.LoadFile(openFileDialog1.FileName, RichTextBoxStreamType.PlainText);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            textBox3.Text = "";
            richTextBox3.Text = "";

        }

        private void button9_Click(object sender, EventArgs e)
        {
            //Form1.textBox5_TextChanged();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            richTextBox3.Text = "";
            richTextBox4.Text = "";
        }






        //CDD!!!!!!!!!!!!!!!!!!!!!!!!
        private void textBox6_CDD_TextChanged(object sender, EventArgs e)
        {
            string text = textBox6.Text;
            string[] str = text.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries); //str[0] - строка заголовока, str[1...N] - строки данных
            string[] tmp = str[0].Split(new string[] { "\t" }, StringSplitOptions.None); //str[0] - строка заголовока, str[1...N] - строки данных
            string[,] cdd = new string[str.GetLength(0), tmp.GetLength(0)];

            for (int i = 0; i < str.GetLength(0); i++) //кол-во строк
            {
                tmp = str[i].Split(new string[] { "\t" }, StringSplitOptions.None);
                for (int j = 0; j < tmp.GetLength(0); j++) //кол-во столбцов
                {
                    cdd[i, j] = tmp[j];     //все cdd в двумерном массиве cdd[i,j]
                }
            }

            //анализ cdd, определяем диапазоны по cid
            //L18 - 7,8,9; L26 - 4,5,6; ARSL18 - 17,18,19; ARSL26 - 14,15,16            
            Boolean L26 = false; Boolean L18 = false; Boolean ARS_L26 = false; Boolean ARS_L18 = false;
            int num_L26_cells = 0; int num_L18_cells = 0; int num_ARS_L26_cells = 0; int num_ARS_L18_cells = 0;
            string earfcndl_L26 = ""; string earfcndl_L18 = ""; string earfcndl_ARS_L26 = ""; string earfcndl_ARS_L18 = "";
            string Bandwidth_L26 = ""; string Bandwidth_L18 = ""; string Bandwidth_ARS_L26 = ""; string Bandwidth_ARS_L18 = "";
            string tac_L26 = ""; string tac_L18 = ""; string tac_ARS_L26 = ""; string tac_ARS_L18 = "";
            string eNodeBID_L26 = ""; string eNodeBID_L18 = ""; string eNodeBID_ARS_L26 = ""; string eNodeBID_ARS_L18 = "";
            string Latitude_L26 = ""; string Latitude_L18 = ""; string Latitude_ARS_L26 = ""; string Latitude_ARS_L18 = "";
            string Longitude_L26 = ""; string Longitude_L18 = ""; string Longitude_ARS_L26 = ""; string Longitude_ARS_L18 = "";
            for (int j = 1; j < cdd.GetLength(1); j++)
            {
                //CellId                
                if (cdd[0,j] == "CellId")
                {
                    for (int i = 1; i < cdd.GetLength(0); i++)
                    {
                        switch (cdd[i, j])
                        {
                            case "4": case "5": case "6":
                                L26 = true;
                                num_L26_cells++;
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "DlEarfcn") { earfcndl_L26 = cdd[i, k].ToString(); } }
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "Bandwidth") { Bandwidth_L26 = cdd[i, k].ToString(); } }
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "TAC") { tac_L26 = cdd[i, k].ToString(); } }
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "eNodeBID") { eNodeBID_L26 = cdd[i, k].ToString(); } }
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "Latitude") { Latitude_L26 = cdd[i, k].ToString(); } }
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "Longitude") { Longitude_L26 = cdd[i, k].ToString(); } }
                                break;
                            case "7": case "8": case "9":
                                L18 = true;
                                num_L18_cells++;
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "DlEarfcn") { earfcndl_L18 = cdd[i, k].ToString(); } }
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "Bandwidth") { Bandwidth_L18 = cdd[i, k].ToString(); } }
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "TAC") { tac_L18 = cdd[i, k].ToString(); } }
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "eNodeBID") { eNodeBID_L18 = cdd[i, k].ToString(); } }
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "Latitude") { Latitude_L18 = cdd[i, k].ToString(); } }
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "Longitude") { Longitude_L18 = cdd[i, k].ToString(); } }
                                break;
                            case "14": case "15": case "16":
                                ARS_L26 = true;
                                num_ARS_L26_cells++;
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "DlEarfcn") { earfcndl_ARS_L26 = cdd[i, k].ToString(); } }
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "Bandwidth") { Bandwidth_ARS_L26 = cdd[i, k].ToString(); } }
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "TAC") { tac_ARS_L26 = cdd[i, k].ToString(); } }
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "eNodeBID") { eNodeBID_ARS_L26 = cdd[i, k].ToString(); } }
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "Latitude") { Latitude_ARS_L26 = cdd[i, k].ToString(); } }
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "Longitude") { Longitude_ARS_L26 = cdd[i, k].ToString(); } }
                                break;
                            case "17": case "18": case "19":
                                ARS_L18 = true;
                                num_ARS_L18_cells++;
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "DlEarfcn") { earfcndl_ARS_L18 = cdd[i, k].ToString(); } }
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "Bandwidth") { Bandwidth_ARS_L18 = cdd[i, k].ToString(); } }
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "TAC") { tac_ARS_L18 = cdd[i, k].ToString(); } }
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "eNodeBID") { eNodeBID_ARS_L18 = cdd[i, k].ToString(); } }
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "Latitude") { Latitude_ARS_L18 = cdd[i, k].ToString(); } }
                                for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "Longitude") { Longitude_ARS_L18 = cdd[i, k].ToString(); } }
                                break;
                            default:
                                
                                break;
                        }
                    }                                                                                                                               
                }
            }

            //только L26
            if (L26 & !L18)
            {
                //richTextBox3.Text = richTextBox3.Text + "$numbandsLTE := 1" + "\n";
                //richTextBox3.Text = richTextBox3.Text + "$LTEband_b1 := 2600" + "\n" + "$LTEband_b2 := 1800" + "\n";
                //richTextBox3.Text = richTextBox3.Text + "$numsectorsLTE_b1 := " + num_L26_cells.ToString() + "\n";
                richTextBox3.Text = richTextBox3.Text + "$earfcndl_b1 := " + earfcndl_L26 + "\n";
                richTextBox3.Text = richTextBox3.Text + "$ChannelBandwidth_b1 := " + Bandwidth_L26 + "\n";
                richTextBox3.Text = richTextBox3.Text + "$tac := " + tac_L26 + "\n";
                richTextBox3.Text = richTextBox3.Text + "$eNBId := " + eNodeBID_L26 + "\n";                
                double tmp_lat = Convert.ToSingle(Latitude_L26);
                tmp_lat = (tmp_lat / 90) * 8388608;
                tmp_lat = Math.Round(tmp_lat,0);
                double tmp_long = Convert.ToSingle(Longitude_L26);
                tmp_long = (tmp_long / 360) * 16777216;
                tmp_long = Math.Round(tmp_long,0);
                richTextBox3.Text = richTextBox3.Text + "$latitude := " + tmp_lat.ToString() + "\n";
                richTextBox3.Text = richTextBox3.Text + "$longitude := " + tmp_long.ToString() + "\n";
                int num_L26 = 0; int num_L18 = 0;
                for (int i = 1; i < cdd.GetLength(0); i++)
                {
                    string CellId = "";
                    for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "CellId") { CellId = cdd[i, k].ToString(); } }
                    if ((CellId == "4") || (CellId == "5") || (CellId == "6")) { num_L26++; }
                    else continue;
                    for (int j = 1; j < cdd.GetLength(1); j++)
                    {
                        if (cdd[0, j] == "CellName") { richTextBox3.Text = richTextBox3.Text + "$CELL_b1["+ num_L26.ToString()+"] := " + cdd[i,j] + "\n"; }
                        if (cdd[0, j] == "CellId") { richTextBox3.Text = richTextBox3.Text + "$cellId_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n"; }
                        if (cdd[0, j] == "Physical Cell ID Group") { richTextBox3.Text = richTextBox3.Text + "$physicalLayerCellIdGroup_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n"; }
                        if (cdd[0, j] == "Physical Layer ID") { richTextBox3.Text = richTextBox3.Text + "$physicalLayerSubCellId_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n"; }
                        if (cdd[0, j] == "RootSequenceIdx") { richTextBox3.Text = richTextBox3.Text + "$rachRootSequence_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n"; }
                        if (cdd[0, j] == "Altitude")
                        {
                            while (cdd[i, j].Length < 4) cdd[i, j] = cdd[i, j] + "0";
                            richTextBox3.Text = richTextBox3.Text + "$altitude_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n";
                        }
                    }
                }


            }
            //только L18
            if (!L26 & L18)
            {
                //richTextBox3.Text = richTextBox3.Text + "$numbandsLTE := 1" + "\n";
                //richTextBox3.Text = richTextBox3.Text + "$LTEband_b1 := 1800" + "\n" + "$LTEband_b2 := 2600" + "\n";
                //richTextBox3.Text = richTextBox3.Text + "$numsectorsLTE_b1 := " + num_L18_cells.ToString() + "\n";
                richTextBox3.Text = richTextBox3.Text + "$earfcndl_b1 := " + earfcndl_L18 + "\n";
                richTextBox3.Text = richTextBox3.Text + "$ChannelBandwidth_b1 := " + Bandwidth_L18 + "\n";
                richTextBox3.Text = richTextBox3.Text + "$tac := " + tac_L18 + "\n";
                richTextBox3.Text = richTextBox3.Text + "$eNBId := " + eNodeBID_L18 + "\n";
                double tmp_lat = Convert.ToSingle(Latitude_L18);
                tmp_lat = (tmp_lat / 90) * 8388608;
                tmp_lat = Math.Round(tmp_lat,0);
                double tmp_long = Convert.ToSingle(Longitude_L18);
                tmp_long = (tmp_long / 360) * 16777216;
                tmp_long = Math.Round(tmp_long,0);
                richTextBox3.Text = richTextBox3.Text + "$latitude := " + tmp_lat.ToString() + "\n";
                richTextBox3.Text = richTextBox3.Text + "$longitude := " + tmp_long.ToString() + "\n";
                int num_L26 = 0; int num_L18 = 0;
                for (int i = 1; i < cdd.GetLength(0); i++)
                {
                    string CellId = "";
                    for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "CellId") { CellId = cdd[i, k].ToString(); } }
                    if ((CellId == "7") || (CellId == "8") || (CellId == "9")) { num_L18++; }
                    else continue;
                    for (int j = 1; j < cdd.GetLength(1); j++)
                    {
                        if (cdd[0, j] == "CellName") { richTextBox3.Text = richTextBox3.Text + "$CELL_b1[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n"; }
                        if (cdd[0, j] == "CellId") { richTextBox3.Text = richTextBox3.Text + "$cellId_b1[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n"; }
                        if (cdd[0, j] == "Physical Cell ID Group") { richTextBox3.Text = richTextBox3.Text + "$physicalLayerCellIdGroup_b1[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n"; }
                        if (cdd[0, j] == "Physical Layer ID") { richTextBox3.Text = richTextBox3.Text + "$physicalLayerSubCellId_b1[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n"; }
                        if (cdd[0, j] == "RootSequenceIdx") { richTextBox3.Text = richTextBox3.Text + "$rachRootSequence_b1[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n"; }
                        if (cdd[0, j] == "Altitude")
                        {
                            while (cdd[i, j].Length < 4) cdd[i, j] = cdd[i, j] + "0";
                            richTextBox3.Text = richTextBox3.Text + "$altitude_b1[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n";
                        }
                    }
                }
            }
            //L18+L26
            if (L26 & L18)
            {
                //richTextBox3.Text = richTextBox3.Text + "$numbandsLTE := 2" + "\n";
                //richTextBox3.Text = richTextBox3.Text + "$LTEband_b1 := 2600" + "\n" + "$LTEband_b2 := 1800" + "\n";
                //richTextBox3.Text = richTextBox3.Text + "$numsectorsLTE_b1 := " + num_L26_cells.ToString() + "\n";
                //richTextBox3.Text = richTextBox3.Text + "$numsectorsLTE_b2 := " + num_L18_cells.ToString() + "\n";
                richTextBox3.Text = richTextBox3.Text + "$earfcndl_b1 := " + earfcndl_L26 + "\n";
                richTextBox3.Text = richTextBox3.Text + "$earfcndl_b2 := " + earfcndl_L18 + "\n";
                richTextBox3.Text = richTextBox3.Text + "$ChannelBandwidth_b1 := " + Bandwidth_L26 + "\n";
                richTextBox3.Text = richTextBox3.Text + "$ChannelBandwidth_b2 := " + Bandwidth_L18 + "\n";
                richTextBox3.Text = richTextBox3.Text + "$tac := " + tac_L26 + "\n";
                richTextBox3.Text = richTextBox3.Text + "$eNBId := " + eNodeBID_L26 + "\n";
                double tmp_lat = Convert.ToSingle(Latitude_L26);
                tmp_lat = (tmp_lat / 90) * 8388608;
                tmp_lat = Math.Round(tmp_lat,0);
                double tmp_long = Convert.ToSingle(Longitude_L26);
                tmp_long = (tmp_long / 360) * 16777216;
                tmp_long = Math.Round(tmp_long,0);
                richTextBox3.Text = richTextBox3.Text + "$latitude := " + tmp_lat.ToString() + "\n";
                richTextBox3.Text = richTextBox3.Text + "$longitude := " + tmp_long.ToString() + "\n";
                int num_L26 = 0; int num_L18 = 0;
                for (int i = 1; i < cdd.GetLength(0); i++)
                {
                    string CellId = "";
                    for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "CellId") { CellId = cdd[i, k].ToString(); } }
                    if ((CellId == "4") || (CellId == "5") || (CellId == "6")) { num_L26++; }
                    else
                    {
                        if ((CellId == "7") || (CellId == "8") || (CellId == "9")) { num_L18++; }
                        else continue;
                    }
                    for (int j = 1; j < cdd.GetLength(1); j++)
                    {                        
                        switch (CellId)
                        {
                                case "4":
                                case "5":
                                case "6":
                                    if (cdd[0, j] == "CellId") { richTextBox3.Text = richTextBox3.Text + "$cellId_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n"; }
                                    if (cdd[0, j] == "CellName") { richTextBox3.Text = richTextBox3.Text + "$CELL_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n"; }
                                    if (cdd[0, j] == "Physical Cell ID Group") { richTextBox3.Text = richTextBox3.Text + "$physicalLayerCellIdGroup_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n"; }
                                    if (cdd[0, j] == "Physical Layer ID") { richTextBox3.Text = richTextBox3.Text + "$physicalLayerSubCellId_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n"; }
                                    if (cdd[0, j] == "RootSequenceIdx") { richTextBox3.Text = richTextBox3.Text + "$rachRootSequence_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n"; }
                                    if (cdd[0, j] == "Altitude")
                                    {
                                    while (cdd[i, j].Length < 4) cdd[i, j] = cdd[i, j] + "0";
                                    richTextBox3.Text = richTextBox3.Text + "$altitude_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n";
                                    }
                                    break;
                                case "7":
                                case "8":
                                case "9":
                                    if (cdd[0, j] == "CellId") { richTextBox3.Text = richTextBox3.Text + "$cellId_b2[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n"; }
                                    if (cdd[0, j] == "CellName") { richTextBox3.Text = richTextBox3.Text + "$CELL_b2[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n"; }
                                    if (cdd[0, j] == "Physical Cell ID Group") { richTextBox3.Text = richTextBox3.Text + "$physicalLayerCellIdGroup_b2[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n"; }
                                    if (cdd[0, j] == "Physical Layer ID") { richTextBox3.Text = richTextBox3.Text + "$physicalLayerSubCellId_b2[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n"; }
                                    if (cdd[0, j] == "RootSequenceIdx") { richTextBox3.Text = richTextBox3.Text + "$rachRootSequence_b2[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n"; }
                                    if (cdd[0, j] == "Altitude")
                                {
                                    while (cdd[i, j].Length < 4) cdd[i, j] = cdd[i, j] + "0";
                                    richTextBox3.Text = richTextBox3.Text + "$altitude_b2[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n";
                                }
                                    break;
                        
                        }
                    }
                }

            }
            
            //ARS
            //только ARSL26
            if (ARS_L26 & !ARS_L18)
            {
                richTextBox3.Text = richTextBox3.Text + "$numbandsLTE_SHARED := 1" + "\n";
                richTextBox3.Text = richTextBox3.Text + "$LTEband_SHARED_b1 := 2600" + "\n" + "$LTEband_SHARED_b2 := 1800" + "\n";
                richTextBox3.Text = richTextBox3.Text + "$numsectorsLTE_SHARED_b1 := " + num_ARS_L26_cells.ToString() + "\n";
                richTextBox3.Text = richTextBox3.Text + "$earfcndl_SHARED_b1 := " + earfcndl_ARS_L26 + "\n";
                richTextBox3.Text = richTextBox3.Text + "$ChannelBandwidth_SHARED_b1 := " + Bandwidth_ARS_L26 + "\n";
                richTextBox3.Text = richTextBox3.Text + "$tac_SHARED := " + tac_ARS_L26 + "\n";
                int num_L26 = 0; int num_L18 = 0;
                for (int i = 1; i < cdd.GetLength(0); i++)
                {
                    string CellId = "";
                    for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "CellId") { CellId = cdd[i, k].ToString(); } }
                    if ((CellId == "14") || (CellId == "15") || (CellId == "16")) num_L26++; else continue;
                    for (int j = 1; j < cdd.GetLength(1); j++)
                    {                                              
                        if (cdd[0, j] == "CellId") { richTextBox3.Text = richTextBox3.Text + "$cellId_SHARED_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n"; }
                        if (cdd[0, j] == "CellName") { richTextBox3.Text = richTextBox3.Text + "$CELL_SHARED_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n"; }
                        if (cdd[0, j] == "Physical Cell ID Group") { richTextBox3.Text = richTextBox3.Text + "$physicalLayerCellIdGroup_SHARED_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n"; }
                        if (cdd[0, j] == "Physical Layer ID") { richTextBox3.Text = richTextBox3.Text + "$physicalLayerSubCellId_SHARED_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n"; }
                        if (cdd[0, j] == "RootSequenceIdx") { richTextBox3.Text = richTextBox3.Text + "$rachRootSequence_SHARED_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n"; }
                        
                    }
                }
            }
            //только ARSL18
            if (!ARS_L26 & ARS_L18)
            {
                richTextBox3.Text = richTextBox3.Text + "$numbandsLTE_SHARED := 1" + "\n";
                richTextBox3.Text = richTextBox3.Text + "$LTEband_SHARED_b1 := 1800" + "\n" + "$LTEband_SHARED_b2 := 2600" + "\n";
                richTextBox3.Text = richTextBox3.Text + "$numsectorsLTE_SHARED_b1 := " + num_ARS_L18_cells.ToString() + "\n";
                richTextBox3.Text = richTextBox3.Text + "$earfcndl_SHARED_b1 := " + earfcndl_ARS_L18 + "\n";
                richTextBox3.Text = richTextBox3.Text + "$ChannelBandwidth_SHARED_b1 := " + Bandwidth_ARS_L18 + "\n";
                richTextBox3.Text = richTextBox3.Text + "$tac_SHARED := " + tac_ARS_L18 + "\n";
                int num_L26 = 0; int num_L18 = 0;
                for (int i = 1; i < cdd.GetLength(0); i++)
                {
                    string CellId = "";
                    for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "CellId") { CellId = cdd[i, k].ToString(); } }
                    if ((CellId == "17") || (CellId == "18") || (CellId == "19")) num_L18++; else continue;
                    for (int j = 1; j < cdd.GetLength(1); j++)
                    {                        
                        if (cdd[0, j] == "CellId") { richTextBox3.Text = richTextBox3.Text + "$cellId_SHARED_b1[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n"; }
                        if (cdd[0, j] == "CellName") { richTextBox3.Text = richTextBox3.Text + "$CELL_SHARED_b1[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n"; }
                        if (cdd[0, j] == "Physical Cell ID Group") { richTextBox3.Text = richTextBox3.Text + "$physicalLayerCellIdGroup_SHARED_b1[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n"; }
                        if (cdd[0, j] == "Physical Layer ID") { richTextBox3.Text = richTextBox3.Text + "$physicalLayerSubCellId_SHARED_b1[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n"; }
                        if (cdd[0, j] == "RootSequenceIdx") { richTextBox3.Text = richTextBox3.Text + "$rachRootSequence_SHARED_b1[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n"; }                                                                         
                    }
                }
            }
            //ARSL26+ARSL18
            if (ARS_L26 & ARS_L18)
            {
                richTextBox3.Text = richTextBox3.Text + "$numbandsLTE_SHARED := 2" + "\n";
                richTextBox3.Text = richTextBox3.Text + "$LTEband_SHARED_b1 := 2600" + "\n" + "$LTEband_SHARED_b2 := 1800" + "\n";
                richTextBox3.Text = richTextBox3.Text + "$numsectorsLTE_SHARED_b1 := " + num_ARS_L26_cells.ToString() + "\n";
                richTextBox3.Text = richTextBox3.Text + "$numsectorsLTE_SHARED_b2 := " + num_ARS_L18_cells.ToString() + "\n";
                richTextBox3.Text = richTextBox3.Text + "$earfcndl_SHARED_b1 := " + earfcndl_ARS_L26 + "\n";
                richTextBox3.Text = richTextBox3.Text + "$earfcndl_SHARED_b2 := " + earfcndl_ARS_L18 + "\n";
                richTextBox3.Text = richTextBox3.Text + "$ChannelBandwidth_SHARED_b1 := " + Bandwidth_ARS_L26 + "\n";
                richTextBox3.Text = richTextBox3.Text + "$ChannelBandwidth_SHARED_b2 := " + Bandwidth_ARS_L18 + "\n";
                richTextBox3.Text = richTextBox3.Text + "$tac_SHARED := " + tac_ARS_L26 + "\n";
                int num_L26 = 0; int num_L18 = 0;
                for (int i = 1; i < cdd.GetLength(0); i++)
                {
                    string CellId = "";
                    for (int k = 1; k < cdd.GetLength(1); k++) { if (cdd[0, k] == "CellId") { CellId = cdd[i, k].ToString(); } }
                    if ((CellId == "14") || (CellId == "15") || (CellId == "16")) num_L26++;
                    else
                    if ((CellId == "17") || (CellId == "18") || (CellId == "19")) num_L18++; else continue;
                    for (int j = 1; j < cdd.GetLength(1); j++)
                    {
                        switch (CellId)
                        {
                            case "14":
                            case "15":
                            case "16":
                                if (cdd[0, j] == "CellId") { richTextBox3.Text = richTextBox3.Text + "$cellId_SHARED_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n"; }
                                if (cdd[0, j] == "CellName") { richTextBox3.Text = richTextBox3.Text + "$CELL_SHARED_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n"; }
                                if (cdd[0, j] == "Physical Cell ID Group") { richTextBox3.Text = richTextBox3.Text + "$physicalLayerCellIdGroup_SHARED_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n"; }
                                if (cdd[0, j] == "Physical Layer ID") { richTextBox3.Text = richTextBox3.Text + "$physicalLayerSubCellId_SHARED_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n"; }
                                if (cdd[0, j] == "RootSequenceIdx") { richTextBox3.Text = richTextBox3.Text + "$rachRootSequence_SHARED_b1[" + num_L26.ToString() + "] := " + cdd[i, j] + "\n"; }                                
                                break;
                            case "17":
                            case "18":
                            case "19":
                                if (cdd[0, j] == "CellId") { richTextBox3.Text = richTextBox3.Text + "$cellId_SHARED_b2[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n"; }
                                if (cdd[0, j] == "CellName") { richTextBox3.Text = richTextBox3.Text + "$CELL_SHARED_b2[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n"; }
                                if (cdd[0, j] == "Physical Cell ID Group") { richTextBox3.Text = richTextBox3.Text + "$physicalLayerCellIdGroup_SHARED_b2[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n"; }
                                if (cdd[0, j] == "Physical Layer ID") { richTextBox3.Text = richTextBox3.Text + "$physicalLayerSubCellId_SHARED_b2[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n"; }
                                if (cdd[0, j] == "RootSequenceIdx") { richTextBox3.Text = richTextBox3.Text + "$rachRootSequence_SHARED_b2[" + num_L18.ToString() + "] := " + cdd[i, j] + "\n"; }                                
                                break;

                        }
                    }
                }
            }


            //string[] results = new string[str.GetLength(0)];
            //for (int i = 0; i < results.Length; i++)
            //{
            //    results[i] = str[i].Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            //}

            //Парсинг строк по разным массивам
            //for (int i = 0; i < str.Length; i++)
            //{
            //    string[] str_0 = str[0].Split(new char[] { '\t' }); //str_0[] - поля заголовока
            //}



            ////Определяем кол-во секторов и диапазоны
            ////определяем диапазоны L18, L26, ARSL18, ARSL26. пределяем по cid
            //string L26 = 0; string L18 = 0; string ARSL26 = 0; string ARSL18 = 0;
            ////CellId
            //for (int i = 0; i < str_0.Length; i++)
            //{
            //    if (str_0[i] == "CellId")
            //    {
            //        for (int j = 1; j < str.Length; j++)
            //        {

            //        }


            //    }

            //}





            //if (str.Length != 2)
            //{
            //    MessageBox.Show("Нужно 2 строки (заголовок и данные)!");
            //    textBox5.Clear();
            //    return;
            //}
            //Анализ
            //string[] str_0 = str[0].Split(new char[] { '\t' }); //str_0[] - поля заголовока
            //string[] str_1 = str[1].Split(new char[] { '\t' }); //str_1[] - поля данных
            //for (int i = 0; i < str_0.Length; i++)
            //{
            //    //ABIS
            //    if (str_0[i] == "ABIS")
            //    {
            //        //провека на корректность ip (содержит 4 точки)                    
            //        int count = str_1[i].ToCharArray().Where(j => j == '.').Count();
            //        if (count != 3) continue;
            //        string abis_mask = str_1[i + 1].Replace(" ", "");
            //        switch (abis_mask)
            //        {
            //            case "255.255.255.224":
            //                abis_mask = "/27";
            //                break;
            //            case "255.255.255.240":
            //                abis_mask = "/28";
            //                break;
            //            case "255.255.255.252":
            //                abis_mask = "/30";
            //                break;
            //            default:
            //                abis_mask = "/???????????";
            //                break;
            //        }
            //        richTextBox3.Text = richTextBox3.Text + "$GSM_IP := " + str_1[i].Replace(" ", "") + abis_mask + "\n";
            //        richTextBox3.Text = richTextBox3.Text + "$GSM_DG := " + str_1[i + 2].Replace(" ", "") + "\n";
            //        richTextBox3.Text = richTextBox3.Text + "$GSM_Vlan := " + str_1[i + 3].Replace(" ", "") + "\n";
            //        //Bridge                    
            //        if (richTextBox3.Text.IndexOf("$Bridge := 1") != -1)
            //        {
            //            if (TNCascade.Substring(TNCascade.Length - 2) == "2G")
            //            {
            //                richTextBox3.Text = richTextBox3.Text + "$OAMVLAN := " + str_1[i + 7].Replace(" ", "") + "\n";
            //                richTextBox3.Text = richTextBox3.Text + "$TRAFFICVLAN := " + str_1[i + 3].Replace(" ", "") + "\n";
            //            }
            //        }

            //    }
            //}
        }
    }
}
