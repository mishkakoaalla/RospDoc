using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KompasAPI7;
using Kompas6Constants;
using Kompas6API5;
using KAPITypes;

using System.Diagnostics;
using reference = System.Int32;
using System.Text.RegularExpressions;


namespace RospDoc
{
    public partial class Form1 : Form
    {
        List<string> path = new List<string>();
        List<string> path_name = new List<string>();
        public bool clear = false;
        


        public string r_bul = @"ToyusanuH";
        public string r_danil = @"DanuILob";
        public string r_krug = @"Kpyr/\ob";        
        public string r_neum = @"Heym";
        public string r_petrov = @"TIeTpob";
        public string r_rud = @"Pygakob";
        public string r_sorok = @"Copokuu";
        public string r_ufr = @"Yqrpymob";
        public string r_shal = @"lllarqrun";
        









        private KompasObject kompas;
        private IApplication appl;         // Интерфейс приложения
        public Form1()
        {
            InitializeComponent();
            GetKompas();
            dateTimePicker1.Value = DateTime.Now;
        }

        private void panel1_DragDrop(object sender, DragEventArgs e)
        {
            var allowedExtensions = new[] { ".cdw", ".spw" };


            foreach (string obj in (string[])e.Data.GetData(DataFormats.FileDrop))
                if (Directory.Exists(obj))
                {
                    // path.AddRange(Directory.GetFiles(obj, "*.*", SearchOption.AllDirectories)
                    //.Where(f=> f.EndsWith(".cdw")|| f.EndsWith(".spw")).ToArray()                

                    // );
                    //MessageBox.Show("Не вабраны файлы с расширением  .cdw или .spw");

                }
                else
                {
                    string q = Path.GetFileName(obj);
                    string w = Path.GetExtension(obj);

                    if (w == ".cdw" || w == ".spw")
                    {
                        path.Add(obj);
                        path_name.Add(q);
                        Console.WriteLine("Документ №" +(path_name.Count)+ "."+ w);
                    }
                }
            label1.Text = string.Join("\r\n", path_name);


            // label1.Text += file + "\n";
        }

        private void panel1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            START();
        }







        void GetKompas()
        {
            try
            {

                kompas = (KompasObject)System.Runtime.InteropServices.Marshal.GetActiveObject("kompas.application.5");
                appl = (IApplication)kompas.ksGetApplication7();
                //MessageBox.Show("Подключение установлено");
                Console.WriteLine("Подключение установлено");
                appl.KompasError.Clear();

            }
            catch
            {
                //MessageBox.Show("Компас не запущен - ЗАПУСКАЕМ ");
                Console.WriteLine("Компас не запущен - ЗАПУСКАЕМ");
                Type t = Type.GetTypeFromProgID("KOMPAS.Application.5");
                kompas = (KompasObject)Activator.CreateInstance(t);
                kompas = (KompasObject)System.Runtime.InteropServices.Marshal.GetActiveObject("kompas.application.5");
                appl = (IApplication)kompas.ksGetApplication7();
                kompas.Visible = true;  //  
                appl.KompasError.Clear();
                //kompas.ActivateControllerAPI();
            }
        }



        void START()
        {
            Console.WriteLine("Количество документов = " + path.Count);
            ksTextItemParam item = null;


            for (int i = 0; i < path.Count; i++)
            {
               

                Console.WriteLine("");
                string w = Path.GetExtension(path[i]);
                Console.WriteLine("Расширение =  " + w);

                if (w == ".cdw")
                {

                    if (clear == false)
                    {
                        Console.WriteLine("START------------");

                        if (comboBox2.SelectedItem.ToString() == "Проверил")
                        {
                            Doc2D(121, 131, textBox2.Text, dateTimePicker1.Text.ToString(), i, true);
                            
                        }

                       if (comboBox2.SelectedItem.ToString() == "Разработал")
                        {
                            Doc2D(120, 130, textBox2.Text, dateTimePicker1.Text.ToString(), i, false);
                        }




                    }
                    else
                    {
                        if (comboBox2.SelectedItem.ToString() == "Проверил")
                        {
                            Doc2D(121, 131, " "," ", i, false);
                            Doc2D(123, 133, " "," ", i, false);

                        }

                        if (comboBox2.SelectedItem.ToString() == "Разработал")
                        {
                            Doc2D(120, 130, " "," ", i, false);
                        }

                    }

                }
                else
                {
                    if (w == ".spw")

                        if (clear == false)
                        {
                            Console.WriteLine("START------------");

                            if (comboBox2.SelectedItem.ToString() == "Проверил")
                            {
                                SpsDoc(121, 131, textBox2.Text, dateTimePicker1.Text.ToString(), i, true);

                            }

                            if (comboBox2.SelectedItem.ToString() == "Разработал")
                            {
                                SpsDoc(120, 130, textBox2.Text, dateTimePicker1.Text.ToString(), i, false);
                            }




                        }
                        else
                        {
                            if (comboBox2.SelectedItem.ToString() == "Проверил")
                            {
                                SpsDoc(121, 131, " ", " ", i, false);
                                SpsDoc(123, 133, " ", " ", i, false);

                            }

                            if (comboBox2.SelectedItem.ToString() == "Разработал")
                            {
                                SpsDoc(120, 130, " ", " ", i, false);
                            }

                        }


                }
            }





        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string vibor = comboBox1.SelectedItem.ToString();

            switch (vibor)
            {
                case "Булганин":
                    textBox2.Text = r_bul;
                    
                    break;
                case "Данилов":
                    textBox2.Text = r_danil;

                    break;

                case "Круглов":
                    textBox2.Text = r_krug;

                    break;

                case "Неумоин":
                    textBox2.Text = r_neum;

                    break;

                case "Петров":
                    textBox2.Text = r_petrov;

                    break;

                case "Рудаков":
                    textBox2.Text = r_rud;

                    break;

                case "Сорокин":
                    textBox2.Text = r_sorok;

                    break;

                case "Уфрутов":
                    textBox2.Text = r_ufr;

                    break;


                case "Шалягин":
                    textBox2.Text = r_shal;

                    break;


                    //default:
                    //Console.WriteLine("Default case");
                    //break;
            }

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox checkBox = (CheckBox)sender; // приводим отправителя к элементу типа CheckBox
            if (checkBox.Checked == true)
            {                
                button1.Text = "Очистить роспись";
                comboBox1.Enabled = false;
                textBox2.Enabled = false;
                dateTimePicker1.Enabled = false;
                clear = true;
            }
            else
            {
                button1.Text = "Вставить роспись";
                comboBox1.Enabled = true;
                textBox2.Enabled = true;
                dateTimePicker1.Enabled = true;
                clear = false;
            }
        }


        public void Doc2D(int n_str, int n_str_dat ,string text, string dat, int n_doc, bool ruk = false)
        {
            IKompasDocument doc = appl.Documents.Open(path[n_doc], true, false);// Получаем интерфейс активного документа 2D в API7
            ksDocument2D docD = (ksDocument2D)kompas.ActiveDocument2D();
            ksStamp stamp = (ksStamp)docD.GetStamp();

            Console.WriteLine("функция Doc2D ");


            stamp.ksOpenStamp();

            //_____________________________________________________________
            LayoutSheets _ls = doc.LayoutSheets;
            LayoutSheet LS = _ls.ItemByNumber[1];
            IStamp isamp = LS.Stamp;
            IText qq = isamp.Text[10];
            string str_ruk = qq.Str;
            LS.LayoutLibraryFileName = "C:\\Program Files\\ASCON\\KOMPAS-3D v20\\Sys\\graphic.lyt";
            LS.Update();

            Console.WriteLine("Есть ли руководитель -------------  " + qq.Str);
            //_____________________________________________________________


            stamp.ksColumnNumber(n_str);
            ksTextItemParam itemParam = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);

            //роспись
            if (itemParam != null)
            {
                itemParam.Init();

                ksTextItemFont itemFont = (ksTextItemFont)itemParam.GetItemFont();
                if (itemFont != null)
                {
                    itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
                    itemParam.s = text;
                    docD.ksTextLine(itemParam);
                }
            }

            //Дата
            stamp.ksColumnNumber(n_str_dat);
            
            if (itemParam != null)
            {
                itemParam.Init();

                ksTextItemFont itemFont = (ksTextItemFont)itemParam.GetItemFont();
                if (itemFont != null)
                {
                    itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
                    itemParam.s = dat;
                    docD.ksTextLine(itemParam);
                }
            }
            ///рук проекта
            if (ruk== true)
            {
                if (str_ruk != "")
                {
                    //роспись
                    stamp.ksColumnNumber(123);
                    itemParam.Init();
                    ksTextItemFont itemFont = (ksTextItemFont)itemParam.GetItemFont();
                    if (itemFont != null)
                    {
                        itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
                        itemParam.s = text;
                        docD.ksTextLine(itemParam);
                    }


                    //дата в росписи
                    stamp.ksColumnNumber(133);
                    itemParam.Init();
                    
                    if (itemFont != null)
                    {
                        itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
                        itemParam.s = dat;
                        docD.ksTextLine(itemParam);
                    }

                }
            }


            

            stamp.ksCloseStamp();



            doc.Close(DocumentCloseOptions.kdSaveChanges);
        }

        public void SpsDoc(int n_str, int n_str_dat, string text, string dat, int n_doc, bool ruk = false)
        {
            IKompasDocument doc = appl.Documents.Open(path[n_doc], true, false);// Получаем интерфейс активного документа 2D в API7                        
            ksSpcDocument DocS = (ksSpcDocument)kompas.SpcActiveDocument();
            ksStamp stamp = DocS.GetStamp();




            stamp.ksOpenStamp();

            //_______________________________________
            LayoutSheets _ls = doc.LayoutSheets;
            LayoutSheet LS = _ls.ItemByNumber[1];
            var q = _ls.ItemByNumber[1].Stamp;
            IStamp isamp = LS.Stamp;
            IText qq = isamp.Text[10];
            LS.LayoutLibraryFileName = "C:\\Program Files\\ASCON\\KOMPAS-3D v20\\Sys\\graphic.lyt";
            LS.Update();
            string str_ruk = qq.Str;
            Console.WriteLine("Есть ли руководитель -------------  " + qq.Str);
            //________________________________________

            stamp.ksColumnNumber(n_str);
            ksTextItemParam itemParam = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);

            //роспись
            if (itemParam != null)
            {
                itemParam.Init();

                ksTextItemFont itemFont = (ksTextItemFont)itemParam.GetItemFont();
                if (itemFont != null)
                {
                    itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
                    itemParam.s = text;
                    stamp.ksTextLine(itemParam);
                }
            }

            //Дата
            stamp.ksColumnNumber(n_str_dat);

            if (itemParam != null)
            {
                itemParam.Init();

                ksTextItemFont itemFont = (ksTextItemFont)itemParam.GetItemFont();
                if (itemFont != null)
                {
                    itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
                    itemParam.s = dat;
                    stamp.ksTextLine(itemParam);
                }
            }
            ///рук проекта
            if (ruk == true)
            {
                if (str_ruk != "")
                {
                    //роспись
                    stamp.ksColumnNumber(123);
                    itemParam.Init();
                    ksTextItemFont itemFont = (ksTextItemFont)itemParam.GetItemFont();
                    if (itemFont != null)
                    {
                        itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
                        itemParam.s = text;
                        stamp.ksTextLine(itemParam);
                    }


                    //дата в росписи
                    stamp.ksColumnNumber(133);
                    itemParam.Init();

                    if (itemFont != null)
                    {
                        itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
                        itemParam.s = dat;
                        stamp.ksTextLine(itemParam);
                    }

                }
            }




            stamp.ksCloseStamp();

            doc.Close(DocumentCloseOptions.kdSaveChanges); //Закрыть документ






















            //stamp.ksColumnNumber(120);
            //ksTextItemParam itemParam = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
            //if (itemParam != null)
            //{
            //    itemParam.Init();

            //    ksTextItemFont itemFont = (ksTextItemFont)itemParam.GetItemFont();
            //    if (itemFont != null)
            //    {
            //        itemFont.SetBitVectorValue(ldefin2d.NEW_LINE, true);
            //        itemFont.fontName = "Staccato222 BT";

            //        itemParam.s = "44444";
            //        stamp.ksTextLine(itemParam);

            //    }
            //}

            //stamp.ksCloseStamp();

        }


    }
}