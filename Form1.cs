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



        private KompasObject kompas;
        private IApplication appl;         // Интерфейс приложения
        public Form1()
        {
            InitializeComponent();
            GetKompas();
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

                    IKompasDocument doc = appl.Documents.Open(path[i], true, false);// Получаем интерфейс активного документа 2D в API7

                    ksDocument2D docD = (ksDocument2D)kompas.ActiveDocument2D();
                    ksStamp stamp = (ksStamp)docD.GetStamp();

                    stamp.ksOpenStamp();
                    stamp.ksColumnNumber(110);
                    ksTextLineParam itemLineText = (ksTextLineParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextLineParam);
                    
                    itemLineText.Init();

                    itemLineText.style = 32768;
                    ksDynamicArray arrpLineText = (ksDynamicArray)kompas.GetDynamicArray(ldefin2d.TEXT_LINE_ARR);
                    item = (ksTextItemParam)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemParam);
                    item.Init();
                    item.s = "УУУУУУУХ";

                    ksTextItemFont itemFont = (ksTextItemFont)kompas.GetParamStruct((short)StructType2DEnum.ko_TextItemFont);
                    itemFont.Init();
                    itemFont.fontName = "GOST type A";


                    Console.WriteLine("Получение данных из документа № - " + Convert.ToInt32(i + 1));


                    stamp.ksCloseStamp();



                    


                    //doc.Close(0); //Закрыть документ

                }
                else
                {
                    if (w == ".spw")
                    {
                        Console.WriteLine("Пропущена спецификация");

                        IKompasDocument docS = appl.Documents.Open(path[i], true, false);// Получаем интерфейс активного документа 2D в API7     


                        LayoutSheets _ls = docS.LayoutSheets;                     
                        LayoutSheet LS = _ls.ItemByNumber[1];

                        var q = _ls.ItemByNumber[1].Stamp;

                        IStamp isamp = LS.Stamp;


                        IText qq = isamp.Text[111];
                        IText ww = isamp.Text[121];
                        Console.WriteLine("ШТАМП Проверил -------------  " + qq.Str);
                        Console.WriteLine("ШТАМП Роспись -------------  " + ww.Str);
                        


                    }

                    
                }
            }





        }
    }
}
