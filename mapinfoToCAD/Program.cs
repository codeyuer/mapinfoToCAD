using System;
using System.Collections.Generic;
using System.Diagnostics;

using System.Text;

using Autodesk.AutoCAD;
using System.Reflection;
using System.Threading;
using System.IO;

namespace mapinfoToCAD
{

    
    class Program
    {
        public static void CloseAllInstance()
        {
            Process[] aCAD =
               Process.GetProcessesByName("acad");

            foreach (Process aCADPro in aCAD)
            {
                aCADPro.CloseMainWindow();
            }
        }

        /// <summary>
        /// 处理dxf文件
        /// 打开文件，按范围缩放，保存
        /// </summary>
        /// <param name="procFile"></param>
        public static void procFile(string procFile)
        {
            //procFile文件夹，要处理的文件路劲
            int totalCount = Searchfile(procFile);
            DirectoryInfo folder = new DirectoryInfo(procFile);


            //try
            //{
                //cadapp.Visible = true;
                float i = 0;
                foreach (FileInfo file in folder.GetFiles("*.dxf"))
                {
                    
                    caddoc = cadapp.Documents.Open(file.FullName.ToString(), Missing.Value, Missing.Value);

                    
                        caddoc.Activate();

                    //string lispPath = "d:/2.lsp";

                    //string loadStr = String.Format("(appload \"{0}\")  tttt\n"/*space after closing paren!!!*/, lispPath);

                    //caddoc.SendCommand(loadStr);
                    

                        string cadCmd = "2 " + file.Name.Substring(0, file.Name.IndexOf(file.Extension) + 1) + " ";

                        caddoc.SendCommand(cadCmd);
                        caddoc.Close();
                caddoc = null;
                    //while ((short)caddoc.GetVariable("CMDACTIVE") == 1)
                    //{
                    //    Thread.Sleep(100);
                    //}

                    i++;
                    Console.WriteLine(file.FullName + "...done" + ((i / totalCount) * 100).ToString() + "%");
                    caddoc = null;
                }

            //}
            //catch(Exception ex)
            //{
            //    Console.WriteLine("程序出错.."+ex.Message.ToString());
            //}
        }
        /// <summary>
        ///读取文件夹内文件数量
        /// </summary>
        /// <param name="Directory"></param>
        /// <returns></returns>
        public static int Searchfile(string Directory)
        {
            int countfile = 0;//文件数量
            int countdir = 0;//文件夹数量
            DirectoryInfo dir = new DirectoryInfo(Directory);
            FileSystemInfo[] fi = dir.GetFileSystemInfos();//获取文件夹下的文件

            foreach (FileSystemInfo f in fi)
            {
                if (f is DirectoryInfo) //判断是否为文件夹
                {
                    countdir += 1;
                    Searchfile(f.FullName); //递归调用

                }
                else
                {
                    countfile += 1;
                }
            }

            return countfile;


        }
       static  Autodesk.AutoCAD.Interop.AcadApplication cadapp = null; //cad
       static  Autodesk.AutoCAD.Interop.AcadDocument caddoc = null;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="args">0.cad .dxf文件夹路径  1.处理数据类型</param>
        public static void Main(string[] args)
        {
            CloseAllInstance();
            Console.WriteLine("开始处理..");
            Console.WriteLine("正在打开cad");
            cadapp = new Autodesk.AutoCAD.Interop.AcadApplication();
            cadapp.Visible = true;
            string procType = args[1];
            switch (procType)
            {
                case "t1":
                    Console.WriteLine("正在处理方式1"+args[0]+"  "+args[1]);
                    Program.procFile(args[0]);
                    break;


            }
                

          

                


        }
    }
}
