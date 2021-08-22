using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.IO;
using System.Diagnostics;
using System.Threading.Tasks;

namespace Smv.DispUpdate
{
  sealed class Program
  {
    private static void CopyAll(DirectoryInfo source, DirectoryInfo target)
    {
      if (source.FullName.ToLower() == target.FullName.ToLower())
        return;
 
      //Check if the target directory exists, if not, create it.
      if (Directory.Exists(target.FullName) == false)
        Directory.CreateDirectory(target.FullName);
      
      //Copy each file into it's new directory.
      foreach (FileInfo fi in source.GetFiles()){
        //Console.WriteLine(@"Copying {0}\{1}", target.FullName, fi.Name);
        fi.CopyTo(Path.Combine(target.ToString(), fi.Name), true);
      }

      //Copy each subdirectory using recursion.
      foreach (DirectoryInfo diSourceSubDir in source.GetDirectories()){
        DirectoryInfo nextTargetSubDir = target.CreateSubdirectory(diSourceSubDir.Name);
        CopyAll(diSourceSubDir, nextTargetSubDir);
      }
    }

    //Копирование всех папок c содержимым из директории источника во другую директорию

    /*
    foreach (var f in di.EnumerateFiles("*")){
      Console.WriteLine(f.FullName);
    }

    foreach (var f in di.EnumerateFiles("*")){
      Console.WriteLine(f.FullName);
    } 
    */

    static void Main(string[] args)
    {
      int cX;
      int cY;

      Console.Title = "Утилита обновления ПО";
      Console.ForegroundColor = ConsoleColor.Yellow;
      Console.WriteLine("Начался процесс обновления ПО...");
      Console.WriteLine("------");
      Console.WriteLine("");
      cX = Console.CursorLeft;
      cY = Console.CursorTop;   

      //Console.WriteLine(args.Length.ToString());
      //Console.WriteLine(args[0]);
      /* 0-UPDATE,  1-получатель,  2-источник*/
      if ((args.Length != 3) || (args[0] != "UPDATE")){
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("Ошибка обновления:");
        Console.WriteLine("Входной параметр программы обновления отсутствует или не верен!");
        Console.WriteLine("Нажмите любую клавишу для выхода");
        Console.ReadKey();
        return;
      }
      
      Process proc = Process.GetProcesses().FirstOrDefault(p => p.ProcessName.StartsWith("Smv.Modules.MgrExt", StringComparison.InvariantCultureIgnoreCase));
      if (proc == null){
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("Ошибка обновления:");
        Console.WriteLine("Процесс основного приложения не найден!");
        Console.WriteLine("Нажмите любую клавишу для выхода");
        Console.ReadKey();
        return;
      }
     
      proc.CloseMainWindow();
      Thread.Sleep(1000);

      //Готовим временную папку
      var tPath = Path.GetTempPath() + "\\Lims";
      if (Directory.Exists(tPath))
        Directory.Delete(tPath, true);

      Directory.CreateDirectory(tPath);

      //Копирование всех папок из директории источника во временную директорию
      var di = new DirectoryInfo(args[2]);
      if (!di.Exists){
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("Ошибка обновления:");
        Console.WriteLine("Папки источника не существует!");
        Console.WriteLine("Нажмите любую клавишу для выхода");
        Console.ReadKey();
        return;
      }

      var dir = from d in di.EnumerateDirectories()
                where d.Name.ToUpper() != "XCONFIGX"
                select d;

      foreach (var d in dir){
        //Console.WriteLine(d.Name);
        CopyAll(new DirectoryInfo(d.FullName), new DirectoryInfo(tPath + "\\" + d.Name));
      }
      Console.SetCursorPosition(0, 1);
      Console.ForegroundColor = ConsoleColor.Yellow;
      Console.Write('*');
      Console.SetCursorPosition(cX, cY);

      //Копирование всех файлов из директории источника во временную директорию
      var fl = from f in di.EnumerateFiles("*")
               where f.Name.ToUpper() != "ZADFG"
               select f;

      foreach (var f in fl){
        //Console.WriteLine(f.Name);
        File.Copy(f.FullName, tPath + "\\" + f.Name, true);
      }
      Console.SetCursorPosition(1, 1);
      Console.ForegroundColor = ConsoleColor.Yellow;
      Console.Write('*');
      Console.SetCursorPosition(cX, cY);


      //Удаление всех папок в корневой директории получателя       
      di = new DirectoryInfo(args[1]);
      if (!di.Exists){
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("Ошибка обновления:");
        Console.WriteLine("Папки получателя не существует!");
        Console.WriteLine("Нажмите любую клавишу для выхода");
        Console.ReadKey();
        return;
      }

      dir = from d in di.EnumerateDirectories()
            where d.Name.ToUpper() != "CONFIG"
            select d;

      foreach(var d in dir){
        //Console.WriteLine(d.Name);
        Directory.Delete(d.FullName, true);
      }
      Console.SetCursorPosition(2,1);
      Console.ForegroundColor = ConsoleColor.Yellow;
      Console.Write('*');
      Console.SetCursorPosition(cX,cY);

      //Удаление всех файлов в корневой директории получателя
      fl = from f in di.EnumerateFiles("*")
           where f.Name.ToUpper() != "SMV.DISPUPDATE.EXE"
           select f;

      foreach(var f in fl){
        //Console.WriteLine(f.Name);
        File.Delete(f.FullName);
      }
      Console.SetCursorPosition(3, 1);
      Console.ForegroundColor = ConsoleColor.Yellow;
      Console.Write('*');
      Console.SetCursorPosition(cX, cY);

      
      //Копирование всех папок из временной директории в директорию получателя
      di = new DirectoryInfo(tPath);
      if (!di.Exists){
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("Ошибка обновления:");
        Console.WriteLine("Временная папки не существует!");
        Console.WriteLine("Нажмите любую клавишу для выхода");
        Console.ReadKey();
        return;
      }

      dir = from d in di.EnumerateDirectories()
            where d.Name.ToUpper() != "CONFIG"
            select d;

      foreach(var d in dir){
        //Console.WriteLine(d.Name);
        CopyAll(new DirectoryInfo(d.FullName), new DirectoryInfo(args[1] + "\\" + d.Name));
      }
      Console.SetCursorPosition(4, 1);
      Console.ForegroundColor = ConsoleColor.Yellow;
      Console.Write('*');
      Console.SetCursorPosition(cX, cY);


      //Копирование всех файлов из временной директории в директорию получателя
      fl = from f in di.EnumerateFiles("*")
           where f.Name.ToUpper() != "SMV.DISPUPDATE.EXE"
           select f;
        
      foreach (var f in fl){
        //Console.WriteLine(f.Name);
        File.Copy(f.FullName, args[1] + "\\" + f.Name, true);
      }
      Console.SetCursorPosition(5, 1);
      Console.ForegroundColor = ConsoleColor.Yellow;
      Console.Write('*');
      Console.SetCursorPosition(cX, cY);

      Console.ForegroundColor = ConsoleColor.White;
      Console.WriteLine("Обновление ПО закончено!");
      Console.ForegroundColor = ConsoleColor.Yellow;
      Console.WriteLine("Запускается обновленное ПО...");

      //Запускаем LIMS
      var startInfo = new ProcessStartInfo(args[1] + "\\Smv.Modules.MgrExt.exe");
      Process.Start(startInfo);
      
      Thread.Sleep(2000);

      //Console.WriteLine("Нажмите любую клавишу для выхода");
      //Console.ReadKey();
    }
  }
}
