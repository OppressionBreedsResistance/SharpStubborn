using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace SharpStubborn
{
    public class Shortcut
    {
        public string Name { get; set; }
        public string TargetPath { get; set; }
        public string IconLocation {  get; set; }
        public string WorkingDirectory {  get; set; }
        public string Arguments { get; set; }

    }
    internal class Program
    {
        static void Main(string[] args)
        {
            // Create list of shortcuts
            List<Shortcut> shortcuts = new List<Shortcut>();
            getShortcuts();

            foreach (Shortcut shortcut in shortcuts) {
                Console.WriteLine("Nazwa to " + shortcut.Name);
                Console.WriteLine("Target path to " + shortcut.TargetPath);
                Console.WriteLine("Icon location to " + shortcut.IconLocation);
            }



            void getShortcuts()
            {
                // Pobranie sciezki do pulpitu
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string[] shortcuts_path = Directory.GetFiles(desktopPath, "*.lnk");

                if (shortcuts_path.Length == 0)
                {
                    Console.WriteLine("Nie znaleziono zadnych skrótów");
                }
                else
                {
                    Console.WriteLine("Znaleziono nastepujace skróty");
                    foreach (string shortcut_path in shortcuts_path)
                    {
                        Console.WriteLine(shortcut_path);
                        IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
                        IWshRuntimeLibrary.IWshShortcut link = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(shortcut_path);
                        Shortcut shortcut = new Shortcut
                        {
                            Name = shortcut_path,
                            TargetPath = link.TargetPath,
                            IconLocation = link.IconLocation,
                            WorkingDirectory = link.WorkingDirectory,
                            Arguments = link.Arguments
                        };
                        shortcuts.Add(shortcut);
                    }
                } 
            }

                // test
        }

    }
}
