using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace PSTfiles
{
    class Program
    {
        static void Main(string[] args)
        {
            string LogFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "PSTfiles.txt");

            Application app = new Application();
            NameSpace ns = app.GetNamespace("MAPI");
            MAPIFolder rootFolder = ns.DefaultStore.GetRootFolder();

            List<Folder> allFolders = GetFolders(rootFolder.Folders).ToList();

            File.WriteAllLines(LogFilePath, allFolders.Select(x => x.FullFolderPath).ToArray());

            string searchResult = "";
            if (args.Any())
            {
                searchResult = "Found: " + (allFolders.FirstOrDefault(x => x.Name.ToLower().Contains(args[0].ToLower()))?.FullFolderPath ?? "Nothing :(");
                Console.WriteLine(searchResult);
                File.AppendAllLines(LogFilePath, new[] { searchResult });
            }

            if (Debugger.IsAttached | true)
            {
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
            }
        }

        static IEnumerable<Folder> GetFolders(Folders folders)
        {
            foreach (Folder folder in folders)
            {
                yield return folder;
                Console.WriteLine(folder.FullFolderPath);
                if (folder.Folders != null)
                    foreach (Folder subFolder in GetFolders(folder.Folders))
                        yield return subFolder;
            }
        }
    }
}
