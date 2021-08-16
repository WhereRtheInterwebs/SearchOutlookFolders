using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using Microsoft.Office.Interop.Outlook;

namespace SearchOutlookFolders
{
    [Cmdlet(VerbsCommon.Search, "OutlookFolders")]
    [OutputType(typeof(SearchResult))]
    public class SearchFolders : Cmdlet
    {
        [Parameter(Position = 1)]
        public string SearchText { get; set; } = "";

        [Parameter(Position = 2)]
        public string Store { get; set; }

        protected override void ProcessRecord()
        {
            Application app = new Application();
            NameSpace ns = app.GetNamespace("MAPI");
            MAPIFolder rootFolder = ns.DefaultStore.GetRootFolder();

            if (Store != null)
                rootFolder = ns.Stores[Store].GetRootFolder();

            List<Folder> allFolders = GetFolders(rootFolder.Folders).ToList();

            allFolders.Where(x => x.Name.ToLower().Contains(SearchText.ToLower()))
                      .Select(x => new SearchResult
                      {
                          Name     = x.Name,
                          FullPath = x.FullFolderPath
                      })
                      .ToList()
                      .ForEach(WriteObject);
        }

        IEnumerable<Folder> GetFolders(Folders folders)
        {
            foreach (Folder folder in folders)
            {
                yield return folder;
                WriteVerbose(folder.FullFolderPath);
                if (folder.Folders != null)
                    foreach (Folder subFolder in GetFolders(folder.Folders))
                        yield return subFolder;
            }
        }

        [Cmdlet(VerbsCommon.Get, "OutlookStores")]
        [OutputType(typeof(string))]
        public class GetStores : Cmdlet
        {
            protected override void ProcessRecord()
            {
                Application app = new Application();
                NameSpace ns = app.GetNamespace("MAPI");
                ns.Stores.Cast<Store>().Select(x => x.DisplayName).ToList().ForEach(WriteObject);
            }
        }
    }

    public class SearchResult
    {
        public string Name { get; set; }
        public string FullPath { get; set; }
    }
}
