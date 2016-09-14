using System;
using System.Diagnostics;
using Microsoft.Exchange.WebServices.Data;

namespace TestEWSSearchPerformance
{
    class Program
    {
        const int NumberOfMessagesToCreate = 10000;
        const int NumberOfSearchesToPerform = 50;
        const int NumberOfIterations = 5;
        static readonly Guid MyNamedPropertyNamespaceGuid = new Guid("6755f378-6d50-4b26-84ae-95824b6b7a1d");
        static readonly ExtendedPropertyDefinition MyNamedProp = 
            new ExtendedPropertyDefinition(
                MyNamedPropertyNamespaceGuid, 
                "SearchingMailboxesViaEWSCustomProp",
                MapiPropertyType.Long);
        const string SearchFolderName = "TestEWSSearchPerformanceSearchFolder";

        static void Main(string[] args)
        {
            var smtpAddressOfMailbox = args[0];
            var exchService = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            exchService.AutodiscoverUrl(smtpAddressOfMailbox, foo => true);
            if (exchService.Url == null)
            {
                Console.WriteLine("Autodiscover failed");
                return;
            }

            var mailbox = new Mailbox(smtpAddressOfMailbox);
            var inboxFolderId = new FolderId(WellKnownFolderName.Inbox, mailbox);
            var inboxFolder = Folder.Bind(exchService, inboxFolderId);

            if (inboxFolder.TotalCount < NumberOfMessagesToCreate)
            {
                CreateLotsOfMessages(exchService, inboxFolderId);
            }

            var searchFolder = CreateSearchFolderIfNeeded(exchService, mailbox);

            Console.WriteLine();
            Console.WriteLine("Seeking the folder directly.");
            for (var x = 0; x < NumberOfIterations; x++)
            {
                SearchBySeek(inboxFolder);
            }

            Console.WriteLine();
            Console.WriteLine("Seeking the search folder.");
            for (var x = 0; x < NumberOfIterations; x++)
            {
                SearchBySeek(searchFolder);
            }

            Console.WriteLine();
            Console.WriteLine("FindItems on the folder without sorting.");
            for (var x = 0; x < NumberOfIterations; x++)
            {
                SearchByFilter(inboxFolder, false);
            }

            Console.WriteLine();
            Console.WriteLine("FindItems on the folder with a sort applied to the view.");
            for (var x = 0; x < NumberOfIterations; x++)
            {
                SearchByFilter(inboxFolder, true);
            }

            Console.WriteLine();
            Console.Write("Done. Hit Enter to exit.");
            Console.ReadLine();
        }

        public static void SearchBySeek(Folder folder)
        {
            var rnd = new Random();
            var view = new ItemView(1, 0);
            view.OrderBy.Add(MyNamedProp, SortDirection.Ascending);
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly, MyNamedProp);
            var sw = new Stopwatch();
            var filter = new SearchFilter.Exists(MyNamedProp);
            sw.Start();
            for (var x = 0; x < NumberOfSearchesToPerform; x++)
            {
                var numberToSearchFor = rnd.Next(0, NumberOfMessagesToCreate);
                int high = folder.FindItems(filter, view).TotalCount - 1;
                var low = 0;
                var found = false;
                while (!found)
                {
                    view.Offset = ((high - low)/2) + low;
                    var resultItem = folder.FindItems(filter, view);
                    if (resultItem.Items.Count < 1) break;

                    long myId;
                    resultItem.Items[0].TryGetProperty(MyNamedProp, out myId);
                    if (myId == numberToSearchFor)
                    {
                        found = true;
                    }
                    else
                    {
                        if (myId > numberToSearchFor)
                        {
                            high = (int)myId - 1;
                        }
                        else
                        {
                            low = (int)myId + 1;
                        }
                    }
                }

                if (!found) Console.WriteLine("Warning! No item found!");
            }

            sw.Stop();
            Console.WriteLine("SearchBySeek finished after: " + sw.ElapsedMilliseconds + " milliseconds.");
        }

        public static void SearchByFilter(Folder folder, bool sort)
        {
            var rnd = new Random();
            var view = new ItemView(1, 0);
            if (sort) view.OrderBy.Add(MyNamedProp, SortDirection.Ascending);
            var sw = new Stopwatch();
            sw.Start();
            for (var x = 0; x < NumberOfSearchesToPerform; x++)
            {
                var numberToSearchFor = rnd.Next(0, NumberOfMessagesToCreate);
                var filter = new SearchFilter.IsEqualTo(MyNamedProp, numberToSearchFor);
                var results = folder.FindItems(filter, view);
            }

            sw.Stop();
            Console.WriteLine("SearchByFilter " + (sort ? "with sort" : "without sort") + " finished after: " + sw.ElapsedMilliseconds + " milliseconds.");
        }

        public static void SearchBySearchFolder(SearchFolder folder)
        {
            var rnd = new Random();
            var view = new ItemView(1, 0);
            view.OrderBy.Add(MyNamedProp, SortDirection.Ascending);
            var sw = new Stopwatch();
            Console.WriteLine("SearchBySearchFolder started.");
            sw.Start();
            for (var x = 0; x < NumberOfSearchesToPerform; x++)
            {
                var numberToSearchFor = rnd.Next(0, NumberOfMessagesToCreate);
                var filter = new SearchFilter.IsEqualTo(MyNamedProp, numberToSearchFor);
                var results = folder.FindItems(filter, view);
            }

            sw.Stop();
            Console.WriteLine("SearchBySearchFolder stopped after: " + sw.ElapsedMilliseconds + " milliseconds.");
        }

        public static void CreateLotsOfMessages(ExchangeService exchService, FolderId folderId)
        {
            for (var x = 0; x < NumberOfMessagesToCreate; x++)
            {
                var msg = new EmailMessage(exchService);
                msg.Subject = "SearchingMailboxViaEWS message " + x.ToString("D5");
                msg.Body = new TextBody { BodyType = BodyType.Text, Text = "Test." };
                msg.SetExtendedProperty(MyNamedProp, x);
                msg.Save(folderId);
                Console.WriteLine("Created message " + x + " of " + NumberOfMessagesToCreate);
            }
        }

        public static SearchFolder CreateSearchFolderIfNeeded(ExchangeService exchService, Mailbox mailbox)
        {
            // Check if it exists already
            var searchFoldersId = new FolderId(WellKnownFolderName.SearchFolders, mailbox);
            var searchFoldersFolder = Folder.Bind(exchService, searchFoldersId);
            var searchFoldersView = new FolderView(2);
            var searchFoldersFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, SearchFolderName);
            var searchFolderResults = searchFoldersFolder.FindFolders(searchFoldersFilter, searchFoldersView);
            SearchFolder mySearchFolder = null;
            if (searchFolderResults.Folders.Count > 1)
            {
                Console.WriteLine("The expected folder name is ambiguous. How did we end up with multiple folders?");
                Console.WriteLine("Dunno, but I'm going to delete all but the first.");
                for (var x = 1; x < searchFolderResults.Folders.Count; x++)
                {
                    searchFolderResults.Folders[x].Delete(DeleteMode.HardDelete);
                }
            }

            if (searchFolderResults.Folders.Count > 0)
            {
                Console.WriteLine("Found existing search folder.");
                mySearchFolder = searchFolderResults.Folders[0] as SearchFolder;
                if (mySearchFolder == null)
                {
                    Console.WriteLine("Somehow this folder isn't a search folder. Deleting it.");
                    searchFolderResults.Folders[0].Delete(DeleteMode.HardDelete);
                }
            }

            if (mySearchFolder == null)
            {
                Console.WriteLine("Creating a new search folder.");
                mySearchFolder = new SearchFolder(exchService);
                mySearchFolder.DisplayName = SearchFolderName;
                mySearchFolder.SearchParameters.SearchFilter =
                    new SearchFilter.Exists(MyNamedProp);
                var inboxId = new FolderId(WellKnownFolderName.Inbox, mailbox);
                mySearchFolder.SearchParameters.RootFolderIds.Add(inboxId);
                mySearchFolder.SearchParameters.Traversal = SearchFolderTraversal.Shallow;
                mySearchFolder.Save(searchFoldersId);
            }

            return mySearchFolder;
        }
    }
}
