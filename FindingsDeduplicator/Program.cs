using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace FindingsDeduplicator
{
    class Program
    {
        static void Main(string[] args)
        {
            var options = new Options();
            if (CommandLine.Parser.Default.ParseArguments(args, options))
            {
                if (File.Exists(options.InputFile))
                {
                    // Open the file
                    var file = new FileInfo(options.InputFile);
                    using (var package = new ExcelPackage(file))
                    {
                        // Look for the findings sheet
                        var worksheet = package.Workbook.Worksheets[options.SheetNumber];
                        List<int> idsToDelete = new List<int>();

                        foreach (var findingsForIp in worksheet.Cells["a:r"]
                            .Where(r => new List<string> { "Microsoft Windows Patching Issues", "Third Party Software Patching Issues" }.Contains(r.Offset(0,1).Value))
                            .GroupBy(r => r.Offset(0, 4).Value))
                        {
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("[*] Finding duplicates for {0}", findingsForIp.Key);

                            foreach(var findings in findingsForIp
                                .Where(r => r.Offset(0,6).Value != null)
                                .GroupBy(r => CreateCustomFinding((string)r.Offset(0,6).Value)))
                            {
                                var findingCount = findings.Count();
                                if (findingCount > 1)
                                {
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    Console.WriteLine("Found {0} findings for {1}", findingCount, findings.Key);

                                    // Add all but the first example to the list to be deleted
                                    idsToDelete.AddRange(findings.Skip(1).Select(i => i.Start.Row));
                                }
                            }
                        }

                        // Due to the rows shifting when deleted, start at the bottom and go up to delete
                        var orderedIdsToDelete = idsToDelete.OrderByDescending(r => r);
                        foreach (var id in orderedIdsToDelete)
                        {
                            worksheet.DeleteRow(id, 1);
                        }

                        var outFile = new FileStream("output.xlsx", FileMode.Create);
                        package.SaveAs(outFile);

                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("Done");
                        Console.ReadLine();
                    }
                }
            }
        }


        /// <summary>
        /// Splits the finding name based on a colon, in case there are multiple CVEs for a single patch
        /// </summary>
        /// <param name="text">The text.</param>
        /// <returns></returns>
        private static string CreateCustomFinding(string text)
        {
            if (!text.Contains(':'))
            {
                return text;
            }

            return text.Substring(0, text.IndexOf(':'));
        }
    }
}
