using System.Text;
using CommandLine;

namespace FindingsDeduplicator
{
    class Options
    {
        [Option('i', "input", Required = true, HelpText = "Input file to read.")]
        public string InputFile { get; set; }

        [Option('v', null, HelpText = "Print details during execution.")]
        public bool Verbose { get; set; }

        [Option('n', DefaultValue = 2, HelpText = "Sheet number to de-duplicate.")]
        public int SheetNumber { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            var usage = new StringBuilder();
            usage.AppendLine("Findings Deduplicator");
            return usage.ToString();
        }
    }
}
