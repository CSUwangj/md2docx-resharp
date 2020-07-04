using System;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using Mono.Options;

namespace md2docx_resharp
{
    class RunArgs {
        public string MarkdonwPath { get; set; }
        public string DocxPath { get; set; }
        public string RulesPath { get; set; }
        public string ConfigPath { get; set; }
    }
    class Program
    {
        /// <summary>
        /// Print usage
        /// </summary>
        /// <param name="options">options</param>
        private static void Usage(OptionSet options)
        {
            Console.WriteLine(@"Usage md2docx [OPTIONS]
Convert markdown to docx with specified options.
Opntions:");
            options.WriteOptionDescriptions(Console.Out);
        }

        /// <summary>
        /// parse command args return a object
        /// </summary>
        /// <param name="args">command line args</param>
        /// <returns>RunArgs</returns>
        private static RunArgs ParseArgs(string[] args) {
            RunArgs runArgs = new RunArgs();
            bool showHelp = false;

            OptionSet p = new OptionSet
            {
                {
                    "i|input=", "{INPUT} markdown file path.",
                    v => runArgs.MarkdonwPath = v
                },
                {
                    "o|output=", "{OUTPUT} docx file path.",
                    v => runArgs.DocxPath = v
                },
                {
                    "r|rules=", "{RULES} for pure text block",
                    v => runArgs.RulesPath = v
                },
                {
                    "c|config=", "{CONFIG} file for coverting",
                    v => runArgs.ConfigPath = v
                },
                {
                    "h|help", "show this message and exit",
                    v => showHelp = v != null
                },
            };

            try {
                p.Parse(args);
            } catch (OptionException e) {
                Console.Write("md2docx: ");
                Console.WriteLine(e.Message);
                Console.WriteLine("Try`md2docx --help' for more information.");
                Environment.Exit(1);
            }

            if (showHelp) {
                Usage(p);
                Environment.Exit(0);
            }

            if (runArgs.GetType().GetProperties()
                .Any(p => string.IsNullOrWhiteSpace((p.GetValue(runArgs) as string)))) {
                Usage(p);
                Environment.Exit(1);
            }
            return runArgs;
        }
        static void Main(string[] args)
        {
            RunArgs runArgs = ParseArgs(args);
        }
    }
}
