using System;
using Mono.Options;

namespace md2docx_resharp
{
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

        static void Main(string[] args)
        {
            string markdonwPath = "";
            string docxPath = "";
            string rulesPath = "";
            string configPath = "";
            bool showHelp = false;
            OptionSet p = new OptionSet
            {
                {
                    "i|input=", "{INPUT} markdown file path.",
                    v => markdonwPath = v
                },
                {
                    "o|output=", "{OUTPUT} docx file path.",
                    v => docxPath = v
                },
                {
                    "r|rules=", "{RULES} for pure text block",
                    v => rulesPath = v
                },
                {
                    "c|config=", "{CONFIG} file for coverting",
                    v => configPath = v
                },
                {   
                    "h|help", "show this message and exit",
                    v => showHelp = v != null 
                },
            };

            try
            {
                p.Parse(args);
            }
            catch (OptionException e)
            {
                Console.Write("md2docx: ");
                Console.WriteLine(e.Message);
                Console.WriteLine("Try`md2docx --help' for more information.");
                return;
            }

            if (showHelp || markdonwPath == "" || configPath == "" || rulesPath == "" || docxPath == "")
            {
                Usage(p);
                return;
            }

        }
    }
}
