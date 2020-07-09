using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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

            OptionSet p = new OptionSet {
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

            // var markdown = Markdown.Parse(runArgs.MarkdonwPath);
            RuleJsonSerializer ruleJsonSerializer = new RuleJsonSerializer();
            var rules = ruleJsonSerializer.ParseJson(System.IO.File.ReadAllText(runArgs.ConfigPath));

            using WordprocessingDocument document = WordprocessingDocument.Create(runArgs.DocxPath, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = document.AddMainDocumentPart();
            GenerateMainPart(mainPart, runArgs.MarkdonwPath);
            StyleDefinitionsPart styleDefinitionsPart = mainPart.AddNewPart<StyleDefinitionsPart>("Styles");
            // TODO: latent config if needed
            GenerateStyleDefinitionsPartContent(styleDefinitionsPart, rules, true);

            FontTablePart fontTablePart1 = mainPart.AddNewPart<FontTablePart>("FontTable");
            GeneratedCode.GenerateFontTablePartContent(fontTablePart1);

            SetPackageProperties(document);
        }

        /// <summary>
        /// Generate document body
        /// </summary>
        /// <param name="mainPart">main body</param>
        /// <param name="md">markdown document</param>
        private static void GenerateMainPart(MainDocumentPart mainPart, string md) {
            // TODO: fill function
            Document document1 = new Document() { MCAttributes = new MarkupCompatibilityAttributes() };

            Body docBody = new Body();

            SectionProperties sectionProperties1 = new SectionProperties();
            PageSize pageSize1 = new PageSize() { Width = 11906U, Height = 16838U };
            PageMargin pageMargin1 = new PageMargin() { Top = 1418, Right = 1134U, Bottom = 1418, Left = 1701U, Header = 851U, Footer = 992U, Gutter = 0U };
            Columns columns1 = new Columns() { Space = "425" };
            DocGrid docGrid1 = new DocGrid() { Type = DocGridValues.Lines, LinePitch = 312 };

            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);
            docBody.Append(sectionProperties1);
            document1.Append(docBody);

            mainPart.Document = document1;
        }

        /// <summary>
        /// Generate styles from config json object
        /// </summary>
        /// <param name="styleDefinitionsPart1">Styles object</param>
        /// <param name="rules">Config json object</param>
        /// <param name="latent">If user need latent style</param>
        private static void GenerateStyleDefinitionsPartContent(StyleDefinitionsPart styleDefinitionsPart1, List<Rule> rules, bool latent) {
            Styles styles = new Styles() { MCAttributes = new MarkupCompatibilityAttributes() };

            DocDefaults docDefaults = new DocDefaults {
                RunPropertiesDefault = new RunPropertiesDefault {
                    RunPropertiesBaseStyle = new RunPropertiesBaseStyle {
                        RunFonts = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "宋体", ComplexScript = "Times New Roman" },
                        Kern = new Kern { Val = 2U },
                        Languages = new Languages { Val = "en-US", EastAsia = "zh-CN", Bidi = "ar-SA" },
                        FontSize = new FontSize { Val = "24" },
                        FontSizeComplexScript = new FontSizeComplexScript { Val = "24" }
                    }
                },
                ParagraphPropertiesDefault = new ParagraphPropertiesDefault()
            };

            styles.Append(docDefaults);

            if (latent) {
                styles.Append(GeneratedCode.GenerateLatentStyles());
            }

            StyleFactory styleFactory = new StyleFactory();
            var result = styleFactory.GenerateStyles(rules);
            foreach (Style style in result) {
                styles.Append(style);
            }

            styleDefinitionsPart1.Styles = styles;
        }

        /// <summary>
        /// Set document's properties like title, creator, etc.
        /// </summary>
        /// <param name="document">Document file</param>
        private static void SetPackageProperties(OpenXmlPackage document) {
            document.PackageProperties.Creator = "";
            document.PackageProperties.Title = "";
            document.PackageProperties.Revision = "3";
            document.PackageProperties.Created = DateTime.Now;
            document.PackageProperties.Modified = DateTime.Now;
            document.PackageProperties.LastModifiedBy = "md2docx_by_CSUwangj";
        }
    }
}
