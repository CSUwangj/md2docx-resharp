using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace md2docx_resharp
{

    public class Rules
    {
        public Dictionary<string, Rule> RulePairs { get; set; }
    }

    public class Rule
    {
        public string MarkdownBlock { get; set; }
        public string ChineseFont { get; set; }
        public string EnglishFont { get; set; }
        public string FontSize { get; set; }
        public string Align { get; set; }
        public double Indents { get; set; }
        public bool PageBreakBefore { get; set; }
        public double BeforeAfterLine { get; set; }
        public double LineSpacingValues { get; set; }
        public bool Outline { get; set; }
        public double OutlineLevel { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strike { get; set; }
    }

    public class RuleJsonSerializer
    {
        public Rule[] ParseJson(string rules)
        {
            var options = new JsonSerializerOptions();
            Dictionary<string, Rule> rulesDictionary = JsonSerializer.Deserialize<Dictionary<string, Rule>>(rules, options);
            return rulesDictionary.Values.ToArray;
        }
    }
    public class StyleFactory
    {
        public Style GenerateStyle(Rule[] rules)
        {
            Style style = new Style { };
            return style;
        }
    }
}
