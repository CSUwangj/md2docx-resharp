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
        public int OutlineLevel { get; set; }
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
            return rulesDictionary.Select(item => item.Value).ToArray();
        }
    }
    public class StyleFactory
    {
        public Style[] GenerateStyles(Rule[] rules)
        {
            List<Style> styles = new List<Style>();
            foreach (Rule rule in rules) {
                styles.Add(GenerateStyle(rule));
            }
            return styles.ToArray();
        }

        private Style GenerateStyle(Rule rule) {
            string size;
            if (int.TryParse(rule.FontSize, out int sz)) {
                size = sz.ToString();
            } else {
                size = fontmap[rule.FontSize];
            }
            Style style = new Style {
                Type = StyleValues.Paragraph,
                StyleId = rule.MarkdownBlock,
                StyleName = new StyleName { Val = rule.MarkdownBlock },
                StyleParagraphProperties = new StyleParagraphProperties {
                    Justification = new Justification { Val = justmap[rule.Align] }
                },
                StyleRunProperties = new StyleRunProperties {
                    RunFonts = new RunFonts { Ascii = rule.EnglishFont, HighAnsi = rule.EnglishFont, ComplexScript = rule.EnglishFont, EastAsia = rule.ChineseFont },
                    FontSize = new FontSize { Val = size },
                    FontSizeComplexScript = new FontSizeComplexScript { Val = size }
                }
            };

            if (rule.Outline) {
                style.StyleParagraphProperties.OutlineLevel = new OutlineLevel { Val = rule.OutlineLevel };
            }
            if (rule.Bold) {
                style.StyleRunProperties.Bold = new Bold();
                style.StyleRunProperties.BoldComplexScript = new BoldComplexScript();
            }
            if (rule.Italic) {
                style.StyleRunProperties.Italic = new Italic();
                style.StyleRunProperties.ItalicComplexScript = new ItalicComplexScript();
            }
            if (rule.Underline) {
                style.StyleRunProperties.Underline = new Underline();
            }
            if (rule.Strike) {
                style.StyleRunProperties.Strike = new Strike();
            }
            if (rule.PageBreakBefore) {
                style.StyleParagraphProperties.PageBreakBefore = new PageBreakBefore();
            }

            if (rule.Indents != 0f) {
                style.StyleParagraphProperties.Indentation = new Indentation {
                    FirstLineChars = (int)rule.Indents * 100
                };
            }
            if (rule.BeforeAfterLine != 0f) {
                style.StyleParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines {
                    BeforeLines = (int)(rule.BeforeAfterLine * 100),
                    AfterLines = (int)(rule.BeforeAfterLine * 100)
                };
            }
            if (rule.LineSpacingValues != 0f) {
                style.StyleParagraphProperties.SpacingBetweenLines = style.StyleParagraphProperties.SpacingBetweenLines ?? new SpacingBetweenLines();
                style.StyleParagraphProperties.SpacingBetweenLines.Line = ((int)(rule.LineSpacingValues * 240)).ToString();
                style.StyleParagraphProperties.SpacingBetweenLines.LineRule = LineSpacingRuleValues.Auto;
            }
            return style;
        }
        #region Chinese font mapping
        static private readonly Dictionary<string, string> fontmap = new Dictionary<string, string>
        {
            {"初号", "84"},
            {"小初", "72"},
            {"一号", "52"},
            {"小一", "48"},
            {"二号", "44"},
            {"小二", "36"},
            {"三号", "32"},
            {"小三", "30"},
            {"四号", "28"},
            {"小四", "24"},
            {"五号", "21"},
            {"小五", "18"},
            {"六号", "15"},
            {"小六", "13"},
            {"七号", "11"},
            {"八号", "10"}
        };
        #endregion
        #region justification mapping
        static private readonly Dictionary<string, JustificationValues> justmap = new Dictionary<string, JustificationValues>
        {
            { "左对齐", JustificationValues.Left },
            { "居中", JustificationValues.Center },
            { "右对齐", JustificationValues.Right },
            { "分散对齐", JustificationValues.Distribute },
            { "两端对齐", JustificationValues.Both }
        };
        #endregion
    }
}
