using DocumentFormat.OpenXml.Wordprocessing;
using md2docx_resharp;
using System.Collections;
using System.Collections.Generic;
using Xunit;
using Xunit.Abstractions;

namespace md2docx_resharp_test {
	public class StyleFactoryTests {

		private readonly ITestOutputHelper output;

		public StyleFactoryTests(ITestOutputHelper output) {
			this.output = output;
		}

		[Theory]
		[ClassData(typeof(CSUStyleTestData))]
		public void CSUStyleTests(string json, Style[] expected) {
			RuleJsonSerializer ruleJsonSerializer = new RuleJsonSerializer();
			Rule[] rulesDictionary = ruleJsonSerializer.ParseJson(json);

			StyleFactory styleFactory = new StyleFactory();
			Style result = styleFactory.GenerateStyle(rulesDictionary);

			// output.WriteLine(expected.OuterXml);
			// output.WriteLine(result.OuterXml);

			Assert.Equal(result, expected);
		}
	}
	public class CSUStyleTestData : IEnumerable<object[]> {
		private readonly List<object[]> _data = new List<object[]> {
			new object[] {@"{
				""heading 1"": {
                  ""MarkdownBlock"": ""heading 1"",
                  ""EnglishFont"": ""Times New Roman"",
                  ""ChineseFont"": ""黑体"",
                  ""FontSize"": ""三号"",
                  ""Align"": ""居中"",
                  ""Outline"": true,
                  ""OutlineLevel"": 0,
                  ""Bold"": false,
                  ""Italic"": false,
                  ""Underline"": false,
                  ""Strike"": false,
                  ""Indents"": 0,
                  ""PageBreakBefore"": true,
                  ""BeforeAfterLine"": 1,
                  ""LineSpacingValues"": 0
                },
				""heading 2"": {
                  ""MarkdownBlock"": ""heading 2"",
                  ""EnglishFont"": ""Times New Roman"",
                  ""ChineseFont"": ""黑体"",
                  ""FontSize"": ""小四"",
                  ""Align"": ""左对齐"",
                  ""Outline"": true,
                  ""OutlineLevel"": 1,
                  ""Bold"": false,
                  ""Italic"": false,
                  ""Underline"": false,
                  ""Strike"": false,
                  ""Indents"": 2,
                  ""PageBreakBefore"": false,
                  ""BeforeAfterLine"": 0,
                  ""LineSpacingValues"": 0
                },
				""heading 3"": {
                  ""MarkdownBlock"": ""heading 3"",
                  ""EnglishFont"": ""Times New Roman"",
                  ""ChineseFont"": ""楷体"",
                  ""FontSize"": ""小四"",
                  ""Align"": ""左对齐"",
                  ""Outline"": true,
                  ""OutlineLevel"": 2,
                  ""Bold"": false,
                  ""Italic"": false,
                  ""Underline"": false,
                  ""Strike"": false,
                  ""Indents"": 2,
                  ""PageBreakBefore"": false,
                  ""BeforeAfterLine"": 0,
                  ""LineSpacingValues"": 0
                },
				""bodytext"": {
                  ""MarkdownBlock"": ""bodytext"",
                  ""EnglishFont"": ""Times New Roman"",
                  ""ChineseFont"": ""宋体"",
                  ""FontSize"": ""小四"",
                  ""Align"": ""左对齐"",
                  ""Outline"": false,
                  ""OutlineLevel"": 0,
                  ""Bold"": false,
                  ""Italic"": false,
                  ""Underline"": false,
                  ""Strike"": false,
                  ""Indents"": 2,
                  ""PageBreakBefore"": false,
                  ""BeforeAfterLine"": 0,
                  ""LineSpacingValues"": 1.5
                },
				""code"": {
                  ""MarkdownBlock"": ""code"",
                  ""EnglishFont"": ""Consolas"",
                  ""ChineseFont"": ""黑体"",
                  ""FontSize"": ""小四"",
                  ""Align"": ""左对齐"",
                  ""Outline"": false,
                  ""OutlineLevel"": 0,
                  ""Bold"": false,
                  ""Italic"": false,
                  ""Underline"": false,
                  ""Strike"": false,
                  ""Indents"": 0,
                  ""PageBreakBefore"": false,
                  ""BeforeAfterLine"": 0,
                  ""LineSpacingValues"": 0
                },
				""reference"": {
                  ""MarkdownBlock"": ""reference"",
                  ""EnglishFont"": ""Times New Roman"",
                  ""ChineseFont"": ""黑体"",
                  ""FontSize"": ""五号"",
                  ""Align"": ""左对齐"",
                  ""Outline"": false,
                  ""OutlineLevel"": 0,
                  ""Bold"": false,
                  ""Italic"": false,
                  ""Underline"": false,
                  ""Strike"": false,
                  ""Indents"": 0,
                  ""PageBreakBefore"": false,
                  ""BeforeAfterLine"": 0,
                  ""LineSpacingValues"": 1.5
                }
			}",
				new Style[] {
					new Style {
						Type = StyleValues.Paragraph,
						StyleId = "heading 1",
						StyleName = new StyleName {
							Val = "heading 1"
						},
						StyleParagraphProperties = new StyleParagraphProperties {
							PageBreakBefore = new PageBreakBefore(),
							SpacingBetweenLines = new SpacingBetweenLines() {
								BeforeLines = 100,
								AfterLines = 100
							},
							Justification = new Justification() {
								Val = JustificationValues.Center
							},
							OutlineLevel = new OutlineLevel() {
								Val = 0
							}
						},
						StyleRunProperties = new StyleRunProperties {
							RunFonts = new RunFonts() {
								Ascii = "Times New Roman",
								HighAnsi = "Times New Roman",
								EastAsia = "黑体",
								ComplexScript = "Times New Roman"
							},
							FontSize = new FontSize {
								Val = "32"
							},
							FontSizeComplexScript = new FontSizeComplexScript {
								Val = "32"
							}
						}
					},
					new Style {
						Type = StyleValues.Paragraph,
						StyleId = "heading 2",
						StyleName = new StyleName {
							Val = "heading 2"
						},
						StyleParagraphProperties = new StyleParagraphProperties {
							Justification = new Justification() {
								Val = JustificationValues.Left
							},
							OutlineLevel = new OutlineLevel() {
								Val = 1
							},
							Indentation = new Indentation() {
								FirstLineChars = 200
							},
						},
						StyleRunProperties = new StyleRunProperties {
							RunFonts = new RunFonts() {
								Ascii = "Times New Roman",
								HighAnsi = "Times New Roman",
								EastAsia = "黑体",
								ComplexScript = "Times New Roman"
							},
							FontSize = new FontSize {
								Val = "24"
							},
							FontSizeComplexScript = new FontSizeComplexScript {
								Val = "24"
							}
						}
					},
					new Style {
						Type = StyleValues.Paragraph,
						StyleId = "heading 3",
						StyleName = new StyleName {
							Val = "heading 3"
						},
						StyleParagraphProperties = new StyleParagraphProperties {
							Justification = new Justification() {
								Val = JustificationValues.Left
							},
							OutlineLevel = new OutlineLevel() {
								Val = 2
							},
							Indentation = new Indentation() {
								FirstLineChars = 200
							}
						},
						StyleRunProperties = new StyleRunProperties {
							RunFonts = new RunFonts() {
								Ascii = "Times New Roman",
								HighAnsi = "Times New Roman",
								EastAsia = "楷体",
								ComplexScript = "Times New Roman"
							},
							FontSize = new FontSize {
								Val = "24"
							},
							FontSizeComplexScript = new FontSizeComplexScript {
								Val = "24"
							}
						}
					},
					new Style {
						Type = StyleValues.Paragraph,
						StyleId = "bodytext",
						StyleName = new StyleName {
							Val = "bodytext"
						},
						StyleParagraphProperties = new StyleParagraphProperties {
							Justification = new Justification() {
								Val = JustificationValues.Left
							},
							Indentation = new Indentation() {
								FirstLineChars = 200
							},
							SpacingBetweenLines = new SpacingBetweenLines {
								Line = "360",
								LineRule = LineSpacingRuleValues.Auto
							}
						},
						StyleRunProperties = new StyleRunProperties {
							RunFonts = new RunFonts() {
								Ascii = "Times New Roman",
								HighAnsi = "Times New Roman",
								EastAsia = "宋体",
								ComplexScript = "Times New Roman"
							},
							FontSize = new FontSize {
								Val = "24"
							},
							FontSizeComplexScript = new FontSizeComplexScript {
								Val = "24"
							}
						}
					},
					new Style {
						Type = StyleValues.Paragraph,
						StyleId = "code",
						StyleName = new StyleName {
							Val = "code"
						},
						StyleParagraphProperties = new StyleParagraphProperties {
							Justification = new Justification() {
								Val = JustificationValues.Left
							},
						},
						StyleRunProperties = new StyleRunProperties {
							RunFonts = new RunFonts() {
								Ascii = "Consolas",
								HighAnsi = "Consolas",
								EastAsia = "黑体",
								ComplexScript = "Consolas"
							},
							FontSize = new FontSize {
								Val = "24"
							},
							FontSizeComplexScript = new FontSizeComplexScript {
								Val = "24"
							}
						}
					},
					new Style {
						Type = StyleValues.Paragraph,
						StyleId = "reference",
						StyleName = new StyleName {
							Val = "reference"
						},
						StyleParagraphProperties = new StyleParagraphProperties {
							Justification = new Justification() {
								Val = JustificationValues.Left
							},
							SpacingBetweenLines = new SpacingBetweenLines {
								Line = "360",
								LineRule = LineSpacingRuleValues.Auto
							}
						},
						StyleRunProperties = new StyleRunProperties {
							RunFonts = new RunFonts() {
								Ascii = "Times New Roman",
								HighAnsi = "Times New Roman",
								EastAsia = "黑体",
								ComplexScript = "Times New Roman"
							},
							FontSize = new FontSize {
								Val = "21"
							},
							FontSizeComplexScript = new FontSizeComplexScript {
								Val = "21"
							}
						}
					}
				}
			}
		};

		public IEnumerator<object[]> GetEnumerator() {
			return _data.GetEnumerator();
		}

		IEnumerator IEnumerable.GetEnumerator() {
			return GetEnumerator();
		}
	}
}