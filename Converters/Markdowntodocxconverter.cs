using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
//using-okat majd átnézni és ami fölösleges kiszedni!


namespace MarkdownDocxConverter.Converters
{
    //konvertálás
    //Markdig használata és ezzel felépíti az új doksit
    public class MarkdownToDocxConverter
    {
        private const int BulletNumberingId = 1;
        private const int OrderedNumberingId = 2;

        public void Convert(string inputPath, string outputPath)
        {
            string markdownText = File.ReadAllText(inputPath);

            //pipeline AST-tá/ba
            MarkdownPipeline pipeline = new MarkdownPipelineBuilder()
                .UseAdvancedExtensions()
                .Build();

            MarkdownDocument markdownDoc = Markdown.Parse(markdownText, pipeline);

            // DOCX létrehozása
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(
                       outputPath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();

                AddNumberingDefinitions(mainPart);

                // stílus és egyéb
                AddStyles(mainPart);

                Body body = new Body();
                mainPart.Document = new Document(body);

                // top-level block átnázáse AST-ba
                foreach (Block block in markdownDoc)
                {
                    List<Paragraph> paragraphs = ConvertBlock(block);
                    foreach (Paragraph para in paragraphs)
                    {
                        body.AppendChild(para);
                    }
                }

                // fix body végződés, kell neki
                body.AppendChild(new SectionProperties());

                mainPart.Document.Save();
            }
        }

        //blokk átalakítása és konvenciója

        private List<Paragraph> ConvertBlock(Block block)
        {
            List<Paragraph> result = new List<Paragraph>();

            if (block is HeadingBlock heading)
            {
                result.Add(ConvertHeading(heading));
            }
            else if (block is ParagraphBlock para)
            {
                result.Add(ConvertParagraph(para));
            }
            else if (block is FencedCodeBlock code)
            {
                result.Add(ConvertCodeBlock(code));
            }
            else if (block is ListBlock list)
            {
                List<Paragraph> listParas = ConvertList(list, 0);
                result.AddRange(listParas);
            }
            else if (block is QuoteBlock quote)
            {
                foreach (Block inner in quote)
                {
                    result.AddRange(ConvertBlock(inner));
                }
            }
            else if (block is ThematicBreakBlock)
            {
                result.Add(ConvertThematicBreak());
            }

            return result;
        }

        private Paragraph ConvertHeading(HeadingBlock heading)
        {
            string styleId = $"Heading{heading.Level}";

            Paragraph para = new Paragraph();
            ParagraphProperties pPr = new ParagraphProperties(
                new ParagraphStyleId() { Val = styleId }
            );
            para.AppendChild(pPr);

            if (heading.Inline != null)
            {
                foreach (Run run in ConvertInlines(heading.Inline))
                {
                    para.AppendChild(run);
                }
            }

            return para;
        }

        private Paragraph ConvertParagraph(ParagraphBlock paraBlock)
        {
            Paragraph para = new Paragraph();

            if (paraBlock.Inline != null)
            {
                foreach (Run run in ConvertInlines(paraBlock.Inline))
                {
                    para.AppendChild(run);
                }
            }

            return para;
        }

        private Paragraph ConvertCodeBlock(FencedCodeBlock code)
        {
            // összekapcsolás
            string codeText = string.Join("\n", code.Lines.Lines
                .Take(code.Lines.Count)
                .Select(line => line.ToString()));

            Paragraph para = new Paragraph();
            ParagraphProperties pPr = new ParagraphProperties(
                new ParagraphStyleId() { Val = "CodeBlock" }
            );
            para.AppendChild(pPr);

            Run run = new Run();
            RunProperties rPr = new RunProperties(
                new RunFonts() { Ascii = "Courier New", HighAnsi = "Courier New" },
                new FontSize() { Val = "20" }
            );
            run.AppendChild(rPr);
            run.AppendChild(new Text(codeText) { Space = SpaceProcessingModeValues.Preserve });
            para.AppendChild(run);

            return para;
        }

        private List<Paragraph> ConvertList(ListBlock list, int depth)
        {
            List<Paragraph> result = new List<Paragraph>();
            bool isOrdered = list.IsOrdered;
            int numbId = isOrdered ? OrderedNumberingId : BulletNumberingId;

            foreach (Block item in list)
            {
                if (item is ListItemBlock listItem)
                {
                    foreach (Block inner in listItem)
                    {
                        if (inner is ParagraphBlock paraBlock)
                        {
                            Paragraph para = new Paragraph();

                            ParagraphProperties pPr = new ParagraphProperties();
                            NumberingProperties numPr = new NumberingProperties(
                                new NumberingLevelReference() { Val = depth },
                                new NumberingId() { Val = numbId }
                            );
                            pPr.AppendChild(numPr);
                            para.AppendChild(pPr);

                            if (paraBlock.Inline != null)
                            {
                                foreach (Run run in ConvertInlines(paraBlock.Inline))
                                {
                                    para.AppendChild(run);
                                }
                            }

                            result.Add(para);
                        }
                        else if (inner is ListBlock nestedList)
                        {
                            result.AddRange(ConvertList(nestedList, depth + 1));
                        }
                    }
                }
            }

            return result;
        }

        private Paragraph ConvertThematicBreak()
        {
            Paragraph para = new Paragraph();
            ParagraphProperties pPr = new ParagraphProperties();

            // alsó része és elrendezés
            ParagraphBorders borders = new ParagraphBorders(
                new BottomBorder()
                {
                    Val = BorderValues.Single,
                    Size = 6,
                    Space = 1,
                    Color = "000000"
                }
            );
            pPr.AppendChild(borders);
            para.AppendChild(pPr);

            return para;
        }

        //belső sor átalakítása

        private List<Run> ConvertInlines(ContainerInline inlines)
        {
            List<Run> runs = new List<Run>();

            foreach (Inline inline in inlines)
            {
                runs.AddRange(ConvertInline(inline, false, false, false));
            }

            return runs;
        }

        private List<Run> ConvertInline(Inline inline, bool bold, bool italic, bool code)
        {
            List<Run> runs = new List<Run>();

            if (inline is LiteralInline literal)
            {
                Run run = BuildRun(literal.Content.ToString(), bold, italic, code);
                runs.Add(run);
            }
            else if (inline is EmphasisInline emphasis)
            {
                bool newBold = bold || emphasis.DelimiterCount == 2;
                bool newItalic = italic || emphasis.DelimiterCount == 1;

                foreach (Inline child in emphasis)
                {
                    runs.AddRange(ConvertInline(child, newBold, newItalic, code));
                }
            }
            else if (inline is CodeInline codeInline)
            {
                Run run = BuildRun(codeInline.Content, bold, italic, true);
                runs.Add(run);
            }
            else if (inline is LinkInline link)
            {
                // segéd függvény
                foreach (Inline child in link)
                {
                    runs.AddRange(ConvertInline(child, bold, italic, code));
                }
            }
            else if (inline is LineBreakInline)
            {
                Run run = new Run();
                run.AppendChild(new Break());
                runs.Add(run);
            }
            else if (inline is ContainerInline container)
            {
                foreach (Inline child in container)
                {
                    runs.AddRange(ConvertInline(child, bold, italic, code));
                }
            }

            return runs;
        }

        private Run BuildRun(string text, bool bold, bool italic, bool code)
        {
            Run run = new Run();
            RunProperties rPr = new RunProperties();

            if (bold)
            {
                rPr.AppendChild(new Bold());
            }
            if (italic)
            {
                rPr.AppendChild(new Italic());
            }
            if (code)
            {
                rPr.AppendChild(new RunFonts() { Ascii = "Courier New", HighAnsi = "Courier New" });
                rPr.AppendChild(new FontSize() { Val = "20" });
            }

            if (rPr.HasChildren)
            {
                run.AppendChild(rPr);
            }

            run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

            return run;
        }

        //Doksi setup helper

        private void AddNumberingDefinitions(MainDocumentPart mainPart)
        {
            NumberingDefinitionsPart numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();

            Numbering numbering = new Numbering();
            //abstract bullet
            AbstractNum bulletAbstract = new AbstractNum() { AbstractNumberId = BulletNumberingId };
            bulletAbstract.AppendChild(new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel });

            for (int i = 0; i < 3; i++)
            {
                Level level = new Level() { LevelIndex = i };
                level.AppendChild(new NumberingFormat() { Val = NumberFormatValues.Bullet });
                level.AppendChild(new LevelText() { Val = "•" });
                level.AppendChild(new ParagraphProperties(
                    new Indentation() { Left = (720 + i * 360).ToString(), Hanging = "360" }
                ));
                bulletAbstract.AppendChild(level);
            }

            // Abstract ordered
            AbstractNum orderedAbstract = new AbstractNum() { AbstractNumberId = OrderedNumberingId };
            orderedAbstract.AppendChild(new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel });

            for (int i = 0; i < 3; i++)
            {
                Level level = new Level() { LevelIndex = i };
                level.AppendChild(new StartNumberingValue() { Val = 1 });
                level.AppendChild(new NumberingFormat() { Val = NumberFormatValues.Decimal });
                level.AppendChild(new LevelText() { Val = $"%{i + 1}." });
                level.AppendChild(new ParagraphProperties(
                    new Indentation() { Left = (720 + i * 360).ToString(), Hanging = "360" }
                ));
                orderedAbstract.AppendChild(level);
            }

            numbering.AppendChild(bulletAbstract);
            numbering.AppendChild(orderedAbstract);

            NumberingInstance bulletInstance = new NumberingInstance() { NumberID = BulletNumberingId };
            bulletInstance.AppendChild(new AbstractNumId() { Val = BulletNumberingId });
            numbering.AppendChild(bulletInstance);

            NumberingInstance orderedInstance = new NumberingInstance() { NumberID = OrderedNumberingId };
            orderedInstance.AppendChild(new AbstractNumId() { Val = OrderedNumberingId });
            numbering.AppendChild(orderedInstance);

            numberingPart.Numbering = numbering;
        }

        private void AddStyles(MainDocumentPart mainPart)
        {
            StyleDefinitionsPart stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            Styles styles = new Styles();

            // alapértelmezett betűtípus
            styles.AppendChild(new DocDefaults(
                new RunPropertiesDefault(
                    new RunPropertiesBaseStyle(
                        new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" },
                        new FontSize() { Val = "24" }  // 12pt
                    )
                )
            ));

            // címsorok H1–H6
            string[] headingNames = { "Heading 1", "Heading 2", "Heading 3",
                                      "Heading 4", "Heading 5", "Heading 6" };
            int[] headingSizes = { 36, 32, 28, 26, 24, 24 };

            for (int i = 0; i < 6; i++)
            {
                int level = i + 1;
                Style style = new Style()
                {
                    Type = StyleValues.Paragraph,
                    StyleId = $"Heading{level}",
                };
                style.AppendChild(new StyleName() { Val = headingNames[i] });
                style.AppendChild(new BasedOn() { Val = "Normal" });
                style.AppendChild(new StyleRunProperties(
                    new Bold(),
                    new FontSize() { Val = headingSizes[i].ToString() }
                ));
                style.AppendChild(new StyleParagraphProperties(
                    new SpacingBetweenLines() { Before = "240", After = "120" },
                    new OutlineLevel() { Val = i }
                ));
                styles.AppendChild(style);
            }

            // kód blokk stílusok
            Style codeStyle = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = "CodeBlock",
            };
            codeStyle.AppendChild(new StyleName() { Val = "Code Block" });
            codeStyle.AppendChild(new StyleRunProperties(
                new RunFonts() { Ascii = "Courier New", HighAnsi = "Courier New" },
                new FontSize() { Val = "20" }
            ));
            codeStyle.AppendChild(new StyleParagraphProperties(
                new SpacingBetweenLines() { Before = "120", After = "120" },
                new Shading() { Val = ShadingPatternValues.Clear, Fill = "F2F2F2" }
            ));
            styles.AppendChild(codeStyle);

            stylesPart.Styles = styles;
        }
    }

}
