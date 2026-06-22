using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;


namespace MarkdownDocxConverter.Converters
{
    
    // Converts a DOCX file to a Markdown file.
    // végigmegy az OpenXml Body elemein és összerakja markdown-ba
    
    public class DocxToMarkdownConverter
    {
        public void Convert(string inputPath, string outputPath)
        {
            StringBuilder markdown = new StringBuilder();

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(inputPath, false))
            {
                MainDocumentPart? mainPart = wordDoc.MainDocumentPart;

                if (mainPart == null || mainPart.Document.Body == null)
                {
                    Console.WriteLine("Error: The DOCX file has no readable content.");
                    return;
                }

                Body body = mainPart.Document.Body;

                // lista detektáció
                string? previousListStyle = null;
                int orderedCounter = 1;

                foreach (var element in body.Elements())
                {
                    if (element is Paragraph para)
                    {
                        string paraText = ConvertParagraph(
                            para,
                            ref previousListStyle,
                            ref orderedCounter
                        );

                        if (paraText.Length > 0)
                        {
                            markdown.AppendLine(paraText);
                        }
                        else
                        {
                            // üres sorok kezelése bekezdések között
                            markdown.AppendLine();
                        }
                    }
                    else if (element is Table table)
                    {
                        string tableText = ConvertTable(table);
                        markdown.AppendLine(tableText);
                    }
                }
            }

            File.WriteAllText(outputPath, markdown.ToString().TrimEnd() + Environment.NewLine);
        }

        //paragrafus kezelés

        private string ConvertParagraph(
            Paragraph para,
            ref string? previousListStyle,
            ref int orderedCounter)
        {
            // paragrafus stílus kezelés, árolsá
            string styleId = GetStyleId(para);

            // lista e a paragrafus, része e egy listának
            NumberingProperties? numPr = para.ParagraphProperties?.NumberingProperties;
            bool isList = numPr != null;

            if (isList)
            {
                return ConvertListItem(para, numPr!, styleId, ref previousListStyle, ref orderedCounter);
            }

            // ha nem visszaálítjuk a lista kezelést
            previousListStyle = null;
            orderedCounter = 1;

            string inlineText = GetInlineText(para);

            // paragrafus/részegység határai (border/margin)
            ParagraphBorders? borders = para.ParagraphProperties?.ParagraphBorders;
            if (borders != null && borders.BottomBorder != null && inlineText.Trim() == "")
            {
                return "---";
            }

            if (inlineText.Trim() == "")
            {
                return "";
            }

            // markdown prefix
            if (styleId.StartsWith("Heading"))
            {
                if (int.TryParse(styleId.Replace("Heading", ""), out int level))
                {
                    string prefix = new string('#', level);
                    return $"{prefix} {inlineText}";
                }
            }

            if (styleId == "CodeBlock" || styleId == "Code")
            {
                return $"```\n{inlineText}\n```";
            }

            // default üres paragrafus
            return inlineText;
        }

        private string ConvertListItem(
            Paragraph para,
            NumberingProperties numPr,
            string styleId,
            ref string? previousListStyle,
            ref int orderedCounter)
        {
            int depth = numPr.NumberingLevelReference?.Val ?? 0;
            string indent = new string(' ', depth * 2);

            string inlineText = GetInlineText(para);

            // rendezett/nem rendezett lista kezelés
            int numbId = numPr.NumberingId?.Val ?? 0;

            // ID=2 rendezett listához
            bool isOrdered = numbId == 2;

            // ha a lista stílusa változik visszaállítás alapértelmezettre
            if (previousListStyle != styleId)
            {
                orderedCounter = 1;
            }
            previousListStyle = styleId;

            if (isOrdered)
            {
                string item = $"{indent}{orderedCounter}. {inlineText}";
                orderedCounter++;
                return item;
            }
            else
            {
                return $"{indent}- {inlineText}";
            }
        }

        // GFM markdown tábla

        private string ConvertTable(Table table)
        {
            StringBuilder sb = new StringBuilder();

            List<List<string>> rows = new List<List<string>>();

            foreach (var row in table.Elements<TableRow>())
            {
                List<string> cells = new List<string>();
                foreach (var cell in row.Elements<TableCell>())
                {
                    // alkotóelemek kezelése space-el elválasztva
                    StringBuilder cellText = new StringBuilder();
                    foreach (var para in cell.Elements<Paragraph>())
                    {
                        cellText.Append(GetInlineText(para));
                    }
                    cells.Add(cellText.ToString().Trim());
                }
                rows.Add(cells);
            }

            if (rows.Count == 0)
            {
                return "";
            }

            // Header sor/line
            sb.AppendLine("| " + string.Join(" | ", rows[0]) + " |");

            // Separator sor/line
            sb.AppendLine("| " + string.Join(" | ", rows[0].Select(_ => "---")) + " |");

            // Data sor/line
            for (int i = 1; i < rows.Count; i++)
            {
                sb.AppendLine("| " + string.Join(" | ", rows[i]) + " |");
            }

            return sb.ToString().TrimEnd();
        }

        // text kibontása

        private string GetInlineText(Paragraph para)
        {
            StringBuilder sb = new StringBuilder();

            foreach (var child in para.Elements())
            {
                if (child is Run run)
                {
                    string runText = GetRunText(run);
                    if (runText.Length == 0)
                    {
                        continue;
                    }

                    RunProperties? rPr = run.RunProperties;

                    bool bold = rPr?.Bold != null;
                    bool italic = rPr?.Italic != null;
                    bool code = IsMonospace(rPr);

                    string formatted = runText;

                    if (code)
                    {
                        formatted = $"`{formatted}`";
                    }
                    else
                    {
                        if (bold && italic)
                        {
                            formatted = $"***{formatted}***";
                        }
                        else if (bold)
                        {
                            formatted = $"**{formatted}**";
                        }
                        else if (italic)
                        {
                            formatted = $"*{formatted}*";
                        }
                    }

                    sb.Append(formatted);
                }
                else if (child is Hyperlink hyperlink)
                {
                    // linket kibontása
                    StringBuilder linkText = new StringBuilder();
                    foreach (Run run2 in hyperlink.Elements<Run>())
                    {
                        linkText.Append(GetRunText(run2));
                    }
                    sb.Append(linkText.ToString());
                }
            }

            return sb.ToString();
        }

        private string GetRunText(Run run)
        {
            StringBuilder sb = new StringBuilder();
            foreach (var child in run.Elements())
            {
                if (child is Text t)
                {
                    sb.Append(t.Text);
                }
                else if (child is Break)
                {
                    sb.Append("  \n"); // Markdown sortörés
                }
            }
            return sb.ToString();
        }

        private bool IsMonospace(RunProperties? rPr)
        {
            if (rPr == null)
            {
                return false;
            }

            string? font = rPr.RunFonts?.Ascii;
            if (font == null)
            {
                return false;
            }

            return font.Contains("Courier") || font.Contains("Mono") || font.Contains("Consolas");
        }

        // Helpers

        private string GetStyleId(Paragraph para)
        {
            return para.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "Normal";
        }
    }

}
