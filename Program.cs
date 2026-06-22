
using MarkdownDocxConverter.Converters;

internal class Program
{
    static void Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("  MarkdownDocxConverter input.md   output.docx");
            Console.WriteLine("  MarkdownDocxConverter input.docx output.md");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found: {inputPath}");
            return;
        }

        string inputExt = Path.GetExtension(inputPath).ToLowerInvariant();
        string outputExt = Path.GetExtension(outputPath).ToLowerInvariant();

        bool mdToDocx = inputExt == ".md" && outputExt == ".docx";
        bool docxToMd = inputExt == ".docx" && outputExt == ".md";

        if (mdToDocx)
        {
            Console.WriteLine($"Converting Markdown → DOCX ...");
            MarkdownToDocxConverter converter = new MarkdownToDocxConverter();
            converter.Convert(inputPath, outputPath);
            Console.WriteLine($"Done! Output: {outputPath}");
        }
        else if (docxToMd)
        {
            Console.WriteLine($"Converting DOCX → Markdown ...");
            DocxToMarkdownConverter converter = new DocxToMarkdownConverter();
            converter.Convert(inputPath, outputPath);
            Console.WriteLine($"Done! Output: {outputPath}");
        }
        else
        {
            Console.WriteLine("Error: Unsupported conversion.");
            Console.WriteLine("  Supported: .md → .docx   or   .docx → .md");
        }
    }
}
