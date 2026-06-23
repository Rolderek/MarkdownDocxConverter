első indítás:
dotnet build

Markdown -> docx:
dotnet run -- sample.md output.docx

docx -> Markdown:
dotnet run -- output.docx roundtrip.md

például:
MarkdownDocxConverter.exe sample.md output.docx
MarkdownDocxConverter.exe output.docx roundtrip.md
