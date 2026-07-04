using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Math;

class Program {
    static void Main() {
        var doc = WordprocessingDocument.Create("test.docx", WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(
            new Paragraph(
                new OfficeMath(
                    new Fraction(
                        new FractionProperties(new FractionType() { Val = FractionTypeValues.Bar }),
                        new Numerator(new Run(new Text("1"))),
                        new Denominator(new Run(new Text("2")))
                    )
                )
            )
        ));
        doc.Close();
    }
}
