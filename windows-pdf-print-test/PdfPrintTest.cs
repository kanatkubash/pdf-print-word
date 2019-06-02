using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using windows_pdf_print;

namespace windows_pdf_print_test
{
  [TestClass]
  public class PdfPrintTest
  {
    [TestMethod]
    public void TestInstantiatesCorrectly()
    {
      using (var pdfPrint = new PdfPrint())
      {
        Assert.IsInstanceOfType(pdfPrint, typeof(PdfPrint));
      }
    }

    [TestMethod]
    public void TestPrintsPdfFromDocx()
    {
      var dir = Directory.GetCurrentDirectory();
      var output = $"{dir}\\sample.pdf";

      if (File.Exists(output))
        File.Delete(output);

      using (var pdfPrint = new PdfPrint())
      {
        pdfPrint.Print($"{dir}\\sample.docx", output);
        Assert.IsTrue(File.Exists(output));
      }
    }

    [TestMethod, ExpectedException(typeof(FileNotFoundException))]
    public void TestThrowsIfFileNotFound()
    {
      using (var pdfPrint = new PdfPrint())
      {
        pdfPrint.Print("C:\\nonexisting-file.nonexistent-extension", "out.pdf");
      }
    }

    [TestMethod, ExpectedException(typeof(InvalidOperationException))]
    public void TestThrowsIfPdfPrinterNotFound()
    {
      var dir = Directory.GetCurrentDirectory();

      using (var pdfPrint = new PdfPrint(pdfPrinterName: "Nonexistent PDF printer"))
      {
        pdfPrint.Print($"{dir}\\sample.docx", "out.pdf");
      }
    }
  }
}
