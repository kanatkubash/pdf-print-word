using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace windows_pdf_print
{
  public class PdfPrint : IDisposable
  {
    /// <summary>
    /// Default printer name for PDF printer
    /// </summary>
    public const string DEFAULT_PDF_PRINTER = "Microsoft Print to PDF";

    private Application wordApp;
    private string pdfPrinterName;

    public PdfPrint(object wordApp = null, string pdfPrinterName = null)
    {
      if (!this.IsWindows10())
        throw new Exception("This library works only on Windows 10");

      if (wordApp == null)
        wordApp = new Application();
      if (pdfPrinterName == null)
        pdfPrinterName = PdfPrint.DEFAULT_PDF_PRINTER;

      this.wordApp = (Application)wordApp;
      this.pdfPrinterName = pdfPrinterName;
    }

    public void Dispose()
    {
      if (this.wordApp != null)
      {
        this.wordApp.Application.Quit();
        this.wordApp = null;
      }
    }

    public void Print(string inFile, string outFile)
    {
      if (!File.Exists(inFile))
        throw new FileNotFoundException("Input file not found", inFile);

      if (!this.IsSupportedFile(inFile))
        throw new InvalidOperationException("File is not supported by word");

      string pdfPrinter = null;
      if ((pdfPrinter = GetPdfPrinter()) == null)
        throw new InvalidOperationException("Cannot find PDF printer");

      Document document = null;
      var dummyArg = Type.Missing;
      wordApp.ActivePrinter = pdfPrinter;

      try
      {
        /// as word interop asks for ref arguments, we should instantiate variables
        object falsy = false;
        object tru = true;
        object file = inFile;
        object output = outFile;
        document = this.wordApp.Documents.OpenNoRepairDialog(
          ref file,
          ref falsy,
          ref tru,
          ref dummyArg,
          ref dummyArg,
          ref dummyArg,
          ref dummyArg,
          ref dummyArg,
          ref dummyArg,
          ref dummyArg,
          ref dummyArg,
          ref dummyArg,
          ref falsy,
          ref dummyArg,
          ref tru,
          ref dummyArg
          );

        document.PrintOut(
          ref falsy,
          ref falsy,
          ref dummyArg,
          ref output,
          ref dummyArg,
          ref dummyArg,
          ref dummyArg,
          ref dummyArg,
          ref dummyArg,
          ref dummyArg,
          ref tru,
          ref dummyArg,
          ref dummyArg,
          ref dummyArg,
          ref dummyArg,
          ref dummyArg,
          ref dummyArg,
          ref dummyArg
          );
      }
      finally
      {
        /// graceful closing document
        if (document != null)
          document.Close(ref dummyArg, ref dummyArg, ref dummyArg);
      }
    }

    private bool IsWindows10()
    {
      return Environment.OSVersion.Platform == PlatformID.Win32NT &&
        Environment.OSVersion.Version.Major == 10;
    }

    private bool IsSupportedFile(string inFile)
    {
      return Regex.IsMatch(inFile, ".(docx?|txt|rtf)$");
    }

    private string GetPdfPrinter()
    {
      return PrinterSettings
        .InstalledPrinters
        .OfType<string>()
        .Where(printer => printer == this.pdfPrinterName)
        .FirstOrDefault();
    }
  }
}
