using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordRemoveHeadersFooters
{
    public class Program
    {
        static void Main(string[] args)
        {
            const string fileToFix = @"C:\Users\fpmor\OneDrive\XGILITY\Clements\Clements-Database - Copy.docx";
            RemoveHeadersAndFooters(fileToFix);
            Console.WriteLine("Press any key to close this window");
            Console.ReadKey();
        }

        // Remove all of the headers and footers from a document.
        private static void RemoveHeadersAndFooters(string filename)
        {
            // Given a document name, remove all of the headers and footers
            // from the document.
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filename, true, new OpenSettings{AutoSave = false}))
            {
                // Get a reference to the main document part.
                var docPart = doc.MainDocumentPart;

                // if there are any headers or footers...
                if (docPart.HeaderParts.Any() || docPart.FooterParts.Any())
                {
                    Console.WriteLine("Starting removal of headers and footers...");
                    // Remove the header and footer parts.
                    docPart.DeleteParts(docPart.HeaderParts);
                    docPart.DeleteParts(docPart.FooterParts);

                    // Get a reference to the root element of the main
                    // document part.
                    Document document = docPart.Document;

                    // Remove all references to the headers and footers.

                    // First, create a list of all descendants of type
                    // HeaderReference. Then, navigate the list and call
                    // Remove on each item to delete the reference.
                    var headers = document.Descendants<HeaderReference>().ToList();
                    foreach (var header in headers)
                    {
                        header.Remove();
                    }

                    // Create a list of all descendants of type
                    // FooterReference. Then, navigate the list and call
                    // Remove on each item to delete the reference.
                    var footers = document.Descendants<FooterReference>().ToList();
                    foreach (var footer in footers)
                    {
                        footer.Remove();
                    }

                    // Save the changes.
                    Console.WriteLine($"Saving {filename}");
                    document.Save();
                }
                else
                {
                    Console.WriteLine($"{filename} contains no headers or footers, so no action taken.");
                }
            }
        }

    }
}
