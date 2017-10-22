# cs-pdf-to-image

A x86 simple library to convert pdf to image for .net. This library is a derivation of Mark Redman's work on PDFConvert using Ghostscript gsdll32.dll. The derivation will automatically make the Ghostscript gsdll32 available on computer that does not have it installed, so that that developers do not need to worry whether ghostscript is available on the end user's computer in order for it to work.



# Usage

```cs
string pdf_filename="sample.pdf";
string png_filename="converted.png"; 
List<string> errors = cs_pdf_to_image.Pdf2Image.Convert(pdf_filename, png_filename);
```


