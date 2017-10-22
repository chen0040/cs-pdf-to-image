using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pdf2png
{
    class Program
    {
        static void Main(string[] args)
        {
            string pdf_filename = cs_pdf_to_image.Pdf2Image.GetProgramFilePath("sample.pdf");
            string png_filename = cs_pdf_to_image.Pdf2Image.GetProgramFilePath("converted.png");

            if(!File.Exists(pdf_filename)) {
                File.WriteAllBytes(pdf_filename, Properties.Resources.pdf_sample);
            }
 
            if (args.Length > 2)
            {
                pdf_filename = args[1];
                png_filename = args[2];
            }

            List<string> errors = cs_pdf_to_image.Pdf2Image.Convert(pdf_filename, png_filename);
            if(errors.Any())
            {
                foreach(string error in errors)
                {
                    Console.WriteLine(error);
                }
            } else
            {
                Console.WriteLine("Conversion is successful.");
            }
        }
    }
}
