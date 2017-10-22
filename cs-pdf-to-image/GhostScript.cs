// ---------------------------------------------------------------------------------------------------------------
// Decription:
// .Net Wrapper for the GhostScript gsdll32.dll Interpreter API
//
// Author:
// Mark Redman
// email: mark@redmanscave.com
// blog:  redmanscave.blogspot.com
// web:   www.redmanscave.com
//
// Copyright © 2008 Mark Redman.
//
// Contributors:
// Special thanks to the following people who have contributed to this project;
// Curia Damiano
// Gabriel Deak Jahn
// Steve Gaetjens
// Jon Gilkison
// Scott Horton
// David Laub
// 
// Change History:
// Version 1.0
// 11/10/2003; Mark Redman; Created, based on Ghostscript version 8.10
// 13/09/2004; Mark Redman; Added Convert function overload to return the converted pages of a file as file paths
// 26/08/2005; Mark Redman; Modification to Print method
// Version 2.0
// 17/08/2007; Mark Redman; Removed unused code
// 18/08/2007; Mark Redman; Updated for GhostScript version 8.60 
// 19/08/2007; Mark Redman; Updated and cleaned up code, Added Stdio callback and progress events
// 20/08/2007; Mark Redman; Added CreateColorSeparations method
// 20/01/2008; Mark Redman; Added Usage examples
// 
// Notes:
// 1) There has been no attempt to make this backwards compatible for previous version of the class.
// 2) Not all output device support has been implemented.
// 3) Some Ghostscript specific naming conventions have been brought over from the Ghostscript
//    documentation to make it easier to see whats going on and will hopefully make updating this
//    class easier.
// 4) Not sure what the Stdio IN callbacks can be used for?
// 5) The Processing Started, Page and Completed Events can be use to get the page count and monitor progress.
//    (I had some corrupt memory errors when consuming these in the GUI on multi-page files? comments welcome)
//
// Usage:
// 1) Get File Separations and Spot Colour Names
//    (Shows the optional support for the callback events)
//    
//    ...
//    lock (typeof(Made4Print.GhostScript))
//    {
//        GC.Collect();
//    
//        using (Made4Print.GhostScript ghostScript = new Made4Print.GhostScript(GhostScriptFolder))
//        {
//            ghostScript.OnProcessingStarted += new GhostScript.ProcessingStartedEventHandler(ghostScript_OnProcessingStarted);
//            ghostScript.OnProcessingPage += new GhostScript.ProcessingPageEventHandler(ghostScript_OnProcessingPage);
//            ghostScript.OnProcessingCompleted += new GhostScript.ProcessingCompletedEventHandler(ghostScript_OnProcessingCompleted);
//            ghostScript.OnStdErrCallbackMessage += new GhostScript.StdioCallbackMessageHandler(ghostScript_OnStdErrCallbackMessage);
//            
//            // Processes theh file separations and returns the list of files created
//            OutputFilenames = ghostScript.CreateColorSeparations(filename, OutputFolder, TempFolder, OutputResolution);
//            
//            // Gets the list of spot colours
//            SpotColorNames.AddRange(ghostScript.SpotColorSeparationNames);
//        }
//    }
//    ...
//
// 2) Convert .pdf file to series of .jpg files
//    (In this case the output file is not specified, so all the pages will be converted)
//  
//    ...
//    lock (typeof(Made4Print.GhostScript))
//    {
//        GC.Collect();
//        
//        // Select output device  
//        Made4Print.GhostScript.OutputDevice outputDevice = GhostScript.OutputDevice.jpeg;
//        // Select output device options
//        Made4Print.GhostScript.DeviceOption[] deviceOptions = Made4Print.GhostScript.DeviceOptions.jpg(100);
//    
//        using (Made4Print.GhostScript ghostScript = new Made4Print.GhostScript(GhostScriptFolder))
//        {
//            OutputFilenames = ghostScript.Convert(outputDevice, deviceOptions, InputFilename, OutputFolder, String.Empty, TempFolder, OutputResolution);
//        }
//     }
//     ...
//
// License:
// THE AUTHOR GRANTS ALL USERS WHO AGREE TO THIS LICENSE A NONEXCLUSIVE, FREE OF CHARGE, LICENSE TO USE THE CODE 
// IN THIS FILE.
// YOU CAN USE THIS CODE FOR ANY NON-COMMERCIAL PURPOSE, INCLUDING DISTRIBUTING DERIVATIVE WORKS.
// YOU CAN USE THIS CODE FOR ANY COMMERCIAL PURPOSE WITH THE ONLY RESTRICTION THAT YOU MAY NOT USE IT, IN WHOLE OR 
// IN PART, TO CREATE A COMMERCIAL GHOSTSCRIPT WRAPPER COMPONENT. BASICALLY, YOU CAN USE THIS CODE AND MODIFY IT 
// TO USE IN YOUR OWN COMMERCIAL SOFTWARE, YOU JUST CAN'T TAKE THE CODE MODIFY IT AND SELL IT AS A PRODUCT.
// 
// Disclaimer:
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT ANY WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT 
// LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. 
// IN NO EVENT SHALL THE AUTHORS OR ITS CONTRBITORS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER 
// IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE CODE IN THIS FILE 
// OR THE USE OR OTHER DEALINGS WITH THIS CODE.
// ---------------------------------------------------------------------------------------------------------------
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.IO;

namespace Made4Print
{
    /// <summary>
    /// GhostScript DLL (gsdll32.dll) Wrapper Class
    /// </summary>
    public class GhostScript : IDisposable
    {
        #region Constants

        const string GSDLL32 = "gsdll32.dll";
        const string KERNEL32 = "kernel32";

        #endregion

        #region Event Handling

        public class ProcessingStartedEventArgs : EventArgs
        {
            private int _PageCount;

            private ProcessingStartedEventArgs()
            {
            }
            public ProcessingStartedEventArgs(int pageCount)
            {
                _PageCount = pageCount;
            }
            public int PageCount
            {
                get
                {
                    return _PageCount;
                }
            }
        }

        public class ProcessingPageEventArgs : EventArgs
        {
            private int _Page;
            private int _PageCount;

            private ProcessingPageEventArgs()
            {
            }
            public ProcessingPageEventArgs(int page, int pageCount)
            {
                _Page = page;
                _PageCount = pageCount;
            }

            public int Page
            {
                get
                {
                    return _Page;
                }
            }

            public int PageCount
            {
                get
                {
                    return _PageCount;
                }
            }
        }

        public class ProcessingCompletedEventArgs : EventArgs
        {
            private int _PageCount;
            private List<string> _OuputFilenames;

            private ProcessingCompletedEventArgs()
            {
            }
            public ProcessingCompletedEventArgs(int pageCount, List<string> ouputFilenames)
            {
                _PageCount = pageCount;
                _OuputFilenames = ouputFilenames;
            }
            public int PageCount
            {
                get
                {
                    return _PageCount;
                }
            }

            public List<string> OuputFilenames
            {
                get
                {
                    return _OuputFilenames;
                }
            }
        }

        public delegate int StdioMessageEventHandler(IntPtr handle, IntPtr pointer, int count);
        public delegate void StdioCallbackMessageHandler(string message);
        public delegate void InMessageEventHandler(string message);
        public delegate void OutMessageEventHandler(string message);
        public delegate void ErrorMessageEventHandler(string message);
        public delegate void ProcessingStartedEventHandler(ProcessingStartedEventArgs e);
        public delegate void ProcessingPageEventHandler(ProcessingPageEventArgs e);
        public delegate void ProcessingCompletedEventHandler(ProcessingCompletedEventArgs e);

        /// <summary>
        /// Stdio IN Callback Message
        /// Returns the number of characters read, 0 for EOF, or -1 for error
        /// </summary>
        public event StdioCallbackMessageHandler OnStdInCallbackMessage;
        /// <summary>
        /// Stdio OUT Callback Message
        /// Return the number of characters written
        /// </summary>
        public event StdioCallbackMessageHandler OnStdOutCallbackMessage;
        /// <summary>
        /// Stdio ERROR Callback Message
        /// Return the number of characters written
        /// </summary>
        public event StdioCallbackMessageHandler OnStdErrCallbackMessage;
        /// <summary>
        /// Event raised when the first page is about to start processing
        /// </summary>
        public event ProcessingStartedEventHandler OnProcessingStarted;
        /// <summary>
        /// Event raised before each page is processed
        /// </summary>
        public event ProcessingPageEventHandler OnProcessingPage;
        /// <summary>
        /// Event raised after all pages have been processed
        /// </summary>
        public event ProcessingCompletedEventHandler OnProcessingCompleted;

        private int RaiseStdInCallbackMessageEvent(IntPtr handle, IntPtr pointer, int count)
        {
            if (OnStdInCallbackMessage != null)
            {
                string message = Marshal.PtrToStringAnsi(pointer);
                OnStdInCallbackMessage(message);
            }
            return count;
        }

        private int RaiseStdOutCallbackMessageEvent(IntPtr handle, IntPtr pointer, int count)
        {
            // Raise StdOut Callback Message Event
            if (OnStdOutCallbackMessage != null)
            {
                OnStdInCallbackMessage(Marshal.PtrToStringAnsi(pointer));
            }

            if (OnProcessingStarted != null || OnProcessingPage != null)
            {
                string message = Marshal.PtrToStringAnsi(pointer).Trim();

                if (_PageCount <= 0)
                {
                    // Attempt to get page count from callback message
                    if (message.StartsWith("Processing"))
                    {
                        try
                        {
                            _PageCount = Int32.Parse(message.Substring(0, message.IndexOf("\n")).Replace("Processing pages 1 through ", String.Empty).Replace(".", String.Empty).Trim());
                        }
                        catch
                        {
                            // Ignore error in case the callback message changes
                        }

                        RaiseProcessingStartedEvent(new ProcessingStartedEventArgs(_PageCount));
                    }
                }

                // Attempt to Get page number from callback message
                if (message.StartsWith("Page"))
                {
                    int page = 0;

                    try
                    {
                        page = Int32.Parse(message.Substring(0, message.IndexOf("\n")).Replace("Page ", String.Empty).Trim());
                    }
                    catch
                    {
                        // Ignore error in case the callback message changes
                    }

                    RaiseProcessingPageEvent(new ProcessingPageEventArgs(page, _PageCount));
                }

            }

            return count;
        }

        private int RaiseStdErrCallbackMessageEvent(IntPtr handle, IntPtr pointer, int count)
        {
            string message = Marshal.PtrToStringAnsi(pointer);

            // Attempt to et spot color separation names from callback message
            if (message.StartsWith("%%SeparationName:"))
            {
                string separationName = message.Replace("%%SeparationName:", String.Empty).Trim();
                if (!_SpotColorSeparationNames.Contains(separationName))
                {
                    _SpotColorSeparationNames.Add(separationName);
                }
            }

            if (OnStdErrCallbackMessage != null)
            {
                OnStdErrCallbackMessage(message);
            }
            return count;
        }

        private void RaiseProcessingStartedEvent(ProcessingStartedEventArgs processingStartedEventArgs)
        {
            if (OnProcessingStarted != null)
            {
                OnProcessingStarted(processingStartedEventArgs);
            }
        }

        private void RaiseProcessingPageEvent(ProcessingPageEventArgs processingPageEventArgs)
        {
            if (OnProcessingPage != null)
            {
                OnProcessingPage(processingPageEventArgs);
            }
        }

        private void RaiseProcessingCompletedEvent(ProcessingCompletedEventArgs processingCompletedEventArgs)
        {
            if (OnProcessingCompleted != null)
            {
                OnProcessingCompleted(processingCompletedEventArgs);
            }
        }

        #endregion

        #region Data Types

        /// <summary>
        /// Represents the Option parameter switch and Value pair that defines an Output Device Option
        /// </summary>
        public struct DeviceOption
        {
            public string Option;
            public string Value;

            public DeviceOption(string option, string optionValue)
            {
                Option = option;
                Value = optionValue;
            }
        }

        #endregion

        #region Enumerations

        /// <summary>
        /// These are the formats GhostScript is capable of interpreting.
        /// </summary>
        public enum SupportedFormats
        {
            /// <summary>
            /// PS, PostScript.
            /// </summary>
            ps,
            /// <summary>
            /// EPS, Encapsulated PostScript.
            /// </summary>
            eps,
            /// <summary>
            /// DOS EPS, DOS Encapsulated PostScript.
            /// </summary>
            epsf,
            /// <summary>
            /// PDF, Portable Document Format.
            /// </summary>
            pdf
        }

        /// <summary>
        /// Output devices for various file formats, high level devices such as PDF, PS and display devices
        /// </summary>
        public enum OutputDevice
        {
            // Image File Formats
            /// <summary>
            /// PNG, Portable Network Graphics format, 24-bit RGB color.
            /// </summary>
            png16m,
            /// <summary>
            /// PNG, Portable Network Graphics format, grayscale.
            /// </summary>
            pnggray,
            /// <summary>
            /// PNG, Portable Network Graphics format, 8-bit color.
            /// </summary>
            png256,
            /// <summary>
            /// PNG, Portable Network Graphics format, 4-bit color.
            /// </summary>
            png16,
            /// <summary>
            /// PNG, Portable Network Graphics format, black-and-white.
            /// </summary>
            pngmono,
            /// <summary>
            /// PNG, Portable Network Graphics format, 32-bit RGBA color with transparency indicating pixel coverage.
            /// </summary>
            pngalpha,
            /// <summary>
            /// JPEG File Interchange Format.
            /// </summary>
            jpeg,
            /// <summary>
            /// Grayscale JPEG File Interchange Format.
            /// </summary>
            jpeggray,
            /// <summary>
            /// PNM, Portable Network Map.
            /// </summary>
            pbm,
            /// <summary>
            /// PNM, Portable Network Map.
            /// </summary>
            pbmraw,
            /// <summary>
            /// PNM, Portable Network Map.
            /// </summary>
            pgm,
            /// <summary>
            /// PNM, Portable Network Map.
            /// </summary>
            pgmraw,
            /// <summary>
            /// PNM, Portable Network Map.
            /// </summary>
            pgnm,
            /// <summary>
            /// PNM, Portable Network Map.
            /// </summary>
            pgnmraw,
            /// <summary>
            /// PNM, Portable Network Map.
            /// </summary>
            pnm,
            /// <summary>
            /// PNM, Portable Network Map.
            /// </summary>
            pnmraw,
            /// <summary>
            /// PNM, Portable Network Map.
            /// </summary>
            ppm,
            /// <summary>
            /// PNM, Portable Network Map.
            /// </summary>
            ppmraw,
            /// <summary>
            /// PNM, Portable Network Map.
            /// </summary>
            pkm,
            /// <summary>
            /// PNM, Portable Network Map.
            /// </summary>
            pkmraw,
            /// <summary>
            /// PNM, Portable Network Map.
            /// </summary>
            pksm,
            /// <summary>
            /// PNM, Portable Network Map.
            /// </summary>
            pksmraw,
            /// <summary>
            /// TIF, Tagged Image File Format, 8-bit RGB uncompressed gray output.
            /// </summary>
            tiffgray,
            /// <summary>
            /// TIF, Tagged Image File Format, 12-bit RGB uncompressed color output (4 bits per component).
            /// </summary>
            tiff12nc,
            /// <summary>
            /// TIF, Tagged Image File Format, 24-bit RGB uncompressed color output (8 bits per component).
            /// </summary>
            tiff24nc,
            /// <summary>
            /// TIF, Tagged Image File Format, 32-bit CMYK uncompressed color output (8 bits per component).
            /// </summary>
            tiff32nc,
            /// <summary>
            /// TIF, Tagged Image File Format, The tiffsep device creates multiple output files.
            /// The device creates a single 32 bit composite CMYK file (tiff32nc format) and multiple tiffgray files. 
            /// A tiffgray file is created for each separation.
            /// See description at:
            /// <see cref="http://ghostscript.com/doc/8.54/Devices.htm#TIFF"/>
            /// </summary>
            tiffsep,
            /// <summary>
            /// TIF, Tagged Image File Format, Black-and White G3 fax encoding with no EOLs.
            /// </summary>
            tiffcrle,
            /// <summary>
            /// TIF, Tagged Image File Format, Black-and White G3 fax encoding with EOLs.
            /// </summary>
            tiffg3,
            /// <summary>
            /// TIF, Tagged Image File Format, Black-and White 2-D G3 fax encoding.
            /// </summary>
            tiffg32d,
            /// <summary>
            /// TIF, Tagged Image File Format, Black-and White G4 fax encoding.
            /// </summary>
            tiffg4,
            /// <summary>
            /// TIF, Tagged Image File Format, Black-and White LZW-compatible (tag = 5) compression.
            /// </summary>
            tifflzw,
            /// <summary>
            /// TIF, Tagged Image File Format, Black-and White PackBits (tag = 32773) compression.
            /// </summary>
            tiffpack,
            /// <summary>
            /// FAX, Raw fax format, Black-and White G3 fax encoding with EOLs.
            /// </summary>
            faxg3,
            /// <summary>
            /// FAX, Raw fax format, Black-and White 2-D G3 fax encoding.
            /// </summary>
            faxg32d,
            /// <summary>
            /// FAX, Raw fax format, Black-and White G4 fax encoding.
            /// </summary>
            faxg4,
            /// <summary>
            /// BMP, MS Windows bitmap.
            /// </summary>
            bmpmono,
            /// <summary>
            /// BMP, MS Windows bitmap.
            /// </summary>
            bmpgray,
            /// <summary>
            /// BMP, MS Windows bitmap.
            /// </summary>
            bmpsep1,
            /// <summary>
            /// BMP, MS Windows bitmap.
            /// </summary>
            bmpsep8,
            /// <summary>
            /// BMP, MS Windows bitmap.
            /// </summary>
            bmp16,
            /// <summary>
            /// BMP, MS Windows bitmap.
            /// </summary>
            bmp256,
            /// <summary>
            /// BMP, MS Windows bitmap.
            /// </summary>
            bmp16m,
            /// <summary>
            /// BMP, MS Windows bitmap.
            /// </summary>
            bmp32b,
            /// <summary>
            /// PCX format.
            /// </summary>
            pcxmono,
            /// <summary>
            /// PCX format.
            /// </summary>
            pcxgray,
            /// <summary>
            /// PCX format.
            /// </summary>
            pcx16,
            /// <summary>
            /// PCX format.
            /// </summary>
            pcx256,
            /// <summary>
            /// PCX format.
            /// </summary>
            pcx24b,
            /// <summary>
            /// PCX format.
            /// </summary>
            pcxcmyk,
            /// <summary>
            /// PSD CMYK format, this device supports spot colors.
            /// </summary>
            psdcmyk,
            /// <summary>
            /// PSD RGB format.
            /// </summary>
            psdrgb,

            // High level devices
            /// <summary>
            /// PDF writer, outputs PDF.
            /// </summary>
            pdfwrite,
            /// <summary>
            /// PS writer, outputs postscript.
            /// </summary>
            pswrite,
            /// <summary>
            /// EPS write, outputs encapsulated postscript.
            /// </summary>
            epswrite,
            /// <summary>
            /// PXL, Mono output in the HP PCL-XL graphic language used by many laser printers.
            /// </summary>
            pxlmono,
            /// <summary>
            /// PXL, Color output in the HP PCL-XL graphic language used by many laser printers.
            /// </summary>
            pxlcolor
            // Others have not been implmeneted yet.
        }

        public enum GraphicsAlphaBits
        {
            NotSet = 0,
            Low = 1,
            Medium = 2,
            Optimum = 4
        }

        /// <summary>
        /// GhostScript error code enumeration. these are taken from the GhostScript error.h file.
        /// Custom errors start at -10000.
        /// </summary>
        public enum ReturnCode : int
        {
            // Postscript level 1 errors
            e_unknownerror = -1,
            e_dictfull = -2,
            e_dictstackoverflow = -3,
            e_dictstackunderflow = -4,
            e_execstackoverflow = -5,
            e_interrupt = -6,
            e_invalidaccess = -7,
            e_invalidexit = -8,
            e_invalidfileaccess = -9,
            e_invalidfont = -10,
            e_invalidrestore = -11,
            e_ioerror = -12,
            e_limitcheck = -13,
            e_nocurrentpoint = -14,
            e_rangecheck = -15,
            e_stackoverflow = -16,
            e_setackunderflow = -17,
            e_syntaxerror = -18,
            e_timeout = -19,
            e_typecheck = -20,
            e_undefined = -21,
            e_undefinedfilename = -22,
            e_undefinedresult = -23,
            e_unmatchedmark = -24,
            e_VMerror = -25,
            // Additional level 2 and DPS errors
            e_configurationerror = -26,
            e_invalidcontext = -27,
            e_undefinedresource = -28,
            e_unregistered = -29,
            // Pseudo-errors used by ghostscript internally
            e_invalidid = -30,                                  // invalidid is for the NeXT DPS extension.
            e_fatal = -100,                                     // Internal code for a fatal error. gs_interpret also returns this for a .quit with a positive exit code.
            e_Quit = -101,                                      // Internal code for the .quit operator. The real quit code is an integer on the operand stack. gs_interpret returns this only for a .quit with a zero exit code.
            e_InterpreterExit = -102,                           // Internal code for a normal exit from the interpreter. Do not use outside of interp.c.
            e_RemapColor = -103,                                // Internal code that indicates that a procedure has been stored in the remap_proc of the graphics state, and should be called before retrying the current token.  This is used for color remapping involving a call back into the interpreter -- inelegant, but effective.
            e_ExecStackUnderflow = -104,                        // Internal code to indicate we have underflowed the top block of the e-stack.
            e_VMreclaim = -105,                                 // Internal code for the vmreclaim operator with a positive operand. We need to handle this as an error because otherwise the interpreter won't reload enough of its state when the operator returns.
            e_NeedInput = -106,                                 // Internal code for requesting more input from run_string.
            e_NeedStdin = -107,                                 // Internal code for stdin callout.
            e_NeedStdout = -108,                                // Internal code for stdout callout.
            e_NeedStderr = -109,                                // Internal code for stderr callout.
            e_Info = -110,                                      // Internal code for a normal exit when usage info is displayed. This allows Window versions of Ghostscript to pause until the message can be read.
            // Custom Errors
            FileTypeNotSupportedByInterpreter = -10000,         // Custom GhostScript Error: Input file type is not supported by the interpreter.
            UnableToLoadGhostScriptDll = -10001,                // Custom GhostScript Error: Unable to load GhostScript DLL (gsdll32.dll)
            GhostScriptDllNotFound = -10002                     // Custom GhostScript Error: GhostScript DLL not found in the specified Library Path
        }

        #endregion

        #region Unmanaged Dynamic-Link Library (DLL) Imports

        [DllImport(KERNEL32)]
        private extern static int LoadLibrary(string filename);

        [DllImport(KERNEL32)]
        private extern static bool FreeLibrary(int handle);

        [DllImport(GSDLL32, EntryPoint = "gsapi_revision", CharSet = CharSet.Ansi)]
        private static extern int gsapi_revision([In, Out] gsapi_revision_t revision, int len);

        [DllImport(GSDLL32, EntryPoint = "gsapi_new_instance", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        public static extern int gsapi_new_instance(ref IntPtr pInstance, out IntPtr pCaller);

        [DllImport(GSDLL32, EntryPoint = "gsapi_delete_instance", CharSet = CharSet.Ansi)]
        private static extern int gsapi_delete_instance(IntPtr pInstance);

        [DllImport(GSDLL32, EntryPoint = "gsapi_exit", CharSet = CharSet.Ansi)]
        private static extern int gsapi_exit(IntPtr pInstance);

        [DllImport(GSDLL32, EntryPoint = "gsapi_init_with_args", CharSet = CharSet.Ansi)]
        private static extern int gsapi_init_with_args(IntPtr pInstance, int argc, [In, Out] String[] argv);

        [DllImport(GSDLL32, EntryPoint = "gsapi_run_file", CharSet = CharSet.Ansi)]
        private static extern int gsapi_run_file(IntPtr pInstance, string strFilename, int nErrors, int nExitCode);

        [DllImport(GSDLL32, EntryPoint = "gsapi_set_stdio", CharSet = CharSet.Ansi)]
        private static extern int gsapi_set_stdio(IntPtr pInstance, StdioMessageEventHandler gsdll_stdin, StdioMessageEventHandler gsdll_stdout, StdioMessageEventHandler gsdll_stderr);

        #endregion

        #region Private Property Members

        private int _Handle = 0;
        private gsapi_revision_t _VersionInfo = null;
        private int _PageCount = 0;
        private List<string> _SpotColorSeparationNames = new List<string>();

        #endregion

        #region Initialization

        private GhostScript()
        {
        }
        public GhostScript(string libraryPath)
        {
            // Check Library Path contains GhostScript DLL
            if (!File.Exists(Path.Combine(libraryPath, GSDLL32)))
            {
                System.Diagnostics.Debug.WriteLine("GhostScriptDllNotFound Exception raised");
                throw (new GhostScriptException((int)ReturnCode.GhostScriptDllNotFound, GetGSErrorMessage((int)ReturnCode.GhostScriptDllNotFound)));
            }

            // Get the GhostScript Version Info
            // Note: This also checks that the GhostScript DLL is available
            _VersionInfo = GetVersion();
            System.Diagnostics.Debug.WriteLine("GhostScript Revision = " + _VersionInfo.Revision.ToString());
            
            InitializePrivateMembers();

            _Handle = LoadLibrary(Path.Combine(libraryPath, "gsdll32.dll"));
            System.Diagnostics.Debug.WriteLine("GhostScript Handle = " + _Handle.ToString());
        }

        /// <summary>
        /// Initializes private property members that may change during file processing.
        /// </summary>
        private void InitializePrivateMembers()
        {
            _PageCount = 0;
            _SpotColorSeparationNames.Clear();
        }

        #endregion

        #region Public Property Members

        /// <summary>
        /// Returns the GhostScript DLL Revision information
        /// </summary>
        public gsapi_revision_t VersionInfo
        {
            get
            {
                return _VersionInfo;
            }
        }

        /// <summary>
        /// Returns the Spot Colour Separation Names<br>
        /// Note: The value is only available when the CreateColorSeparations is used.
        /// </summary>
        public List<string> SpotColorSeparationNames
        {
            get
            {
                return _SpotColorSeparationNames;
            }
        }

        #endregion

        #region Public Methods

        public gsapi_revision_t GetVersion()
        {
            try
            {
                gsapi_revision_t revisionInfo = new gsapi_revision_t();
                gsapi_revision(revisionInfo, Marshal.SizeOf(revisionInfo));
                return revisionInfo;
            }
            catch (System.DllNotFoundException)
            {
                throw (new GhostScriptException((int)ReturnCode.UnableToLoadGhostScriptDll, GetGSErrorMessage((int)ReturnCode.UnableToLoadGhostScriptDll)));
            }
        }

        /// <summary>
        /// Returns true if the file type for the specified file extension is supported.
        /// Supported formats (.ps, .eps, epsf, pdf)
        /// </summary>
        /// <param name="extension"></param>
        /// <returns></returns>
        public static bool FileTypeSupported(string extension)
        {
            extension = extension.Replace(".", String.Empty).ToLower();

            foreach (SupportedFormats supportedFormats in Enum.GetValues(typeof(SupportedFormats)))
            {
                if (supportedFormats.ToString().ToLower() == extension)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Generic Convert function to convert a file to the specified output device.
        /// If an output filename is specified, the first page will be converted to the file specified
        /// If the output filename is in the form of a filename template for multiple pages, 
        /// it will convert the file accordingly.
        /// If no output filename is not specified, a template to convert all pages will be used.
        /// Returns the List of output file names for all files created.
        /// </summary>
        /// <param name="outputDevice"></param>
        /// <param name="deviceOptions"></param>
        /// <param name="inputFileName"></param>
        /// <param name="outputPath"></param>
        /// <param name="outputFileName"></param>
        /// <param name="temporaryFileFolder"></param>
        /// <param name="resolution"></param>
        /// <returns></returns>
        /// Notes:
        /// Command Line parameters:
        /// The first argument is always ignored
        /// -dSAFER               : Disables the deletefile and renamefile operators, and the ability to open piped commands (%pipe%cmd) at all. Only %stdout and %stderr can be opened for writing. Disables reading of files other than %stdin, those given as a command line argument, or those contained on one of the paths given by LIBPATH and FONTPATH and specified by the system params /FontResourceDir and /GenericResourceDir.
        /// -dBATCH, -dPAUSE      : The -dBATCH -dNOPAUSE options disable the interactive prompting. The interpreter also quits gracefully if it encounters end-of-file or control-C. -sDEVICE      : Ghostscript has a notion of 'output devices' which handle saving or displaying the results in a particular format. Ghostscript comes with a diverse variety of such devices supporting vector and raster file output, screen display, driving various printers and communicating with other applications.
        /// -dCOLORSCREEN         : On high-resolution devices (at least 150 dpi resolution, or -dDITHERPPI specified), -dCOLORSCREEN forces the use of separate halftone screens with different angles for CMYK or RGB if halftones are needed (this produces the best-quality output)
        /// -dCOLORSCREEN=0       : -dCOLORSCREEN=0 uses separate screens with the same frequency and angle
        /// -dCOLORSCREEN=false   : -dCOLORSCREEN=false forces the use of a single binary screen. The default if COLORSCREEN is not specified is to use separate screens with different angles if the device has fewer than 5 bits per color, and a single binary screen (which is never actually used under normal circumstances) on all other devices.
        /// -dDOINTERPOLATE       : Turns on image interpolation for all images, improving image quality for scaled images at the expense of speed. Note that -dNOINTERPOLATE overrides -dDOINTERPOLATE if both are specified.
        /// -dTextAlphaBits=n,
        /// -dGraphicsAlphaBits=n : These options control the use of subsample antialiasing. Their use is highly recommended for producing high quality rasterizations. The subsampling box size n should be 4 for optimum output, but smaller values can be used for faster rendering. Antialiasing is enabled separately for text and graphics content. Allowed values are 1, 2 or 4.
        /// -dUseCIEColor         : Set UseCIEColor in the page device dictionary, remapping device-dependent color values through a CIE color space. This can can improve conversion of CMYK documents to RGB.
        ///
        /// How GhostScript Finds Files:
        /// When looking for initialization files (gs_*.ps, pdf_*.ps), font files, the Fontmap file, files named on 
        /// the command line, and resource files, Ghostscript first tests whether the file name specifies an 
        /// absolute path:
        /// Does the name have ':' as its second character, or begin with '/', '\', or '//servername/share/'
        /// If the test succeeds, Ghostscript tries to open the file using the name given. Otherwise it tries 
        /// directories in this order:  
        /// 1) The current directory (unless disabled by the -P- switch)
        /// 2) The directories specified by -I switches in the command line, if any
        /// 3) The directories specified by the GS_LIB environment variable, if any
        /// 4) The directories specified by the GS_LIB_DEFAULT macro (if any) in the makefile when this executable was 
        ///    built
        [FileIOPermissionAttribute(SecurityAction.Demand, Unrestricted = true)]
        public List<string> Convert(OutputDevice outputDevice, DeviceOption[] deviceOptions, string inputFileName, string outputPath, string outputFileName, string temporaryFileFolder, int resolution)
        {
            System.Diagnostics.Debug.WriteLine("GhostScript Convert called, inputFileName=" + inputFileName + " temporaryFileFolder=" + temporaryFileFolder);

            InitializePrivateMembers();

            List<string> outputFilenames = new List<string>();

            // Check if input file is supported
            if (!FileTypeSupported(Path.GetExtension(inputFileName)))
            {
                throw (new GhostScriptException((int)ReturnCode.FileTypeNotSupportedByInterpreter, GhostScript.GetGSErrorMessage((int)ReturnCode.FileTypeNotSupportedByInterpreter)));
            }

            // Construct Ouput Filename Template
            string outputFilenameTemplate = Path.Combine(outputPath, outputFileName);

            if (outputFileName.Trim() == String.Empty)
            {
                // Construct output filename template for multipage document
                string filename = Path.GetFileNameWithoutExtension(inputFileName.Replace("%", "_"));
                //filename = String.Concat(filename,  (filename.EndsWith("_") ? String.Empty : "_"));
                outputFilenameTemplate = Path.Combine(outputPath, String.Concat(filename, "%01d.", GetOutputDeviceFileExtension(outputDevice)));
            }

            // Copy source file and perform processing on the copy
            string temporaryInputFilename = Path.Combine(temporaryFileFolder, String.Concat(Guid.NewGuid().ToString("N"), Path.GetExtension(inputFileName)));

            System.IO.File.Copy(inputFileName, temporaryInputFilename, true);

            // Construct Command Parameters to pass to the GhostScript DLL
            List<string> commandList = new List<string>();
            commandList.Add("convert"); //First parameter is ignored
            commandList.Add("-dSAFER");
            commandList.Add("-dBATCH");
            commandList.Add("-dNOPAUSE");
            commandList.Add("-dCOLORSCREEN");
            commandList.Add("-dDOINTERPOLATE");
            commandList.Add("-dTextAlphaBits=4");
            commandList.Add("-dGraphicsAlphaBits=4");
            commandList.Add("-dUseCIEColor");

            // Specify Output Device
            commandList.Add(GetOutputDeviceParameter(outputDevice));

            // Specify Output Device Options
            commandList.AddRange(GetDeviceOptionParameters(deviceOptions));

            // Specify Resolution
            if (resolution > 0)
            {
                commandList.Add(GetResolutionParameter(resolution));
            }

            // other parameters
            //commandList.Add("-c3000000");
            //commandList.Add("setvmthreshold -f");
            //commandList.Add("-r600");                 // Sets DPI resolution where XRES = YRES
            //commandList.Add("-r300x400");             // Sets DPI resolution (-rXRESxYRES) where XRES and YRES are defferent

            // Specify Output File
            commandList.Add(String.Concat("-sOutputFile=", outputFilenameTemplate));

            // Specify Input File
            commandList.Add(temporaryInputFilename);

            // Convert the command list into a string array
            string[] commandParameters = new string[commandList.Count];
            for (int counter = 0; counter <= commandList.Count - 1; counter++)
            {
                commandParameters[counter] = commandList[counter];
            }

            try
            {
                System.Diagnostics.Debug.WriteLine("GhostScript Convert.CallGSDLL");
                // Call GhostScript DLL
                CallGSDll(commandParameters);

                // Construct the result, listing output files in the output folder in the correct order 
                string outputFile = Path.Combine(outputPath, Path.GetFileNameWithoutExtension(outputFilenameTemplate.Replace("%01d", String.Empty)));
                string outputFileExtension = Path.GetExtension(outputFilenameTemplate);
                int counter = 1;
                string outputFileNameCheck = String.Concat(outputFile, counter, outputFileExtension);
                while (System.IO.File.Exists(outputFileNameCheck))
                {
                    // Add ouput file name to result
                    outputFilenames.Add(outputFileNameCheck);

                    // Increment counter, we use this number to check for the next file
                    counter++;

                    // Construct next filename to check
                    outputFileNameCheck = String.Concat(outputFile, counter, outputFileExtension);
                }

                RaiseProcessingCompletedEvent(new ProcessingCompletedEventArgs(_PageCount, outputFilenames));
            }
            finally
            {
                // Delete Temporary Input File
                if (System.IO.File.Exists(temporaryInputFilename))
                {
                    try
                    {
                        System.IO.File.Delete(temporaryInputFilename);
                    }
                    catch
                    {
                        // Ignore error the temporary file cannot be deleted
                    }
                }
            }

            System.Diagnostics.Debug.WriteLine("GhostScript Convert outputFileNameCount = " + outputFilenames.Count.ToString());

            return outputFilenames;
        }

        /// <summary>
        /// Creates multiple output files, a full color tif and a grayscale tif for each CMYK and Spot Color separation.<br>
        /// The SpotColorSeparationNames property will contain the spot color names.<br>
        /// The outputFileName determines the output file naming convention.
        /// if the outputFileName is not specified, the output files will be saved with the following naming convention:<br>
        /// Full Color tif : FILENAME.tif<br>
        /// Grayscale tif  : FILENAME_N.SEPARATION.tif<br>
        /// Where:<br>
        /// FILENAME = Input filename without extension<br>
        /// N = Page number<br>
        /// SEPARATION = Cyan, Magenta, Yellow, Black for CMYK Separations or the Separation Color Name<br>
        /// If an outputFileName is specified, FILENAME above is replaced with the name supplied.
        /// Notes:<br>
        /// If a spot color separation file is missing for a page, there was no content on that page, even if the spot colour
        /// is listed in SpotColorSeparationNames.
        /// </summary>
        /// <param name="inputFileName"></param>
        /// <param name="outputPath"></param>
        /// <param name="outputFileName"></param>
        /// <param name="temporaryFileFolder"></param>
        /// <param name="resolution"></param>
        /// <returns></returns>
        public List<string> CreateColorSeparations(string inputFileName, string outputPath, string outputFileName, string temporaryFileFolder, int resolution)
        {
            InitializePrivateMembers();

            List<string> outputFilenames = new List<string>();

            // Check if input file is supported
            if (!FileTypeSupported(Path.GetExtension(inputFileName)))
            {
                throw (new GhostScriptException((int)ReturnCode.FileTypeNotSupportedByInterpreter, GhostScript.GetGSErrorMessage((int)ReturnCode.FileTypeNotSupportedByInterpreter)));
            }

            System.Diagnostics.Debug.WriteLine("GhostScript - After File Support");

            // Construct output filename template for multipage document
            string filename = Path.GetFileNameWithoutExtension(inputFileName.Replace("%", "_"));

            // Use supplied prefix if specified
            if (outputFileName.Trim() != String.Empty)
            {
                filename = outputFileName.Trim().Replace("%", "_");
            }
            filename = String.Concat(filename, (filename.EndsWith("_") ? String.Empty : "_"));
            string outputFilenameTemplate = Path.Combine(outputPath, String.Concat(filename, "%01d.tif"));

            // Copy source file and perform processing on the copy
            string temporaryInputFilename = Path.Combine(temporaryFileFolder, String.Concat(Guid.NewGuid().ToString("N"), Path.GetExtension(inputFileName)));

            System.IO.File.Copy(inputFileName, temporaryInputFilename, true);

            // Construct Command Parameters to pass to the GhostScript DLL
            List<string> commandList = new List<string>();
            commandList.Add("tiffsep");
            commandList.Add("-dBATCH");
            commandList.Add("-dNOPAUSE");
            commandList.Add("-sDEVICE=tiffsep");
            //commandList.Add("-dUseCIEColor");
            commandList.Add("-dDOINTERPOLATE");
            commandList.Add("-dTextAlphaBits=4");
            commandList.Add("-dGraphicsAlphaBits=4");
            commandList.Add(String.Concat("-sOutputFile=", outputFilenameTemplate));
            commandList.Add(String.Concat("-r", (resolution > 0 ? resolution : 120)));
            commandList.Add("-MaxSeparations=8");
            commandList.Add(temporaryInputFilename);

            // Convert the command list into a string array
            string[] commandParameters = new string[commandList.Count];
            for (int counter = 0; counter <= commandList.Count - 1; counter++)
            {
                commandParameters[counter] = commandList[counter];
            }

            try
            {
                System.Diagnostics.Debug.WriteLine("GhostScript - Before CallGSDll");
                // Call GhostScript DLL
                CallGSDll(commandParameters);

                // Rename the numbered spot color filenames to include the spot colour name and construct result
                for (int counter = 1; counter <= _PageCount; counter++)
                {
                    string outputFile = Path.GetFileNameWithoutExtension(outputFilenameTemplate.Replace("%01d", String.Empty));
                    string[] pageFilenames = Directory.GetFiles(outputPath, String.Concat(outputFile, counter, "*"));

                    foreach (string pageFilename in pageFilenames)
                    {
                        bool filenameAdded = false;

                        string outputFilename = pageFilename;

                        for (int separationNumber = 0; separationNumber < _SpotColorSeparationNames.Count; separationNumber++)
                        {
                            if (outputFilename.Contains(String.Concat(".s", separationNumber, ".")))
                            {
                                outputFilename = outputFilename.Replace(String.Concat(".s", separationNumber, "."), String.Concat(".", _SpotColorSeparationNames[separationNumber], "."));

                                if (File.Exists(outputFilename))
                                {
                                    File.Delete(outputFilename);
                                }

                                // Copy File
                                File.Move(pageFilename, outputFilename);

                                // Update result
                                outputFilenames.Add(outputFilename);

                                filenameAdded = true;

                                break;
                            }
                        }

                        if (!filenameAdded)
                        {
                            // Update result
                            outputFilenames.Add(outputFilename);
                        }
                    }
                }

                RaiseProcessingCompletedEvent(new ProcessingCompletedEventArgs(_PageCount, outputFilenames));
            }
            finally
            {
                // Delete Temporary Input File
                if (System.IO.File.Exists(temporaryInputFilename))
                {
                    try
                    {
                        System.IO.File.Delete(temporaryInputFilename);
                    }
                    catch
                    {
                        // Ignore error the temporary file cannot be deleted
                    }
                }
            }

            return outputFilenames;
        }

        /// <summary>
        /// Calls the GhostScript interpreter with the command arguments specified.
        /// </summary>
        /// <param name="commandArguments"></param>
        public void CallGSDll(string[] commandParameters)
        {
            int errorCode = (int)ReturnCode.e_unknownerror;

            IntPtr ghostScriptPtr = IntPtr.Zero;
            IntPtr callerPtr = IntPtr.Zero;

            // Load new instance of Ghostscript
            errorCode = gsapi_new_instance(ref ghostScriptPtr, out callerPtr);

            // Setup Callback functions
            errorCode = gsapi_set_stdio(ghostScriptPtr, new StdioMessageEventHandler(RaiseStdInCallbackMessageEvent), new StdioMessageEventHandler(RaiseStdOutCallbackMessageEvent), new StdioMessageEventHandler(RaiseStdErrCallbackMessageEvent));

            if (errorCode >= 0)
            {
                try
                {
                    // Init the GhostScript interpreter
                    errorCode = gsapi_init_with_args(ghostScriptPtr, commandParameters.Length, commandParameters);

                    // Stop the Ghostscript interpreter
                    gsapi_exit(ghostScriptPtr);
                }
                finally
                {
                    // Release the Ghostscript instance handle
                    gsapi_delete_instance(ghostScriptPtr);
                }
            }

            // Ignore e_Quit error, Note: if stdio is used, there are more return codes to ignore.
            errorCode = (errorCode == (int)ReturnCode.e_Quit ? 0 : errorCode);

            // Throw custom exception if error occured calling the ghostscript interpreter API
            if (errorCode < 0)
            {
                StringBuilder commands = new StringBuilder();

                foreach (string command in commandParameters)
                {
                    commands.Append(String.Concat(command, " "));
                }

                throw (new GhostScriptException(errorCode, String.Concat("Command Arguments=[" + commands.ToString().Trim() + "] ", GetGSErrorMessage(errorCode), " (", errorCode.ToString(), ")")));
            }
        }

        /*
        /// <summary>
        /// Prints a file to a local printer.
        /// </summary>
        /// <param name="strFilePath"></param>
        /// <param name="strPrinter"></param>
        public static void Print(string strFilePath, string strPrinter)
        {
            string strFilename = String.Empty;
            string strExtension = String.Empty;
            string strTempSetupFile = String.Empty;

            // Get Filename.ext from specified file path
            strFilename = strFilePath.Substring(strFilePath.LastIndexOf("\\") + 1);
            int dotPosition = strFilename.LastIndexOf(".");
            if (dotPosition >= 0)
            {
                strExtension = strFilename.Substring(dotPosition + 1);
                strFilename = strFilename.Substring(0, dotPosition);
            }

            // Check if strFilePath is a supported file type
            bool bSupportFileType = false;
            foreach (SupportedFormats supportedFormats in System.Enum.GetValues(typeof(SupportedFormats)))
            {
                if (supportedFormats.ToString().Trim() == strExtension.Trim()) bSupportFileType = true;
            }
            if (bSupportFileType == false)
            {
                throw (new GhostScriptException((int)GSError.FileTypeNotSupportedByInterpreter, GhostScript.GetGSErrorMessage((int)GSError.FileTypeNotSupportedByInterpreter)));
            }

            ArrayList argsList = new ArrayList();
            argsList.Add("print"); //Ignored
            argsList.Add("-dBATCH");
            argsList.Add("-dNOPAUSE");

            string strSetupFile = ConfigurationSettings.AppSettings["GhostScriptPrintSetupFile"];

            if (strSetupFile != String.Empty && File.Exists(strSetupFile))
            {
                StreamReader oStream = File.OpenText(strSetupFile);
                string strSetup = oStream.ReadToEnd();
                oStream.Close();
                strTempSetupFile = strSetupFile.Replace(".ps", "temp.ps");
                StreamWriter oNewStream = File.CreateText(strTempSetupFile);
                oNewStream.Write(String.Format(strSetup, strPrinter.Replace("\\", "\\\\"), String.Concat(strFilename, ".", strExtension)));
                oNewStream.Flush();
                oNewStream.Close();
                argsList.Add(strTempSetupFile);
            }
            else
            {
                argsList.Add("-sDEVICE=mswinpr2");
                argsList.Add("-dNoCancel");
                //argsList.Add(String.Format("-sOutputFile=\"%printer%{0}\"", strPrinter)); original
                argsList.Add(String.Format("-sOutputFile=%printer%{0}", strPrinter)); // modified
            }
            argsList.Add(strFilePath);

            // convert the array list into a string array
            // This is so Arguments can be added and removed easily
            string[] strArgs = new string[argsList.Count];
            for (int counter = 0; counter <= argsList.Count - 1; counter++)
            {
                strArgs[counter] = argsList[counter].ToString();
            }

            try
            {
                CallGSDll(strArgs);
            }
            catch (GhostScriptException e)
            {
                //ignore io errors
                if (e.ErrorCode != (int)GSError.e_ioerror)
                    throw e;
            }
            finally
            {
                if (File.Exists(strTempSetupFile))
                    File.Delete(strTempSetupFile);
            }
        }
         * */

        #endregion

        #region Ouput Device Options

        /// <summary>
        /// DeviceOptions class
        /// </summary>
        public class DeviceOptions
        {
            /// <summary>
            /// Use DefaultOptions() to return an empty DeviceOptions[] array when the output device used 
            /// has no options, or the defaults are to be used.
            /// </summary>
            /// <returns></returns>
            public static DeviceOption[] DefaultOptions()
            {
                DeviceOption[] Options = new DeviceOption[0];
                return Options;
            }

            /// <summary>
            /// pngalpha output device options (RGB color in the form #RRGGBB, default white = #ffffff)
            /// </summary>
            /// <param name="backgroundColor">
            /// For the pngalpha device only, this sets the suggested background color in the PNG bKGD chunk. 
            /// When a program reading a PNG file does not support alpha transparency, the PNG library converts 
            /// the image using either a background color if supplied by the program or the bKGD chunk. 
            /// One common web browser has this problem, so when using color attributes eg: bgcolor="CCCC00" in a body tag on a web page 
            /// this option would need to be set to "#CCCC00" when creating alpha transparent PNG images for use on the page.
            /// </param>
            /// <returns></returns>
            public static DeviceOption[] pngalpha(string backgroundColor)
            {
                DeviceOption BackgroundColor = new DeviceOption("-dBackgroundColor", "16" + backgroundColor);

                DeviceOption[] Options = new DeviceOption[1];
                Options.SetValue(BackgroundColor, 0);

                return Options;
            }

            /// <summary>
            /// jpeg output device option for all jpeg devices. (integer from 0 to 100, default 75)
            /// </summary>
            /// <param name="quality"></param>
            /// Sets the quality level according to the widely used IJG quality scale, which balances the extent of 
            /// compression against the fidelity of the image when reconstituted. Lower values drop more information 
            /// from the image to achieve higher compression, and therefore have lower quality when reconstituted.
            /// <returns></returns>
            public static DeviceOption[] jpg(int quality)
            {
                DeviceOption Quality = new DeviceOption("-dJPEGQ=", quality.ToString());

                DeviceOption[] Options = new DeviceOption[1];
                Options.SetValue(Quality, 0);

                return Options;
            }

            /// <summary>
            /// tif options for black and white tif devices only.
            /// </summary>
            /// <param name="maxStripSize"></param>
            /// Set the maximum (uncompressed) size of a strip. (non-negative integer; default = 0)
            /// <param name="adjustWidth"></param>
            ///  (0 or 1; default = 1)
            ///  If this option set to 1 then, if the requested page width is close to either A4 (1728 columns) or 
            ///  B4 (2048 columns), set the page width to A4 or B4 respectively.
            /// <returns></returns>
            public static DeviceOption[] tif(int maxStripSize, int adjustWidth)
            {
                DeviceOption MaxStripSize = new DeviceOption("-dMaxStripSize=", maxStripSize.ToString());
                DeviceOption AdjustWidth = new DeviceOption("-dAdjustWidth=", adjustWidth.ToString());

                DeviceOption[] Options = new DeviceOption[2];
                Options.SetValue(MaxStripSize, 0);
                Options.SetValue(AdjustWidth, 0);

                return Options;
            }

            /// <summary>
            /// The tiffsep device creates multiple output files.
            /// The device creates a single 32 bit composite CMYK file (tiff32nc format) and multiple tiffgray files. 
            /// A tiffgray file is created for each separation.
            /// See description at:
            /// <see cref="http://ghostscript.com/doc/8.54/Devices.htm#TIFF"/>
            /// </summary>
            /// </summary>
            /// <returns></returns>
            public static DeviceOption[] tiffsep()
            {
                DeviceOption[] Options = new DeviceOption[0];
                return Options;
            }

            /// <summary>
            /// ps options for writing postscript. (Set to 1, 1.5, 2 or 3, default is 2)
            /// </summary>
            /// <param name="languageLevel"></param>
            /// Set the language level of the generated file. 
            /// Language level 1.5 is language level 1 with color extensions. 
            /// Currently language level 3 generates the same PostScript as 2.
            /// <returns></returns>
            public static DeviceOption[] ps(int languageLevel)
            {
                DeviceOption LanguageLevel = new DeviceOption("-dLanguageLevel=", languageLevel.ToString());

                DeviceOption[] Options = new DeviceOption[1];
                Options.SetValue(LanguageLevel, 0);

                return Options;
            }

            /// <summary>
            /// eps options for writing encapsulated postscript. (Set to 1, 1.5, 2 or 3, default is 2)
            /// </summary>
            /// <param name="languageLevel"></param>
            /// Set the language level of the generated file. 
            /// Language level 1.5 is language level 1 with color extensions. 
            /// Currently language level 3 generates the same PostScript as 2.
            /// <returns></returns>
            public static DeviceOption[] eps(int languageLevel)
            {
                DeviceOption LanguageLevel = new DeviceOption("-dLanguageLevel=", languageLevel.ToString());

                DeviceOption[] Options = new DeviceOption[1];
                Options.SetValue(LanguageLevel, 0);

                return Options;
            }

        }

        #endregion

        #region Helper Functions

        /// <summary>
        /// Returns the Output Device parameter
        /// </summary>
        /// <param name="outputDevice"></param>
        /// <returns></returns>
        public static string GetOutputDeviceParameter(OutputDevice outputDevice)
        {
            if (Enum.IsDefined(typeof(OutputDevice), outputDevice))
            {
                return String.Concat("-sDEVICE=", outputDevice);
            }
            else
            {
                return "-sDEVICE=unknown";
            }
        }

        /// <summary>
        /// Returns the Output Device Option parameters
        /// </summary>
        /// <param name="deviceOptions"></param>
        /// <returns></returns>
        public static List<string> GetDeviceOptionParameters(DeviceOption[] deviceOptions)
        {
            List<string> result = new List<string>();

            for (int counter = 0; counter <= deviceOptions.Length - 1; counter++)
            {
                result.Add(String.Concat(deviceOptions[counter].Option, deviceOptions[counter].Value));
            }

            return result;
        }

        /// <summary>
        /// Returns the DPI resolution parameter (where XRES = YRES)
        /// </summary>
        /// <param name="resolution"></param>
        /// <returns></returns>
        public static string GetResolutionParameter(int resolution)
        {
            return String.Concat("-r", resolution.ToString());
        }

        /// <summary>
        /// Returns the file extension for the specified output device
        /// </summary>
        /// <param name="outputDevice"></param>
        /// <returns></returns>
        public static string GetOutputDeviceFileExtension(OutputDevice outputDevice)
        {
            switch (outputDevice)
            {
                case OutputDevice.png16m:
                case OutputDevice.pnggray:
                case OutputDevice.png256:
                case OutputDevice.png16:
                case OutputDevice.pngmono:
                case OutputDevice.pngalpha:
                    {
                        return "png";
                    }
                case OutputDevice.jpeg:
                case OutputDevice.jpeggray:
                    {
                        return "jpg";
                    }
                case OutputDevice.pbm:
                case OutputDevice.pbmraw:
                case OutputDevice.pgm:
                case OutputDevice.pgmraw:
                case OutputDevice.pgnm:
                case OutputDevice.pgnmraw:
                case OutputDevice.pnm:
                case OutputDevice.pnmraw:
                case OutputDevice.ppm:
                case OutputDevice.ppmraw:
                case OutputDevice.pkm:
                case OutputDevice.pkmraw:
                case OutputDevice.pksm:
                case OutputDevice.pksmraw:
                    {
                        return "pnm";
                    }
                case OutputDevice.tiffgray:
                case OutputDevice.tiff12nc:
                case OutputDevice.tiff24nc:
                case OutputDevice.tiff32nc:
                case OutputDevice.tiffsep:
                case OutputDevice.tiffcrle:
                case OutputDevice.tiffg3:
                case OutputDevice.tiffg32d:
                case OutputDevice.tiffg4:
                case OutputDevice.tifflzw:
                case OutputDevice.tiffpack:
                    {
                        return "tif";
                    }
                case OutputDevice.faxg3:
                case OutputDevice.faxg32d:
                case OutputDevice.faxg4:
                    {
                        return "raw";
                    }
                case OutputDevice.bmpgray:
                case OutputDevice.bmpsep1:
                case OutputDevice.bmpsep8:
                case OutputDevice.bmp16:
                case OutputDevice.bmp256:
                case OutputDevice.bmp16m:
                case OutputDevice.bmp32b:
                    {
                        return "bmp";
                    }
                case OutputDevice.pcxmono:
                case OutputDevice.pcxgray:
                case OutputDevice.pcx16:
                case OutputDevice.pcx256:
                case OutputDevice.pcx24b:
                case OutputDevice.pcxcmyk:
                    {
                        return "pcx";
                    }
                case OutputDevice.psdcmyk:
                case OutputDevice.psdrgb:
                    {
                        return "psd";
                    }
                case OutputDevice.pdfwrite:
                    {
                        return "pdf";
                    }
                case OutputDevice.pswrite:
                    {
                        return "ps";
                    }
                case OutputDevice.epswrite:
                    {
                        return "eps";
                    }
                case OutputDevice.pxlmono:
                case OutputDevice.pxlcolor:
                    {
                        return "pxl";
                    }

                default: return String.Empty;
            }
        }

        #endregion

        #region Error Handling

        /// <summary>
        /// GhostScriptException Class
        /// </summary>
        public class GhostScriptException : System.Exception
        {
            private int _ErrorCode = -1;
            private DateTime _TimeStamp = DateTime.MinValue;

            /// <summary>
            /// Initializes the GhostScriptException class
            /// </summary>
            /// <param name="errorCode"></param>
            /// <param name="message"></param>
            public GhostScriptException(int errorCode, string message)
                : base(message)
            {
                _ErrorCode = errorCode;
                _TimeStamp = DateTime.Now;
            }

            /// <summary>
            /// Returns the error code returned by the call to the GhostScript Interpreter API
            /// </summary>
            public int ErrorCode
            {
                get { return _ErrorCode; }
            }

            /// <summary>
            /// Returns the System.DateTime for the date and time the error was raised.
            /// </summary>
            public DateTime TimeStamp
            {
                get { return _TimeStamp; }
            }
        }

        /// <summary>
        /// Returns the error message for the specified GhostScript error code
        /// </summary>
        /// <param name="errorCode"></param>
        /// <returns></returns>
        private static string GetGSErrorMessage(int returnCode)
        {
            switch (returnCode)
            {
                // Level 1 PostScript errors
                case (int)ReturnCode.e_unknownerror: return "Unknown error";
                case (int)ReturnCode.e_dictfull: return "level 1 error: e_dictfull";
                case (int)ReturnCode.e_dictstackoverflow: return "level 1 error: e_dictstackoverflow";
                case (int)ReturnCode.e_dictstackunderflow: return "level 1 error: e_dictstackunderflow";
                case (int)ReturnCode.e_execstackoverflow: return "level 1 error: e_execstackoverflow";
                case (int)ReturnCode.e_interrupt: return "level 1 error: e_interrupt";
                case (int)ReturnCode.e_invalidaccess: return "level 1 error: e_invalidaccess";
                case (int)ReturnCode.e_invalidexit: return "level 1 error: e_invalidexit";
                case (int)ReturnCode.e_invalidfileaccess: return "level 1 error: e_invalidfileaccess";
                case (int)ReturnCode.e_invalidfont: return "level 1 error: e_invalidfont";
                case (int)ReturnCode.e_invalidrestore: return "level 1 error: e_invalidrestore";
                case (int)ReturnCode.e_ioerror: return "level 1 error: e_ioerror";
                case (int)ReturnCode.e_limitcheck: return "level 1 error: e_limitcheck";
                case (int)ReturnCode.e_nocurrentpoint: return "level 1 error: e_nocurrentpoint";
                case (int)ReturnCode.e_rangecheck: return "level 1 error: e_rangecheck error";
                case (int)ReturnCode.e_stackoverflow: return "level 1 error: e_stackoverflow";
                case (int)ReturnCode.e_setackunderflow: return "level 1 error: e_setackunderflow";
                case (int)ReturnCode.e_syntaxerror: return "level 1 error: e_syntaxerror";
                case (int)ReturnCode.e_timeout: return "level 1 error: e_timeout";
                case (int)ReturnCode.e_typecheck: return "level 1 error: e_typecheck";
                case (int)ReturnCode.e_undefined: return "level 1 error: e_undefined";
                case (int)ReturnCode.e_undefinedfilename: return "level 1 error: e_undefinedfilename";
                case (int)ReturnCode.e_undefinedresult: return "level 1 error: e_undefinedresult";
                case (int)ReturnCode.e_unmatchedmark: return "level 1 error: e_unmatchedmark";
                case (int)ReturnCode.e_VMerror: return "level 1 error: e_VMerror error";
                // Level 2 and DPS errors
                case (int)ReturnCode.e_configurationerror: return "Level 2 error: e_configurationerror";
                case (int)ReturnCode.e_invalidcontext: return "Level 2 error: e_invalidcontext";
                case (int)ReturnCode.e_undefinedresource: return "Level 2: e_undefinedresource";
                case (int)ReturnCode.e_unregistered: return "Level 2: e_unregistered";
                case (int)ReturnCode.e_invalidid: return "Level 2: e_invalidid";
                // Pseudo internal ghostscript errors
                case (int)ReturnCode.e_fatal: return "Internal GhostScript code: e_fatal";
                case (int)ReturnCode.e_Quit: return "Internal GhostScript code: e_Quit";
                case (int)ReturnCode.e_InterpreterExit: return "Internal GhostScript code: e_InterpreterExit";
                case (int)ReturnCode.e_RemapColor: return "Internal GhostScript code: e_RemapColor";
                case (int)ReturnCode.e_ExecStackUnderflow: return "Internal GhostScript code: e_ExecStackUnderflow";
                case (int)ReturnCode.e_VMreclaim: return "Internal GhostScript code: e_VMreclaim";
                case (int)ReturnCode.e_NeedInput: return "Internal GhostScript code: e_NeedInput";
                case (int)ReturnCode.e_NeedStdin: return "Internal GhostScript code: e_NeedStdin";
                case (int)ReturnCode.e_NeedStdout: return "Internal GhostScript code: e_NeedStdout";
                case (int)ReturnCode.e_NeedStderr: return "Internal GhostScript code: e_NeedStderr";
                case (int)ReturnCode.e_Info: return "Internal GhostScript code: e_Info";
                // Custom Ghostscript errors
                case (int)ReturnCode.FileTypeNotSupportedByInterpreter: return "GhostScript Wrapper Class error: File type not supported by the ghostscript interpreter.";
                case (int)ReturnCode.UnableToLoadGhostScriptDll: return "GhostScript Wrapper Class error: Unable to load GhostScript DLL (gsdll32.dll).";
                case (int)ReturnCode.GhostScriptDllNotFound: return "GhostScript Wrapper Class error: GhostScript DLL (gsdll32.dll) not found in specified library path.";
                default: return "Unknown error.";
            }
        }

        #endregion

        #region Revision Class

        /// <summary>
        /// Revision class
        /// </summary>
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public class gsapi_revision_t
        {
            public string Product;
            public string Copyright;
            public int Revision;
            public int RevisionDate;
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            if (_Handle != -1)
            {
                FreeLibrary(_Handle);
            }
        }

        #endregion

        #region Finalization

        ~GhostScript()
        {
            Dispose();
        }

        #endregion

    }
}