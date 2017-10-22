using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Collections;
using System.IO;

namespace PdfToImage
{
    /// <summary>
    /// Create by : TaGoH
    /// URL of the last version: http://www.codeproject.com/KB/cs/GhostScriptUseWithCSharp.aspx
    /// Description:
    /// Class to convert a pdf to an image using GhostScript DLL
    /// A big Credit for this code go to:Rangel Avulso he made me start with the project!
    /// </summary>
    /// <see cref="http://www.codeproject.com/KB/cs/GhostScriptUseWithCSharp.aspx"/>
    /// <seealso cref="http://www.hrangel.com.br/index.php/2006/12/04/converter-pdf-para-imagem-jpeg-em-c/"/>
    public class PDFConvert
    {
        #region Static
        /// <summary>The name of the DLL i'm using to work</summary>
        public const string GhostScriptDLLName = "gsdll32.dll";
        /// <summary>Use to check for default transformation</summary>
        /// <summary>Thanks to 	tchu_2000 to remind that u should never hardcode strings! :)</summary>
        private const string GS_OutputFileFormat = "-sOutputFile={0}";
        private const string GS_DeviceFormat = "-sDEVICE={0}";
        private const string GS_FirstParameter = "pdf2img";
        private const string GS_ResolutionXFormat = "-r{0}";
        private const string GS_ResolutionXYFormat = "-r{0}x{1}";
        private const string GS_GraphicsAlphaBits = "-dGraphicsAlphaBits={0}";
        private const string GS_TextAlphaBits = "-dTextAlphaBits={0}";
        private const string GS_FirstPageFormat = "-dFirstPage={0}";
        private const string GS_LastPageFormat = "-dLastPage={0}";
        private const string GS_FitPage = "-dPDFFitPage";
        private const string GS_PageSizeFormat = "-g{0}x{1}";
        private const string GS_DefaultPaperSize = "-sPAPERSIZE={0}";
        private const string GS_JpegQualityFormat = "-dJPEGQ={0}";
        private const string GS_RenderingThreads = "-dNumRenderingThreads={0}";
        private const string GS_Fixed1stParameter = "-dNOPAUSE";
        private const string GS_Fixed2ndParameter = "-dBATCH";
        private const string GS_Fixed3rdParameter = "-dSAFER";
        private const string GS_FixedMedia = "-dFIXEDMEDIA";
        private const string GS_QuiteOperation = "-q";
        private const string GS_StandardOutputDevice = "-";
        private const string GS_MultiplePageCharacter = "%";
        //Thanks to davalv for this font related options
        //http://www.codeproject.com/script/Membership/View.aspx?mid=3255201
        private const string GS_FontPath = "-sFONTPATH={0}";
        private const string GS_NoPlatformFonts = "-dNOPLATFONTS";
        private const string GS_NoFontMap = "-dNOFONTMAP";
        private const string GS_FontMap = "-sFONTMAP={0}";
        private const string GS_SubstitutionFont = "-sSUBSTFONT={0}";
        private const string GS_FCOFontFile = "-sFCOfontfile={0}";
        private const string GS_FAPIFontMap = "-sFAPIfontmap={0}";
        private const string GS_NoPrecompiledFonts = "-dNOCCFONTS";
        
        /// <summary>Mutex used to be sure that the DLL is not called twice</summary>
        private static System.Threading.Mutex mutex;
        #endregion
        #region Windows Import
        /// <summary>Needed to copy memory from one location to another, used to fill the struct</summary>
        /// <param name="Destination"></param>
        /// <param name="Source"></param>
        /// <param name="Length"></param>
        [DllImport("kernel32.dll", EntryPoint = "RtlMoveMemory")]
        static extern void CopyMemory(IntPtr Destination, IntPtr Source, uint Length);
        #endregion
        #region GhostScript Import

        /// <summary>Create a new instance of Ghostscript. This instance is passed to most other gsapi functions. The caller_handle will be provided to callback functions.
        ///  At this stage, Ghostscript supports only one instance. </summary>
        /// <param name="pinstance"></param>
        /// <param name="caller_handle"></param>
        /// <returns></returns>
        [DllImport(GhostScriptDLLName, EntryPoint = "gsapi_new_instance")]
        private static extern int gsapi_new_instance (out IntPtr pinstance, IntPtr caller_handle);

        /// <summary>This is the important function that will perform the conversion</summary>
        /// <param name="instance"></param>
        /// <param name="argc"></param>
        /// <param name="argv"></param>
        /// <returns></returns>
        [DllImport("gsdll32.dll", EntryPoint="gsapi_init_with_args")]
        private static extern int gsapi_init_with_args (IntPtr instance, int argc, IntPtr argv);
        /// <summary>
        /// Exit the interpreter. This must be called on shutdown if gsapi_init_with_args() has been called, and just before gsapi_delete_instance(). 
        /// </summary>
        /// <param name="instance"></param>
        /// <returns></returns>
        [DllImport("gsdll32.dll", EntryPoint="gsapi_exit")]
        private static extern int gsapi_exit (IntPtr instance);

        /// <summary>
        /// Destroy an instance of Ghostscript. Before you call this, Ghostscript must have finished. If Ghostscript has been initialised, you must call gsapi_exit before gsapi_delete_instance. 
        /// </summary>
        /// <param name="instance"></param>
        [DllImport("gsdll32.dll", EntryPoint="gsapi_delete_instance")]
        private static extern void gsapi_delete_instance (IntPtr instance);
        /// <summary>Get info about the version of Ghostscript i'm using</summary>
        /// <param name="pGSRevisionInfo"></param>
        /// <param name="intLen"></param>
        /// <returns></returns>
        [DllImport("gsdll32.dll", EntryPoint="gsapi_revision")]
        private static extern int gsapi_revision (ref GS_Revision pGSRevisionInfo , int intLen );
        /// <summary>Use a different I/O</summary>
        /// <param name="lngGSInstance"></param>
        /// <param name="gsdll_stdin">Function that menage the Standard INPUT</param>
        /// <param name="gsdll_stdout">Function that menage the Standard OUTPUT</param>
        /// <param name="gsdll_stderr">Function that menage the Standard ERROR output</param>
        /// <returns></returns>
        [DllImport("gsdll32.dll", EntryPoint = "gsapi_set_stdio")]
        private static extern int gsapi_set_stdio(IntPtr lngGSInstance, StdioCallBack gsdll_stdin, StdioCallBack gsdll_stdout, StdioCallBack gsdll_stderr);

        #endregion
        #region Const
        const int e_Quit=-101;
        const int e_NeedInput = -106;
        #endregion
        #region Variables
        private string _sDeviceFormat;
        private string _sParametersUsed;

        private int _iWidth;
        private int _iHeight;
        private int _iResolutionX;
        private int _iResolutionY;
        private int _iJPEGQuality;
        /// <summary>The first page to convert in image</summary>
        private int _iFirstPageToConvert = -1;
        /// <summary>The last page to conver in an image</summary>
        private int _iLastPageToConvert = -1;
        /// <summary>This parameter is used to control subsample antialiasing of graphics</summary>
        private int _iGraphicsAlphaBit = -1;
        /// <summary>This parameter is used to control subsample antialiasing of text</summary>
        private int _iTextAlphaBit = -1;
        /// <summary>In how many thread i should perform the conversion</summary>
        /// <remarks>This is a Major innovation since 8.63 NEVER use it with previous version!</remarks>
        private int _iRenderingThreads = -1;

        //set the path for the GS
        private string GS_Path = "";

        public string GSPath
        {
            set { GS_Path = value; }
        }

        /// <summary>In how many thread i should perform the conversion</summary>
        /// <remarks>This is a Major innovation since 8.63 NEVER use it with previous version!</remarks>
        /// <value>Set it to 0 made the program set it to Environment.ProcessorCount HT machine could want to perform a check for this..</value>
        public int RenderingThreads
        {
            get { return _iRenderingThreads; }
            set {
                if (value == 0)
                    _iRenderingThreads = Environment.ProcessorCount;
                else
                    _iRenderingThreads = value; }
        }
        private bool _bFitPage;
        private bool _bThrowOnlyException = false;
        private bool _bRedirectIO = false;
        private bool _bForcePageSize = false;
        /// <summary>The pagesize of the output</summary>
        private string _sDefaultPageSize;
        private IntPtr _objHandle;
        /// <summary>If true i will try to output everypage to a different file!</summary>
        private bool _didOutputToMultipleFile = false;

        private System.Diagnostics.Process myProcess;
        public StringBuilder output;
        //public string output;
        //private List<byte> outputBytes;
        //public string error;
        #region Fonts related variables thanks to devalv
        private List<string> _sFontPath = new List<string>();
        private bool _bDisablePlatformFonts = false;
        private bool _bDisableFontMap = false;
        private List<string> _sFontMap = new List<string>();
        private string _sSubstitutionFont;
        private string _sFCOFontFile;
        private string _sFAPIFontMap;
        private bool _bDisablePrecompiledFonts = false;
        #endregion
        #endregion
        #region Proprieties
        /// <summary>
        /// What format to use to convert
        /// is suggested to use png256 instead of jpeg for document!
        /// they are smaller and better suited!
        /// </summary>
        /// <see cref="http://pages.cs.wisc.edu/~ghost/doc/cvs/Devices.htm"/>
        public string OutputFormat
        {
            get { return _sDeviceFormat; }
            set { _sDeviceFormat = value; }
        }

        /// <summary>The pagesize of the output</summary>
        /// <remarks>Without this parameter the output should be letter, complain to USA for this :) if the document specify a different size it will take precedece over this!</remarks>
        /// <see cref="http://pages.cs.wisc.edu/~ghost/doc/cvs/Use.htm#Known_paper_sizes"/>
        public string DefaultPageSize
        {
            get { return _sDefaultPageSize; }
            set { _sDefaultPageSize = value; }
        }

        /// <summary>If set to true and page default page size will force the rendering in that output format</summary>
        public bool ForcePageSize
        {
            get { return _bForcePageSize; }
            set { _bForcePageSize = value; }
        }

        public string ParametersUsed
        {
            get { return _sParametersUsed; }
            set { _sParametersUsed = value; }
        }

        public int Width
        {
            get { return _iWidth; }
            set { _iWidth = value; }
        }

        public int Height
        {
            get { return _iHeight; }
            set { _iHeight = value; }
        }

        public int ResolutionX
        {
            get { return _iResolutionX; }
            set { _iResolutionX = value; }
        }

        public int ResolutionY
        {
            get { return _iResolutionY; }
            set { _iResolutionY = value; }
        }

        /// <summary>This parameter is used to control subsample antialiasing of graphics</summary>
        /// <value>Value MUST BE below or equal 0 if not set, or 1,2,or 4 NO OTHER VALUES!</value>
        /// <see cref="http://pages.cs.wisc.edu/~ghost/doc/cvs/Use.htm"/>
        public int GraphicsAlphaBit
        {
            get { return _iGraphicsAlphaBit; }
            set
            {
                if ((value > 4) | (value == 3))
                    throw new ArgumentOutOfRangeException("The Graphics Alpha Bit must have a value between 1 2 and 4, or <= 0 if not set");
                _iGraphicsAlphaBit = value;
            }
        }
        /// <summary>This parameter is used to control subsample antialiasing of text</summary>
        /// <value>Value MUST BE below or equal 0 if not set, or 1,2,or 4 NO OTHER VALUES!</value>
        /// <see cref="http://pages.cs.wisc.edu/~ghost/doc/cvs/Use.htm"/>
        public int TextAlphaBit
        {
            get { return _iTextAlphaBit; }
            set
            {
                if ((value > 4) | (value == 3))
                    throw new ArgumentOutOfRangeException("The Text Alpha Bit must have a value between 1 2 and 4, or <= 0 if not set");
                _iTextAlphaBit = value;
            }
        }

        public Boolean FitPage
        {
            get { return _bFitPage; }
            set { _bFitPage = value; }
        }
        /// <summary>Quality of compression of JPG</summary>
        public int JPEGQuality
        {
            get { return _iJPEGQuality; }
            set { _iJPEGQuality = value; }
        }
        /// <summary>The first page to convert in image</summary>
        public int FirstPageToConvert
        {
            get { return _iFirstPageToConvert; }
            set { _iFirstPageToConvert = value; }
        }
        /// <summary>The last page to conver in an image</summary>
        public int LastPageToConvert
        {
            get { return _iLastPageToConvert; }
            set { _iLastPageToConvert = value; }
        }
        /// <summary>Set to True if u want the program to never display Messagebox
        /// but otherwise throw exception</summary>
        public Boolean ThrowOnlyException
        {
            get { return _bThrowOnlyException; }
            set { _bThrowOnlyException = value; }
        }
        /// <summary>[NOT WORKING]If i should redirect the Output of Ghostscript library somewhere</summary>
        /// <remarks>NOT WORKING!!!</remarks>
        public bool RedirectIO
        {
            get { return _bRedirectIO; }
            set { _bRedirectIO = value; }
        }
        /// <summary>If true i will try to output everypage to a different file!</summary>
        public bool OutputToMultipleFile
        {
            get { return _didOutputToMultipleFile; }
            set { _didOutputToMultipleFile = value; }
        }
        /// <summary>
        /// If set to true the library will use a mutex to be sure that the library is never called twice at the same time!
        /// </summary>
        public bool UseMutex
        {
            get { return mutex != null; }
            set {
                if (!value)//if i don't want to use it
                {
                    if (mutex != null)//if it exist
                    {   //close and delete it
                        mutex.ReleaseMutex();
                        mutex.Close();
                        mutex = null;
                    }
                }
                else//If i want to use mutex create it if it doesn't exist
                {
                    if (mutex == null)
                        mutex = new System.Threading.Mutex(false, "MutexGhostscript");
                }
            }
        }

        #region Fonts related thanks to devalv
        public List<string> FontPath
        {
            get { return _sFontPath; }
            set { _sFontPath = value; }
        }

        public bool DisablePlatformFonts
        {
            get { return _bDisablePlatformFonts; }
            set { _bDisablePlatformFonts = value; }
        }

        public bool DisableFontMap
        {
            get { return _bDisableFontMap; }
            set { _bDisableFontMap = value; }
        }

        public List<string> FontMap
        {
            get { return _sFontMap; }
            set { _sFontMap = value; }
        }

        public string SubstitutionFont
        {
            get { return _sSubstitutionFont; }
            set { _sSubstitutionFont = value; }
        }

        public string FCOFontFile
        {
            get { return _sFCOFontFile; }
            set { _sFCOFontFile = value; }
        }

        public string FAPIFontMap
        {
            get { return _sFAPIFontMap; }
            set { _sFAPIFontMap = value; }
        }

        public bool DisablePrecompiledFonts
        {
            get { return _bDisablePrecompiledFonts; }
            set { _bDisablePrecompiledFonts = value; }
        }
        #endregion
        #endregion
        #region Init
        public PDFConvert(IntPtr objHandle)
        {
            _objHandle = objHandle;
        }

        public PDFConvert()
        {
            _objHandle = IntPtr.Zero;
        }
        #endregion

        #region Convert
        /// <summary>Convert a single file!</summary>
        /// <param name="inputFile">The file PDf to convert</param>
        /// <param name="outputFile">The image file that will be created</param>
        /// <remarks>You must pass all the parameter for the conversion
        /// as Proprieties of this class</remarks>
        /// <returns>True if the conversion succed!</returns>
        public bool Convert(string inputFile, string outputFile)
        {
            return Convert(inputFile, outputFile, _bThrowOnlyException,null);
        }

        /// <summary>Convert a single file!</summary>
        /// <param name="inputFile">The file PDf to convert</param>
        /// <param name="outputFile">The image file that will be created</param>
        /// <param name="parameters">You must pass all the parameter for the conversion here</param>
        /// <remarks>Thanks to 	tchu_2000 for the help!</remarks>
        /// <returns>True if the conversion succed!</returns>
        public bool Convert(string inputFile, string outputFile,string parameters)
        {
            return Convert(inputFile, outputFile, _bThrowOnlyException, parameters);
        }

        /// <summary>Convert a single file!</summary>
        /// <param name="inputFile">The file PDf to convert</param>
        /// <param name="outputFile">The image file that will be created</param>
        /// <param name="throwException">if the function should throw an exception
        /// or display a message box</param>
        /// <remarks>You must pass all the parameter for the conversion
        /// as Proprieties of this class</remarks>
        /// <returns>True if the conversion succed!</returns>
        private bool Convert(string inputFile, string outputFile,bool throwException,string options)
        {
            #region Check Input
            //Avoid to work when the file doesn't exist
            if (string.IsNullOrEmpty(inputFile))
                throw new ArgumentNullException("inputFile");
            if (!System.IO.File.Exists(inputFile))
                throw new ArgumentException(string.Format("The file :'{0}' doesn't exist", inputFile), "inputFile");
            if (string.IsNullOrEmpty(_sDeviceFormat))
                throw new ArgumentNullException("Device");
            //be sure that if i specify multiple page outpage i added the % to the filename!
            #endregion
            //If i create a Mutex it means i want to protect concurrent access to the library
            if (mutex != null) mutex.WaitOne();
            bool result = false;
            try
            {
                result = ExecuteGhostscriptCommand(GetGeneratedArgs(inputFile, outputFile, options));
            }
            finally {if (mutex != null) mutex.ReleaseMutex();}
            return result;
        }

        /// <summary>Print a file</summary>
        /// <param name="inputFile">THe file to print</param>
        /// <param name="printParametersFile">The file with the configuration of the printer</param>
        /// <returns>True if i send the work to the printer queue</returns>
        public bool Print(string inputFile, string printParametersFile)
        {
            #region Check Input
            //Avoid to work when the file doesn't exist
            if (string.IsNullOrEmpty(inputFile))
                throw new ArgumentNullException("inputFile");
            if (!System.IO.File.Exists(inputFile))
                throw new ArgumentException(string.Format("The file :'{0}' doesn't exist", inputFile), "inputFile");
            //Avoid to work when the file doesn't exist
            if (string.IsNullOrEmpty(printParametersFile))
                throw new ArgumentNullException("printParametersFile");
            if (!System.IO.File.Exists(printParametersFile))
                throw new ArgumentException(string.Format("The file :'{0}' doesn't exist", printParametersFile), "printParametersFile");
            #endregion
            // Example : gswin32.exe" -dNOPAUSE -dBATCH -dFirstPage=1 -dLastPage=1 setup.ps mio.pdf -c quit
            List<string> args = new List<string>(7);
            args.Add("printPdf");
            args.Add("-dNOPAUSE");
            args.Add("-dBATCH");
            if (_iFirstPageToConvert > 0)
                args.Add(string.Format("-dFirstPage={0}", _iFirstPageToConvert));
            if ((_iLastPageToConvert > 0) && (_iLastPageToConvert >= _iFirstPageToConvert))
                args.Add(string.Format("-dLastPage={0}", _iLastPageToConvert));
            args.Add(printParametersFile);
            args.Add(inputFile);
            bool result = false;
            if (mutex != null) mutex.WaitOne();
            try{result = ExecuteGhostscriptCommand(args.ToArray());}
            finally { if (mutex != null) mutex.ReleaseMutex(); }
            return result;
        }

        /// <summary>Execute a Ghostscript command with a certain list of arguments</summary>
        /// <param name="sArgs">The list of the arguments</param>
        /// <returns>true if it succed, false otherwise</returns>
        private bool ExecuteGhostscriptCommand(string[] sArgs)
        {
            //Console.WriteLine("gs: {0}", string.Join(" ", sArgs));

            #region Variables
            int intReturn, intCounter, intElementCount;
            //The pointer to the current istance of the dll
            IntPtr intGSInstanceHandle = IntPtr.Zero;
            object[] aAnsiArgs;
            IntPtr[] aPtrArgs;
            GCHandle[] aGCHandle;
            IntPtr callerHandle, intptrArgs;
            GCHandle gchandleArgs;
            #endregion
            #region Convert Unicode strings to null terminated ANSI byte arrays
            // Convert the Unicode strings to null terminated ANSI byte arrays
            // then get pointers to the byte arrays.
            intElementCount = sArgs.Length;
            aAnsiArgs = new object[intElementCount];
            aPtrArgs = new IntPtr[intElementCount];
            aGCHandle = new GCHandle[intElementCount];
            //Convert the parameters
            for (intCounter = 0; intCounter < intElementCount; intCounter++)
            {
                aAnsiArgs[intCounter] = StringToAnsiZ(sArgs[intCounter]);
                aGCHandle[intCounter] = GCHandle.Alloc(aAnsiArgs[intCounter], GCHandleType.Pinned);
                aPtrArgs[intCounter] = aGCHandle[intCounter].AddrOfPinnedObject();
            }
            gchandleArgs = GCHandle.Alloc(aPtrArgs, GCHandleType.Pinned);
            intptrArgs = gchandleArgs.AddrOfPinnedObject();
            #endregion
            #region Create a new istance of the library!
            intReturn = -1;
            try
            {
                intReturn = gsapi_new_instance(out intGSInstanceHandle, _objHandle);
                //Be sure that we create an istance!
                if (intReturn < 0)
                {
                    ClearParameters(ref aGCHandle, ref gchandleArgs);
                    throw new ApplicationException("I can't create a new istance of Ghostscript please verify no other istance are running!");
                }
            }
            catch (BadImageFormatException )//99.9% of time i'm just loading a 32bit dll in a 64bit enviroment!
            {
                ClearParameters(ref aGCHandle, ref gchandleArgs);
                //Check if i'm in a 64bit enviroment or a 32bit
                if (IntPtr.Size == 8) // 8 * 8 = 64
                {
                    throw new ApplicationException(string.Format("The gsdll32.dll you provide is not compatible with the current architecture that is 64bit,"+
                    "Please download any version above version 8.64 from the original website in the 64bit or x64 or AMD64 version!"));
                }
                else if (IntPtr.Size == 4) // 4 * 8 = 32
                {
                    throw new ApplicationException(string.Format("The gsdll32.dll you provide is not compatible with the current architecture that is 32bit,"+
                    "Please download any version above version 8.64 from the original website in the 32bit or x86 or i386 version!"));
                }
            }
            catch (DllNotFoundException )//in this case the dll we r using isn't the dll we expect
            {
                ClearParameters(ref aGCHandle, ref gchandleArgs);
                throw new ApplicationException("The gsdll32.dll wasn't found in default dlls search path" +
                    "or is not in correct version (doesn't expose the required methods). Please download " +
                    "at least the version 8.64 from the original website");
            }
            callerHandle = IntPtr.Zero;//remove unwanter handler
            #endregion
            #region Capture the I/O
            //Not working
            if (_bRedirectIO)
            {
                StdioCallBack stdinCallback = new StdioCallBack(gsdll_stdin);
                StdioCallBack stdoutCallback = new StdioCallBack(gsdll_stdout);
                StdioCallBack stderrCallback = new StdioCallBack(gsdll_stderr);
                intReturn = gsapi_set_stdio(intGSInstanceHandle, stdinCallback, stdoutCallback, stderrCallback);
                if (output == null) output = new StringBuilder();
                else output.Remove(0, output.Length);
                myProcess = System.Diagnostics.Process.GetCurrentProcess();
                myProcess.OutputDataReceived += new System.Diagnostics.DataReceivedEventHandler(SaveOutputToImage);
            }
            #endregion
            intReturn = -1;//if nothing change it there is an error!
            //Ok now is time to call the interesting method
            try { intReturn = gsapi_init_with_args(intGSInstanceHandle, intElementCount, intptrArgs); }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message, ex);
            }
            finally//No matter what happen i MUST close the istance!
            {   //free all the memory
                ClearParameters(ref aGCHandle, ref gchandleArgs);
                gsapi_exit(intGSInstanceHandle);//Close the istance
                gsapi_delete_instance(intGSInstanceHandle);//delete it
                //In case i was looking for output now stop
                if ((myProcess != null) && (_bRedirectIO)) myProcess.OutputDataReceived -= new System.Diagnostics.DataReceivedEventHandler(SaveOutputToImage);
            }
            //Conversion was successfull if return code was 0 or e_Quit
            return (intReturn == 0) | (intReturn == e_Quit);//e_Quit = -101
        }

        /// <summary>Remove the memory allocated</summary>
        /// <param name="aGCHandle"></param>
        /// <param name="gchandleArgs"></param>
        private void ClearParameters(ref GCHandle[] aGCHandle, ref GCHandle gchandleArgs)
        {
            for (int intCounter = 0; intCounter < aGCHandle.Length; intCounter++)
                aGCHandle[intCounter].Free();
            gchandleArgs.Free();
        }
        #region Test (code not used)
        void SaveOutputToImage(object sender, System.Diagnostics.DataReceivedEventArgs e)
        {
            output.Append(e.Data);
        }

        //public System.Drawing.Image Convert(string inputFile)
        //{
        //    _bRedirectIO = true;
        //    if (Convert(inputFile, "%stdout", _bThrowOnlyException))
        //    {
        //        if ((output != null) && (output.Length > 0))
        //        {
        //            //StringReader sr = new StringReader(output.ToString());
        //            //MemoryStream ms = new MemoryStream(UTF8Encoding.Default.GetBytes(output.ToString()));
        //            System.Drawing.Image returnImage = System.Drawing.Image.FromStream(myProcess.StandardOutput.BaseStream).Clone() as System.Drawing.Image;
        //            //ms.Close();
        //            return returnImage;
        //        }
        //    }
        //    return null;
        //}
        #endregion
        #endregion

        #region Accessory Functions
        /// <summary>This function create the list of parameters to pass to the dll with parameters given directly from the program</summary>
        /// <param name="inputFile"></param>
        /// <param name="outputFile"></param>
        /// <param name="otherParameters">The other parameters i could be interested</param>
        /// <remarks>Be very Cautious using this! code provided and modified from tchu_2000</remarks>
        /// <returns></returns>
        private string[] GetGeneratedArgs(string inputFile, string outputFile, string otherParameters)
        {
            if (!string.IsNullOrEmpty(otherParameters))
                return GetGeneratedArgs(inputFile,outputFile,otherParameters.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
            else
                return GetGeneratedArgs(inputFile,outputFile,(string[])null);
        }

        /// <summary>This function create the list of parameters to pass to the dll</summary>
        /// <param name="inputFile">the file to convert</param>
        /// <param name="outputFile">where to write the image</param>
        /// <returns>the list of the arguments</returns>
        private string[] GetGeneratedArgs(string inputFile, string outputFile,string[] presetParameters)
        {
            string[] args;
            ArrayList lstExtraArgs = new ArrayList();
            //ok if i haven't been passed a list of parameters create my own
            if ((presetParameters == null) || (presetParameters.Length == 0))
            {
                #region Parameters
                //Ok now check argument per argument and compile them
                //If i want a jpeg i can set also quality
                if (_sDeviceFormat == "jpeg" && _iJPEGQuality > 0 && _iJPEGQuality < 101)
                    lstExtraArgs.Add(string.Format(GS_JpegQualityFormat, _iJPEGQuality));
                //if i provide size it will override the paper size
                if (_iWidth > 0 && _iHeight > 0)
                    lstExtraArgs.Add(string.Format(GS_PageSizeFormat, _iWidth, _iHeight));
                else//otherwise if aviable use the papersize
                {
                    if (!string.IsNullOrEmpty(_sDefaultPageSize))
                    {
                        lstExtraArgs.Add(string.Format(GS_DefaultPaperSize, _sDefaultPageSize));
                        //It have no meaning to set it if the default page is not set!
                        if (_bForcePageSize)
                            lstExtraArgs.Add(GS_FixedMedia);
                    }
                }

                //not set antialiasing settings
                if (_iGraphicsAlphaBit > 0)
                    lstExtraArgs.Add(string.Format(GS_GraphicsAlphaBits, _iGraphicsAlphaBit));
                if (_iTextAlphaBit > 0)
                    lstExtraArgs.Add(string.Format(GS_TextAlphaBits, _iTextAlphaBit));
                //Should i try to fit?
                if (_bFitPage) lstExtraArgs.Add(GS_FitPage);
                //Do i have a forced resolution?
                if (_iResolutionX > 0)
                {
                    if (_iResolutionY > 0)
                        lstExtraArgs.Add(String.Format(GS_ResolutionXYFormat, _iResolutionX, _iResolutionY));
                    else
                        lstExtraArgs.Add(String.Format(GS_ResolutionXFormat, _iResolutionX));
                }
                if (_iFirstPageToConvert > 0)
                    lstExtraArgs.Add(String.Format(GS_FirstPageFormat, _iFirstPageToConvert));
                if (_iLastPageToConvert > 0)
                {
                    if ((_iFirstPageToConvert > 0) && (_iFirstPageToConvert > _iLastPageToConvert))
                        throw new ArgumentOutOfRangeException(string.Format("The 1st page to convert ({0}) can't be after then the last one ({1})",_iFirstPageToConvert,_iLastPageToConvert));
                    lstExtraArgs.Add(String.Format(GS_LastPageFormat, _iLastPageToConvert));
                }
                //Set in how many threads i want to do the work
                if (_iRenderingThreads > 0)
                    lstExtraArgs.Add(String.Format(GS_RenderingThreads, _iRenderingThreads));

                //If i want to redirect write it to the standard output!
                if (_bRedirectIO)
                {
                    //In this case you must also use the -q switch to prevent Ghostscript
                    //from writing messages to standard output which become
                    //mixed with the intended output stream. 
                    //outputFile = GS_StandardOutputDevice;
                    //lstExtraArgs.Add(GS_QuiteOperation);
                }
                #region Fonts
                if ((_sFontPath != null) && (_sFontPath.Count > 0))
                    lstExtraArgs.Add(String.Format(GS_FontPath, String.Join(";", _sFontPath.ToArray())));
                if (_bDisablePlatformFonts)
                    lstExtraArgs.Add(GS_NoPlatformFonts);
                if (_bDisableFontMap)
                    lstExtraArgs.Add(GS_NoFontMap);
                if ((_sFontMap != null) && (_sFontMap.Count > 0))
                    lstExtraArgs.Add(String.Format(GS_FontMap, String.Join(";", _sFontMap.ToArray())));
                if (!string.IsNullOrEmpty(_sSubstitutionFont))
                    lstExtraArgs.Add(string.Format(GS_SubstitutionFont, _sSubstitutionFont));
                if (!string.IsNullOrEmpty(_sFCOFontFile))
                    lstExtraArgs.Add(string.Format(GS_FCOFontFile, _sFCOFontFile));
                if (!string.IsNullOrEmpty(_sFAPIFontMap))
                {
                    lstExtraArgs.Add(string.Format(GS_FAPIFontMap, _sFAPIFontMap));
                }
                if (_bDisablePrecompiledFonts)
                    lstExtraArgs.Add(GS_NoPrecompiledFonts);
                #endregion

                //the following line is added by me to force the system to look for gs-related files in my own specified folder
                if (!string.IsNullOrEmpty(GS_Path))
                {
                    lstExtraArgs.Add(string.Format("-I{0}", GS_Path));
                }

                #endregion
                int iFixedCount = 7;//This are the mandatory options
                int iExtraArgsCount = lstExtraArgs.Count;
                args = new string[iFixedCount + lstExtraArgs.Count];
                args[1] = GS_Fixed1stParameter;//"-dNOPAUSE";//I don't want interruptions
                args[2] = GS_Fixed2ndParameter;//"-dBATCH";//stop after
                args[3] = GS_Fixed3rdParameter;//"-dSAFER";
                args[4] = string.Format(GS_DeviceFormat, _sDeviceFormat);//what kind of export format i should provide
                //For a complete list watch here:
                //http://pages.cs.wisc.edu/~ghost/doc/cvs/Devices.htm
                //Fill the remaining parameters
                for (int i = 0; i < iExtraArgsCount; i++)
                    args[5 + i] = (string)lstExtraArgs[i];
            }
            else
            {//3 arguments MUST be added 0 (meaningless) and at the end the output and the inputfile
                args = new string[presetParameters.Length + 3];
                //now use the parameters i receive (thanks CrucialBT to point this out!)
                //and thanks to Barbara who pointout that i was skipping the last parameter
                for (int i = 1; i <= presetParameters.Length; i++)
                    args[i] = presetParameters[i - 1];
            }
            args[0]=GS_FirstParameter;//this parameter have little real use
            //Now check if i want to update to 1 file per page i have to be sure do add % to the output filename
            if ((_didOutputToMultipleFile) && (!outputFile.Contains(GS_MultiplePageCharacter)))
            {// Thanks to Spillie to show me the error!
                int lastDotIndex = outputFile.LastIndexOf('.');
                if (lastDotIndex > 0)
                    outputFile = outputFile.Insert(lastDotIndex, "%d");
            }
            //Ok now save them to be shown 4 debug use
            _sParametersUsed = string.Empty;
            //Copy all the args except the 1st that is useless and the last 2
            for (int i = 1; i < args.Length - 2; i++)
                _sParametersUsed += " " + args[i];
            //Fill outputfile and inputfile as last 2 arguments!
            args[args.Length - 2] = string.Format(GS_OutputFileFormat, outputFile);
            args[args.Length - 1] = string.Format("{0}", inputFile);

            _sParametersUsed += " " + string.Format(GS_OutputFileFormat, string.Format("\"{0}\"", outputFile))
            + " " + string.Format("\"{0}\"", inputFile);
            return args;
        }

        /// <summary>
        /// Convert a Unicode string to a null terminated Ansi string for Ghostscript.
        /// The result is stored in a byte array
        /// </summary>
        /// <param name="str">The parameter i want to convert</param>
        /// <returns>the byte array that contain the string</returns>
        private static byte[] StringToAnsiZ(string str)
        {   //This with Encoding.Default should work also with Chineese Japaneese
            //Thanks to tchu_2000 I18N related patch
            if (str == null) str = String.Empty;
            return Encoding.Default.GetBytes(str);
        }

        /// <summary>Convert a Pointer to a string to a real string</summary>
        /// <param name="strz">the pointer to the string in memory</param>
        /// <returns>The string</returns>
        public static string AnsiZtoString(IntPtr strz)
        {
            if (strz != IntPtr.Zero)
                return Marshal.PtrToStringAnsi(strz);
            else
                return string.Empty;
        }

        /// <summary>Check if i find the DLL that i need to continue!</summary>
        /// <returns>true if i found it</returns>
        public static bool CheckDll()
        {
            return File.Exists(GhostScriptDLLName);
        }
        #endregion
        #region Menage Standard Input & Standard Output
        public int gsdll_stdin(IntPtr intGSInstanceHandle, IntPtr strz, int intBytes)
        {
            // This is dumb code that reads one byte at a time
           // Ghostscript doesn't mind this, it is just very slow
            if (intBytes == 0) 
                return 0;
            else
            {
                int ich = Console.Read();
                if (ich == -1)
                    return 0; // EOF
                else
                {
                    byte bch = (byte)ich;
                    GCHandle gcByte = GCHandle.Alloc(bch, GCHandleType.Pinned);
                    IntPtr ptrByte = gcByte.AddrOfPinnedObject();
                    CopyMemory(strz, ptrByte, 1);
                    ptrByte = IntPtr.Zero;
                    gcByte.Free();
                    return 1;
                }
            }
        }

        public int gsdll_stdout(IntPtr intGSInstanceHandle, IntPtr strz , int intBytes)
        {
            // If you can think of a more efficient method, please tell me!
            // We need to convert from a byte buffer to a string
            // First we create a byte array of the appropriate size
            byte[] aByte = new byte[intBytes];
            // Then we get the address of the byte array
            GCHandle gcByte = GCHandle.Alloc(aByte, GCHandleType.Pinned);
            IntPtr ptrByte = gcByte.AddrOfPinnedObject();
            // Then we copy the buffer to the byte array
            CopyMemory(ptrByte, strz, (uint)intBytes);
            // Release the address locking
            ptrByte = IntPtr.Zero;
            gcByte.Free();
            // Then we copy the byte array to a string, character by character
            string str = "";
            for (int i = 0; i < intBytes; i++)
            {
                str += (char)aByte[i];
            }
            // Finally we output the message
            //Console.Write(str);
            output.Append(str);
            return intBytes;
            //if (intBytes > 0)
            //{
            //    Console.Write(Marshal.PtrToStringAnsi(strz));
            //}
            //return 0;
        }

        public int gsdll_stderr(IntPtr intGSInstanceHandle, IntPtr strz, int intBytes )
        {
            return gsdll_stdout(intGSInstanceHandle, strz, intBytes);
            //Console.Write(Marshal.PtrToStringAnsi(strz));
            //return intBytes;
        }
        #endregion
        #region Menage Revision
        public GhostScriptRevision GetRevision()
        {
            // Check revision number of Ghostscript
            int intReturn;
            GS_Revision udtGSRevInfo = new GS_Revision();
            GhostScriptRevision output;
            GCHandle gcRevision;
            gcRevision = GCHandle.Alloc(udtGSRevInfo, GCHandleType.Pinned);
            intReturn = gsapi_revision(ref udtGSRevInfo, 16);
            output.intRevision = udtGSRevInfo.intRevision;
            output.intRevisionDate = udtGSRevInfo.intRevisionDate;
            output.ProductInformation = AnsiZtoString(udtGSRevInfo.strProduct);
            output.CopyrightInformations = AnsiZtoString(udtGSRevInfo.strCopyright);
            gcRevision.Free();
            return output;
        }
        #endregion
    }

    /// <summary>Delegate used by Ghostscript to perform I/O operations</summary>
    /// <param name="handle"></param>
    /// <param name="strptr"></param>
    /// <param name="count"></param>
    /// <returns></returns>
    public delegate int StdioCallBack(IntPtr handle, IntPtr strptr, int count);

    /// <summary>This struct is filled with the information of the version of this ghostscript</summary>
    /// <remarks>Have the layout defined cuz i will fill it with a kernel copy memory</remarks>
    [StructLayout(LayoutKind.Sequential)]
    struct GS_Revision
    {
        public IntPtr strProduct;
        public IntPtr strCopyright;
        public int intRevision;
        public int intRevisionDate;
    }

    public struct GhostScriptRevision
    {
        public string ProductInformation;
        public string CopyrightInformations;
        public int intRevision;
        public int intRevisionDate;
    }
}
