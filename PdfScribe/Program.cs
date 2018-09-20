using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Reflection;
using PdfScribeCore;
namespace PdfScribe
{
 
    public class Program
    {

        #region Message constants

        const string errorDialogCaption = "PDF Scribe"; // Error taskdialog caption text
        public static string CurrentDir=""; //Take Empty Variable for Current Directory
        public static string instance = "";
        const string errorDialogInstructionPDFGeneration = "There was a PDF generation error.";
        const string errorDialogInstructionCouldNotWrite = "Could not create the output file.";
        const string errorDialogInstructionUnexpectedError = "There was an internal error. Enable tracing for details.";

        const string errorDialogTextFileInUse = "{0} is being used by another process.";
        const string errorDialogTextGhostScriptConversion = "Ghostscript error code {0}.";

        const string warnFileNotDeleted = "{0} could not be deleted.";

        #endregion

        #region Other constants
        const string traceSourceName = "PdfScribe";

        const string defaultOutputFilename = "PDFSCRIBE.PDF";

        #endregion

        static TraceSource logEventSource = new TraceSource(traceSourceName);
        [STAThread]
        static void Main(string[] args)
        {
            // Install the global exception handler
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(Application_UnhandledException);


            String standardInputFilename = Path.GetTempFileName();
            String outputFilename = String.Empty;
            String outputfilepath = String.Empty;
            String path1 = String.Empty;

            //Default Values for Printer page and Color
            string papersize = "A0";
            string DEVICEWIDTHPOINTS = "2384";
            string DEVICEHEIGHTPOINTS = "3370";
            string color = "Color";

            String[] ghostScriptArguments_o;

            try
            {
                using (BinaryReader standardInputReader = new BinaryReader(Console.OpenStandardInput()))
                {
                    using (FileStream standardInputFile = new FileStream(standardInputFilename, FileMode.Create, FileAccess.ReadWrite))
                    {
                        standardInputReader.BaseStream.CopyTo(standardInputFile);
                    }
                }

                if (GetPdfOutputFilename(ref outputFilename))
                {
                    try
                    {
                        // Get Installed Instance Path And Read App.config File
                        string loc = System.Reflection.Assembly.GetExecutingAssembly().Location;
                        string Config = Path.Combine(Path.GetDirectoryName(loc) + "\\App.config");
                        System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
                        string fileName = @Config;
                        xmlDoc.Load(fileName);
                        CurrentDir = xmlDoc["configuration"]["userSettings"]["PdfScribe.Properties.Settings"]["setting"]["value"].InnerText;

                        instance = CurrentDir;
                        string folder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        string specificFolder = Path.Combine(folder, "PdfScribe");
                        System.Xml.Linq.XDocument doc = new System.Xml.Linq.XDocument();
                        //MessageBox.Show("Instance = " + instance);
                        if (instance == "" || instance == ":instance1")
                        {

                            string printer1 = Path.Combine(specificFolder, "scribe");
                            doc = System.Xml.Linq.XDocument.Load(@printer1 + "\\" + "config.xml");
                        }

                        if (instance == ":instance2")
                        {
                            string printer2 = Path.Combine(specificFolder, "scribe2");
                            doc = System.Xml.Linq.XDocument.Load(@printer2 + "\\" + "config.xml");
                        }

                        if (instance == ":instance3")
                        {
                            string printer3 = Path.Combine(specificFolder, "scribe3");
                            doc = System.Xml.Linq.XDocument.Load(@printer3 + "\\" + "config.xml");
                        }
                        if (instance == ":instance4")
                        {
                            string printer4 = Path.Combine(specificFolder, "scribe4");
                            doc = System.Xml.Linq.XDocument.Load(@printer4 + "\\" + "config.xml");
                        }

                        if (instance == ":instance5")
                        {
                            string printer5 = Path.Combine(specificFolder, "scribe5");
                            doc = System.Xml.Linq.XDocument.Load(@printer5 + "\\" + "config.xml");
                        }

                        if (instance == ":instance6")
                        {
                            string printer6 = Path.Combine(specificFolder, "scribe6");
                            doc = System.Xml.Linq.XDocument.Load(@printer6 + "\\" + "config.xml");
                        }

                        if (instance == ":instance7")
                        {
                            string printer7 = Path.Combine(specificFolder, "scribe7");
                            doc = System.Xml.Linq.XDocument.Load(@printer7 + "\\" + "config.xml");
                        }

                        if (instance == ":instance8")
                        {
                            string printer8 = Path.Combine(specificFolder, "scribe8");
                            doc = System.Xml.Linq.XDocument.Load(@printer8 + "\\" + "config.xml");
                        }

                        if (instance == ":instance9")
                        {
                            string printer9 = Path.Combine(specificFolder, "scribe9");
                            doc = System.Xml.Linq.XDocument.Load(@printer9 + "\\" + "config.xml");
                        }

                        if (instance == ":instance10")
                        {
                            string printer10 = Path.Combine(specificFolder, "scribe10");
                            doc = System.Xml.Linq.XDocument.Load(@printer10 + "\\" + "config.xml");
                        }

                        //Read Config.xml file in Current Working Directory
                        //System.Xml.Linq.XDocument doc = new System.Xml.Linq.XDocument();
                        //doc = System.Xml.Linq.XDocument.Load(CurrentDir + "\\" + "config.xml");
                        var parameters = doc.Descendants("Parameter").ToDictionary(n => n.Attribute("Name").Value, v => v.Attribute("Value").Value);
                        
                        if (parameters.Any())
                        {
                            //Set the Wix Properties in the Session object from the XML file
                            foreach (var parameter in parameters)
                            {
                                if (parameter.Key == "outputFileName")
                                {
                                    outputFilename = parameter.Value;
                                }
                                if (parameter.Key == "outputFilePath")
                                {
                                    outputfilepath = parameter.Value;
                                }
                                if (parameter.Key == "color")
                                {
                                    color = parameter.Value;
                                }
                                if (parameter.Key == "paperSize")
                                {
                                    papersize = parameter.Value;
                                }
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    

                    //Logic to Append _1...._2.. to pdf file name
                    int count = 1;
                    outputFilename = outputfilepath + "\\" + outputFilename;
                    string fileNameOnly = Path.GetFileNameWithoutExtension(outputFilename);

                    string extension = Path.GetExtension(outputFilename);
                    string path = Path.GetDirectoryName(outputFilename);
                    string newFullPath = outputFilename;
                    while (File.Exists(newFullPath))
                    {
                        string tempFileName = string.Format("{0}_{1}", fileNameOnly, count++);
                        newFullPath = Path.Combine(path, tempFileName + extension);
                    }

                    outputFilename = newFullPath;
                    
                    // Remove the existing PDF file if present
                    //File.Delete(outputFilename);


                    // Only set absolute minimum parameters, let the postscript input
                    // dictate as much as possible
                    //"-dDEVICEWIDTHPOINTS=8","-dDEVICEHEIGHTPOINTS=7","-sColor=1"

                    /*Original//String[] ghostScriptArguments = { "-dBATCH", "-dNOPAUSE", "-dSAFER",  "-sDEVICE=pdfwrite",
                                                String.Format("-sOutputFile={0}", outputFilename), standardInputFilename };*/

                  
                    //Setup Height And Width Points according to selected page size
                    if (papersize == "A0")
                    {
                        DEVICEWIDTHPOINTS = "2384";
                        DEVICEHEIGHTPOINTS = "3370";

                    }

                    if (papersize == "A1")
                    {
                        DEVICEWIDTHPOINTS = "1684";
                        DEVICEHEIGHTPOINTS = "2384";

                    }

                    if (papersize == "A2")
                    {
                        DEVICEWIDTHPOINTS = "1191";
                        DEVICEHEIGHTPOINTS = "1684";

                    }

                    if (papersize == "A3")
                    {
                        DEVICEWIDTHPOINTS = "842";
                        DEVICEHEIGHTPOINTS = "1191";

                    }

                    if (papersize == "A4")
                    {
                        DEVICEWIDTHPOINTS = "595";
                        DEVICEHEIGHTPOINTS = "842";
                    }

                    if (papersize == "A5")
                    {
                        DEVICEWIDTHPOINTS = "420";
                        DEVICEHEIGHTPOINTS = "595";

                    }

                    if (papersize == "Letter")
                    {
                        DEVICEWIDTHPOINTS = "612";
                        DEVICEHEIGHTPOINTS = "792";

                    }
                    //Tabloid		 792x1224
                    if (papersize == "Tabloid")
                    {
                        DEVICEWIDTHPOINTS = "792";
                        DEVICEHEIGHTPOINTS = "1224";

                    }

                    //Ledger		1224x792
                    if (papersize == "Ledger")
                    {
                        DEVICEWIDTHPOINTS = "1224";
                        DEVICEHEIGHTPOINTS = "792";

                    }
                    //Legal		 612x1008
                    if (papersize == "Legal")
                    {
                        DEVICEWIDTHPOINTS = "612";
                        DEVICEHEIGHTPOINTS = "1008";

                    }
                    //Statement	 396x612
                    if (papersize == "Statement")
                    {
                        DEVICEWIDTHPOINTS = "396";
                        DEVICEHEIGHTPOINTS = "612";

                    }
                    //Executive	 540x720

                    if (papersize == "Executive")
                    {
                        DEVICEWIDTHPOINTS = "540";
                        DEVICEHEIGHTPOINTS = "720";

                    }
                    //B4		 729x1032
                    if (papersize == "B4")
                    {
                        DEVICEWIDTHPOINTS = "729";
                        DEVICEHEIGHTPOINTS = "1032";

                    }

                    //B5		 516x729
                    if (papersize == "B5")
                    {
                        DEVICEWIDTHPOINTS = "516";
                        DEVICEHEIGHTPOINTS = "729";

                    }

                    //Prepare GhostScript Argument Based on parameters.
                    if (color == "Color")
                    {
                        String[] ghostScriptArguments = { "-dBATCH", "-dNOPAUSE", "-dSAFER",  "-sDEVICE=pdfwrite",
                                                String.Format("-sOutputFile={0}", outputFilename),String.Format("-sDEFAULTPAPERSIZE={0}", papersize),String.Format("-dDEVICEWIDTHPOINTS={0}", DEVICEWIDTHPOINTS),String.Format("-dDEVICEHEIGHTPOINTS={0}", DEVICEHEIGHTPOINTS),"-sPAPERSIZE=legal","-dFIXEDMEDIA","-dFitPage", standardInputFilename };
                        ghostScriptArguments_o = ghostScriptArguments;

                    }
                    else
                    {
                        String[] ghostScriptArguments = { "-dBATCH", "-dNOPAUSE", "-dSAFER",  "-sDEVICE=pdfwrite","-sColorConversionStrategy=Gray","-dProcessColorModel=/DeviceGray",
                                                String.Format("-sOutputFile={0}", outputFilename),String.Format("-sDEFAULTPAPERSIZE={0}", papersize),String.Format("-dDEVICEWIDTHPOINTS={0}", DEVICEWIDTHPOINTS),String.Format("-dDEVICEHEIGHTPOINTS={0}", DEVICEHEIGHTPOINTS),"-sPAPERSIZE=legal","-dFIXEDMEDIA","-dFitPage", standardInputFilename };
                        ghostScriptArguments_o = ghostScriptArguments;

                    }
                    GhostScript64.CallAPI(ghostScriptArguments_o);
                    //MessageBox.Show("File Saved as " + outputFilename);
                }
            }
            catch (IOException ioEx)
            {
                // We couldn't delete, or create a file
                // because it was in use
                logEventSource.TraceEvent(TraceEventType.Error, 
                                          (int)TraceEventType.Error,
                                          errorDialogInstructionCouldNotWrite +
                                          Environment.NewLine +
                                          "Exception message: " + ioEx.Message);
                DisplayErrorMessage(errorDialogCaption,
                                    errorDialogInstructionCouldNotWrite + Environment.NewLine +
                                    String.Format("{0} is in use.", outputFilename));
            }
            catch (UnauthorizedAccessException unauthorizedEx)
            {
                // Couldn't delete a file
                // because it was set to readonly
                // or couldn't create a file
                // because of permissions issues
                logEventSource.TraceEvent(TraceEventType.Error, 
                                          (int)TraceEventType.Error, 
                                          errorDialogInstructionCouldNotWrite +
                                          Environment.NewLine +
                                          "Exception message: " + unauthorizedEx.Message);
                DisplayErrorMessage(errorDialogCaption,
                                    errorDialogInstructionCouldNotWrite + Environment.NewLine +
                                    String.Format("Insufficient privileges to either create or delete {0}", outputFilename));


            }
            catch (ExternalException ghostscriptEx)
            {
                // Ghostscript error
                logEventSource.TraceEvent(TraceEventType.Error, 
                                          (int)TraceEventType.Error, 
                                          String.Format(errorDialogTextGhostScriptConversion, ghostscriptEx.ErrorCode.ToString()) +
                                          Environment.NewLine +
                                          "Exception message: " + ghostscriptEx.Message);
                DisplayErrorMessage(errorDialogCaption,
                                    errorDialogInstructionPDFGeneration + Environment.NewLine +
                                    String.Format(errorDialogTextGhostScriptConversion, ghostscriptEx.ErrorCode.ToString()));

            }
            finally
            {
                try
                {
                    File.Delete(standardInputFilename);
                }
                catch 
                {
                    logEventSource.TraceEvent(TraceEventType.Warning,
                                              (int)TraceEventType.Warning,
                                              String.Format(warnFileNotDeleted, standardInputFilename));
                }
                logEventSource.Flush();
            }
        }
        /// <summary>
        /// All unhandled exceptions will bubble their way up here -
        /// a final error dialog will be displayed before the crash and burn
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static void Application_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            logEventSource.TraceEvent(TraceEventType.Critical,
                                      (int)TraceEventType.Critical,
                                      ((Exception)e.ExceptionObject).Message + Environment.NewLine +
                                                                        ((Exception)e.ExceptionObject).StackTrace);
            DisplayErrorMessage(errorDialogCaption,
                                errorDialogInstructionUnexpectedError);
        }

        static bool GetPdfOutputFilename(ref String outputFile)
        {
            bool filenameRetrieved = false;

            switch (Properties.Settings.Default.AskUserForOutputFilename)
            {
                case (true) :
                    using (SetOutputFilename dialogOwner = new SetOutputFilename())
                    {
                        dialogOwner.TopMost = true;
                        dialogOwner.TopLevel = true;
                        dialogOwner.Show(); // Form won't actually show - Application.Run() never called
                                            // but having a topmost/toplevel owner lets us bring the SaveFileDialog to the front
                        dialogOwner.BringToFront();
                        using (SaveFileDialog pdfFilenameDialog = new SaveFileDialog())
                        {
                            pdfFilenameDialog.AddExtension = true;
                            pdfFilenameDialog.AutoUpgradeEnabled = true;
                            pdfFilenameDialog.CheckPathExists = true;
                            pdfFilenameDialog.Filter = "pdf files (*.pdf)|*.pdf";
                            pdfFilenameDialog.ShowHelp = false;
                            pdfFilenameDialog.Title = "PDF Scribe - Set output filename";
                            pdfFilenameDialog.ValidateNames = true;
                            if (pdfFilenameDialog.ShowDialog(dialogOwner) == DialogResult.OK)
                            {
                                outputFile = pdfFilenameDialog.FileName;
                                filenameRetrieved = true;
                            }
                        }
                        dialogOwner.Close();
                    }
                    break;
                default:
                    outputFile = GetOutputFilename();
                    filenameRetrieved = true;
                    break;
            }
            return filenameRetrieved;

        }

        private static String GetOutputFilename()
        {

            String outputFilename = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), defaultOutputFilename);
            if (!String.IsNullOrEmpty(Properties.Settings.Default.OutputFile) &&
                !String.IsNullOrWhiteSpace(Properties.Settings.Default.OutputFile))
            {
                if (IsFilePathValid(Properties.Settings.Default.OutputFile))
                {
                    outputFilename = Properties.Settings.Default.OutputFile;
                }
                else
                {
                    if (IsFilePathValid(Environment.ExpandEnvironmentVariables(Properties.Settings.Default.OutputFile)))
                    {
                        outputFilename = Environment.ExpandEnvironmentVariables(Properties.Settings.Default.OutputFile);
                    }
                }
            }
            else
            {
                logEventSource.TraceEvent(TraceEventType.Warning,
                                          (int)TraceEventType.Warning,
                                          String.Format("Using default output filename {0}",
                                                        outputFilename));
            }
            return outputFilename;
        }

        static bool IsFilePathValid(String filePath)
        {
            bool pathIsValid = false;

            if (!String.IsNullOrEmpty(filePath) && filePath.Length <= 260)
            {
                String directoryName = Path.GetDirectoryName(filePath);
                String filename = Path.GetFileName(filePath);

                if (Directory.Exists(directoryName))
                {
                    // Check for invalid filename chars
                    Regex containsABadCharacter = new Regex("["
                                                    + Regex.Escape(new String(System.IO.Path.GetInvalidPathChars())) + "]");
                    pathIsValid = !containsABadCharacter.IsMatch(filename);
                }
            }
            else
            {
                logEventSource.TraceEvent(TraceEventType.Warning,
                                          (int)TraceEventType.Warning,
                                          "Output filename is longer than 260 characters, or blank.");
            }
            return pathIsValid;
        }

        /// <summary>
        /// Displays up a topmost, OK-only message box for the error message
        /// </summary>
        /// <param name="boxCaption">The box's caption</param>
        /// <param name="boxMessage">The box's message</param>
        static void DisplayErrorMessage(String boxCaption,
                                        String boxMessage)
        {

            MessageBox.Show(boxMessage,
                            boxCaption,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error,
                            MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.DefaultDesktopOnly);
            
        }
    }
}
