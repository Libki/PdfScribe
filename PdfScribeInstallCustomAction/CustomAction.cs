using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;
using PdfScribeCore;
//using PdfScribe.Properties;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Configuration;
namespace PdfScribeInstallCustomAction
{
    /// <summary>
    /// Lotsa notes from here:
    /// http://stackoverflow.com/questions/835624/how-do-i-pass-msiexec-properties-to-a-wix-c-sharp-custom-action
    /// </summary>
    public class CustomActions
    {
        [CustomAction]
        public static ActionResult CheckIfPrinterNotInstalled(Session session)
        {
            ActionResult resultCode;
            SessionLogWriterTraceListener installTraceListener = new SessionLogWriterTraceListener(session);
            PdfScribeInstaller installer = new PdfScribeInstaller();

            installer.AddTraceListener(installTraceListener);
            try
            {
                if (installer.IsPdfScribePrinterInstalled())
                {
                    resultCode = ActionResult.Success;
                }
                else
                {
                    resultCode = ActionResult.Failure;
                }
            }
            finally
            {
                if (installTraceListener != null)
                    installTraceListener.Dispose();
            }

            return resultCode;
        }
        [CustomAction]
        public static ActionResult InstallPdfScribePrinter(Session session)
        {
            ActionResult printerInstalled;
            String driverSourceDirectory = session.CustomActionData["DriverSourceDirectory"];
            String outputCommand = session.CustomActionData["OutputCommand"];
            String outputCommandArguments = session.CustomActionData["OutputCommandArguments"];
            String CurrentDir = session.CustomActionData["CurrentDir"];
            String transform = session.CustomActionData["TRANSFORMS"];
            SessionLogWriterTraceListener installTraceListener = new SessionLogWriterTraceListener(session);
            installTraceListener.TraceOutputOptions = TraceOptions.DateTime;
            
            PdfScribeInstaller installer = new PdfScribeInstaller();
            
            installer.AddTraceListener(installTraceListener);


            try
            {
                System.Xml.Linq.XDocument doc = new System.Xml.Linq.XDocument();
                doc = System.Xml.Linq.XDocument.Load(CurrentDir + "\\" + "config.xml");

                session.Log("Parameters Loaded:" + (doc.Root != null));
                session.Log("Parameter Count:" + doc.Descendants("Parameter").Count());
                var parameters = doc.Descendants("Parameter").ToDictionary(n => n.Attribute("Name").Value, v => v.Attribute("Value").Value);

                if (parameters.Any())
                {
                    session.Log("Parameters loaded into Dictionary Count: " + parameters.Count());

                    //Set the Wix Properties in the Session object from the XML file
                    foreach (var parameter in parameters)
                    {
                        if (parameter.Key == "PrinterName" || parameter.Key == "PortName" || parameter.Key == "HardwareId")
                        { 
                            session.CustomActionData[parameter.Key] = parameter.Value;
                        }
                    }
                }
                else
                {
                    session.Log("No Parameters loaded");
                }


                string folder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string specificFolder = Path.Combine(folder, "PdfScribe");

                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if (transform == ":instance2")
                {
                    string printer2 = System.IO.Path.Combine(specificFolder, "scribe2");
                    System.IO.Directory.CreateDirectory(printer2);
                    doc.Save(@printer2 + "\\" + "Config.xml");
                }

                else if (transform == ":instance3")
                {
                    string printer3 = System.IO.Path.Combine(specificFolder, "scribe3");
                    System.IO.Directory.CreateDirectory(printer3);
                    doc.Save(@printer3 + "\\" + "Config.xml");
                }

                else if (transform == ":instance4")
                {
                    string printer4 = System.IO.Path.Combine(specificFolder, "scribe4");
                    System.IO.Directory.CreateDirectory(printer4);
                    doc.Save(@printer4 + "\\" + "Config.xml");
                }

                else if (transform == ":instance5")
                {
                    string printer5 = System.IO.Path.Combine(specificFolder, "scribe5");
                    System.IO.Directory.CreateDirectory(printer5);
                    doc.Save(@printer5 + "\\" + "Config.xml");
                }

                else if (transform == ":instance6")
                {
                    string printer6 = System.IO.Path.Combine(specificFolder, "scribe6");
                    System.IO.Directory.CreateDirectory(printer6);
                    doc.Save(@printer6 + "\\" + "Config.xml");
                }

                else if (transform == ":instance7")
                {
                    string printer7 = System.IO.Path.Combine(specificFolder, "scribe7");
                    System.IO.Directory.CreateDirectory(printer7);
                    doc.Save(@printer7 + "\\" + "Config.xml");
                }

                else if (transform == ":instance8")
                {
                    string printer8 = System.IO.Path.Combine(specificFolder, "scribe8");
                    System.IO.Directory.CreateDirectory(printer8);
                    doc.Save(@printer8 + "\\" + "Config.xml");
                }

                else if (transform == ":instance9")
                {
                    string printer9 = System.IO.Path.Combine(specificFolder, "scribe9");
                    System.IO.Directory.CreateDirectory(printer9);
                    doc.Save(@printer9 + "\\" + "Config.xml");
                }

                else if (transform == ":instance10")
                {
                    string printer10 = System.IO.Path.Combine(specificFolder, "scribe10");
                    System.IO.Directory.CreateDirectory(printer10);
                    doc.Save(@printer10 + "\\" + "Config.xml");
                }

                else if (transform == "" || transform == ":instance")
                {
                    string printer1 = System.IO.Path.Combine(specificFolder, "scribe");
                    System.IO.Directory.CreateDirectory(printer1);

                    doc.Save(@printer1 + "\\" + "Config.xml");
                }
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            }
            catch (Exception ex)
            {
                session.Log("ERROR in custom action SetInstallerProperties {0}", ex.ToString());
                MessageBox.Show(ex.Message);
                return ActionResult.Failure;
            }

            //Update Printer variables for Specific Instance
            if (installer.UpdateVal(session.CustomActionData["PrinterName"],session.CustomActionData["PortName"],session.CustomActionData["HardwareId"]))
            {
                //MessageBox.Show("Success");
                //MessageBox.Show(installer.PRINTERNAME+"POrt="+installer.PORTNAME+"HD="+installer.HARDWAREID);
            }
            else
            {
                MessageBox.Show("Failed To Update Instance Setup");
            }

            try
            {


                if (installer.InstallPdfScribePrinter(driverSourceDirectory,
                                                      outputCommand,
                                                      outputCommandArguments))
                {

                    printerInstalled = ActionResult.Success;
                }
                else
                {
                    printerInstalled = ActionResult.Failure;
                }
                installTraceListener.CloseAndWriteLog();
            }
            finally
            {
                if (installTraceListener != null)
                    installTraceListener.Dispose();
                
            }
            return printerInstalled;
        }
        [CustomAction]
        public static ActionResult UninstallPdfScribePrinter(Session session)
        {
            ActionResult printerUninstalled;
            SessionLogWriterTraceListener installTraceListener = new SessionLogWriterTraceListener(session);
            installTraceListener.TraceOutputOptions = TraceOptions.DateTime;

            PdfScribeInstaller installer = new PdfScribeInstaller();
            installer.AddTraceListener(installTraceListener);
            try
            {
                if (installer.UninstallPdfScribePrinter())
                    printerUninstalled = ActionResult.Success;
                else
                    printerUninstalled = ActionResult.Failure;
                installTraceListener.CloseAndWriteLog();
            }
            finally
            {
                if (installTraceListener != null)
                    installTraceListener.Dispose();
            }
            return printerUninstalled;
        }
    }
}
