using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using Microsoft.Office.Interop.Outlook;
using Patholab_Common;
using Patholab_DAL_V1;
using FAXCOMEXLib;
using Exception = System.Exception;

//using FAXCOMLib;

namespace FaxRecievedContainers
{

    public partial class FaxRecievedContainersService : ServiceBase
    {

        private System.Diagnostics.EventLog eventLog1;
        private static DataLayer dal;
        private static FAXCOMEXLib.FaxServer faxServer ;
        private static FAXCOMEXLib.FaxSender faxSender;
        private static string scanDirectory;
        private static string outDirectory;
        private Timer timer;
        private static bool running = false;

        public FaxRecievedContainersService(string[] args)
        {
            InitializeComponent();
            string eventSourceName = "MySource";
            string logName = "MyNewLog";
            if (args.Count() > 0)
            {
                eventSourceName = args[0];
            }
            if (args.Count() > 1)
            {
                logName = args[1];
            }
            eventLog1 = new System.Diagnostics.EventLog();
            if (!System.Diagnostics.EventLog.SourceExists(eventSourceName))
            {
                System.Diagnostics.EventLog.CreateEventSource(eventSourceName, logName);
            }
            eventLog1.Source = eventSourceName;
            eventLog1.Log = logName;
        }

        protected override void OnStart(string[] args)
        {

            //get data from "Shipment Scan Directory" in "System Parameters" phrase
            Debugger.Launch();
            dal = new DataLayer();
            string connectionStrings = ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;
            dal.MockConnect(connectionStrings);
            if (!GetScanDirectoyFromPhrase()) return;
            faxServer=new FaxServer();
            faxServer.Connect("");

            GetDocumentsAndSendFax(scanDirectory, outDirectory);


            //-=================================-
            Directory.CreateDirectory(scanDirectory.TrimEnd('\\'));
            FileSystemWatcher watcher = new FileSystemWatcher();
            watcher.Path = scanDirectory;
            watcher.Filter = "file";
            watcher.Created += new FileSystemEventHandler(watcher_Created);
            watcher.EnableRaisingEvents = true;
            this.timer = new System.Timers.Timer(30000D);  // 30000 milliseconds = 30 seconds

            this.timer.AutoReset = true;
            this.timer.Elapsed += new System.Timers.ElapsedEventHandler(this.timer_Elapsed);
            this.timer.Start();
        }

        private void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            GetDocumentsAndSendFax(scanDirectory, outDirectory);
        }


        static void watcher_Created(object sender, FileSystemEventArgs e)
        {
            GetDocumentsAndSendFax(scanDirectory, outDirectory);

            //Will this instance of p stick around until the timer within it is finished?
        }

        private static void GetDocumentsAndSendFax(string directoryPath, string outputPath)
        {
            try
            {

                if (running) return;
                running = true;
                // create the output dir
                DateTime now = DateTime.Now;
                string yearMonth = now.Year.ToString() + @"\" + now.Month.ToString("00");
               // string newOutputPath = outputPath + yearMonth + @"\";

                outputPath += yearMonth + @"\";

                //get all Tiff files in dir, and process them
                string[] fileEntries = Directory.GetFiles(directoryPath);
                if (fileEntries!=null)
                {
                    Directory.CreateDirectory(outputPath.TrimEnd('\\'));
                  //  eventLog1.WriteEntry("Monitoring the System", EventLogEntryType.Information, eventId++);


                }
                foreach (string fileName in fileEntries)
                {
                    try
                    {
                        if (Path.GetExtension(fileName).ToUpper() == ".TIFF" &&
                            Path.GetFileName(fileName).Substring(0, 1) != "$" &&
                            Path.GetFileName(fileName).Substring(0, 1) != "~")
                        {
                            // CloseWord(fileName);
                            //Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                            ProcessScanFile(fileName, outputPath);
                        }
                    }
                    catch
                        (System.Exception ex)
                    {
                        Logger.WriteLogFile(ex);
                    }
                }
            }
            catch
                (System.Exception ex)
            {
                running = false;
                Logger.WriteLogFile(ex);
            }
        }

        public static void ProcessScanFile(string filePath, string outputDirPath)
        {
            //container name is the file name
            string containerName = Path.GetFileNameWithoutExtension(filePath);
            U_CONTAINER_USER receivedContainer =
                dal.FindBy<U_CONTAINER_USER>(
                    cu =>
                    cu.U_CONTAINER.NAME == containerName  && cu.U_STATUS != "X" &&
                    cu.U_CLINIC != null).SingleOrDefault();//&& cu.U_FAX_SEND_ON == null
            if (receivedContainer != null)
            {
                string newFilePath = outputDirPath + Path.GetFileName(filePath);
                try
                {
                    File.Move(filePath, newFilePath);
                }
                catch (Exception ex)
                {
                    
                }
               

                if (receivedContainer.U_CLINIC1.U_CLINIC_USER.U_FAX_NBR != null)
                {

                    Fax(newFilePath, receivedContainer.U_CLINIC1.U_CLINIC_USER.U_FAX_NBR, receivedContainer.U_CLINIC1.NAME);
                }
                if (receivedContainer.U_CLINIC1.U_CLINIC_USER.U_EMAIL_ADDRESS != null && receivedContainer.U_CLINIC1.U_CLINIC_USER.U_EMAIL_ADDRESS.Contains('@'))
                {

                    Email(newFilePath, receivedContainer.U_CLINIC1.U_CLINIC_USER.U_EMAIL_ADDRESS, receivedContainer.U_CLINIC1.NAME);
                }
            }



        }

        private static void Email(string scanFile, string emailAddress, string recipientName)
        {
            var oApp = new Microsoft.Office.Interop.Outlook.Application();
            MailItem oMailItem = (MailItem)oApp.CreateItem(OlItemType.olMailItem);
            oMailItem.To = emailAddress;
            oMailItem.Attachments.Add(scanFile, OlAttachmentType.olByValue, 1, scanFile);
            oMailItem.Subject = "צידנית #" + Path.GetFileNameWithoutExtension(scanFile) + " התקבלה בפתולאב";
            ((ItemEvents_10_Event)oMailItem).Send += (MailService_Send);
            ((ItemEvents_10_Event)oMailItem).Close += (ThisAddIn_Close);
            oMailItem.Display(true);
        }
        static void MailService_Send(ref bool Cancel)
        {
            // SentFromOutlook = true;
        }



        static void ThisAddIn_Close(ref bool Cancel)
        {

        }
        private static void Fax(string scanFile, string faxnumber, string recipientName)
        {

            FAXCOMEXLib.FaxDocument faxDoc = new FaxDocument();
            
            
            try
            {
                faxServer.Connect(Environment.MachineName);
            }
            catch (System.Exception e)
            {
                Logger.WriteLogFile(e);
            }

            try
            {
                //faxDoc =faxServer.CreateDocument(scanFile);
                
            }
            catch (System.Exception e)
            {
                Logger.WriteLogFile(e);
            }

            try
            {
                faxDoc.Recipients.Add(faxnumber, recipientName);
                faxDoc.Subject = "Container:"+ Path.GetFileNameWithoutExtension(scanFile);
                faxDoc.DocumentName = Path.GetFileName(scanFile);
                faxDoc.AttachFaxToReceipt = true;
                faxDoc.Body = scanFile;
                faxDoc.ReceiptType = FAX_RECEIPT_TYPE_ENUM.frtMAIL;
                //Set the cover page type and the path to the cover page
                faxDoc.CoverPageType = FAXCOMEXLib.FAX_COVERPAGE_TYPE_ENUM.fcptSERVER;
                faxDoc.CoverPage = "  , פתולאב" + Path.GetFileNameWithoutExtension(scanFile);
                faxDoc.Note = "הצידנית התקבלה בחברתנו, פתולאב";

            }
            catch (System.Exception e)
            {
                Logger.WriteLogFile(e);
            }


            try
            {

                var jobId = faxDoc.ConnectedSubmit(faxServer);
            }
            catch (System.Exception e)
            {
                Logger.WriteLogFile(e); 
            }

        }

        private bool GetScanDirectoyFromPhrase()
        {
            bool result = false;
            try
            {
                PHRASE_HEADER systemParams = dal.GetPhraseByName("System Parameters");
                result =
                   systemParams.PhraseEntriesDictonary.TryGetValue("Shipment Scan Directory", out scanDirectory)
                   &&
                     systemParams.PhraseEntriesDictonary.TryGetValue("Shipment Out Directory", out outDirectory);
                scanDirectory += scanDirectory.EndsWith(@"\") ? "" : @"\";
                outDirectory += outDirectory.EndsWith(@"\") ? "" : @"\";



            }
            catch (System.Exception ex)
            {
                Logger.WriteLogFile(ex);
                // if (debug) MessageBox.Show(@"Error: Could not find entry  ""Shipment Scan Directory"" in Phrase ""System Parameters""");
                Logger.WriteLogFile(new System.Exception(@"Error: Could not find entry  ""Shipment Scan Directory"" or ""Shipment Out Directory"" in Phrase ""System Parameters"""));
                return false;
            }

            return result;
        }

        protected override void OnStop()
        {
        }
    }
}
