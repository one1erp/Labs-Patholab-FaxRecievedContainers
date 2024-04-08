using System.Windows.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using Microsoft.Office.Interop.Outlook;
using Patholab_Common;
using Patholab_DAL_V1;
//using FAXCOMEXLib;
using FAXCOMLib;

using Exception = System.Exception;

//using FAXCOMLib;

namespace FaxContainersApp
{

    public partial class FaxRecievedContainersapp: Form
    {

        private System.Diagnostics.EventLog eventLog1;
        private   DataLayer dal;
        //private FAXCOMEXLib.FaxServer faxServer;
        private FAXCOMLib.FaxServer faxServer = new FaxServer();
        //private FAXCOMEXLib.FaxSender faxSender;
        //private FAXCOMLib.FaxSender faxSender;
        private string _scanDirectory;
        private string _outDirectory;
        private string _errorDirectory;
        //private string _coverPageLocation;
        private double _scanInterval;
        private string _faxServerName;
        private System.Timers.Timer timer;
        private   bool running = false;


        public FaxRecievedContainersapp()
        {
            InitializeComponent();
            string eventSourceName = "MySource";
            string logName = "MyNewLog";
           try
           {
               eventSourceName = ConfigurationSettings.AppSettings["eventSourceName"];
               logName = ConfigurationSettings.AppSettings["logName"];
               //_coverPageLocation = ConfigurationSettings.AppSettings["coverPageLocation"];
               _scanInterval = 0;
               double.TryParse(ConfigurationSettings.AppSettings["scanInterval"] + "000",out _scanInterval);
               _faxServerName = ConfigurationSettings.AppSettings["faxServerName"];

           }
           catch(Exception e)
           {
               log(e);
           }
            eventLog1 = new System.Diagnostics.EventLog();
            if (!System.Diagnostics.EventLog.SourceExists(eventSourceName))
            {
                System.Diagnostics.EventLog.CreateEventSource(eventSourceName, logName);
            }
            eventLog1.Source = eventSourceName;
            eventLog1.Log = logName;
        }

         private void FaxRecievedContainersapp_Load(object sender, EventArgs e)
        {
             try
             {

            
            //get data from "Shipment Scan Directory" in "System Parameters" phrase
            //Debugger.Launch();
            dal = new DataLayer();
            string connectionStrings = ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;
            dal.MockConnect(connectionStrings);
      
            if (!GetScanDirectoyFromPhrase()) return;
            //FAXCOMLib.FaxServer faxServer111 = new FAXCOMLib.FaxServer();
          //  faxServer111.Connect("192.168.0.242");
            //string machineName = Environment.MachineName;
            string machineName = _faxServerName;// "VM-RR";
            if (machineName == "THIS-PC") machineName = Environment.MachineName;
            logTextBox.Text = DateTime.Now.ToString("dd/MM/yy HH:mm:ss ") + "Connecting To '" + machineName + "' Fax Server";
            faxServer.Connect(machineName);
            log("Connected");

            Directory.CreateDirectory(_scanDirectory.TrimEnd('\\'));
            Directory.CreateDirectory(_outDirectory.TrimEnd('\\'));
            Directory.CreateDirectory(_errorDirectory.TrimEnd('\\'));

            
             GetDocumentsAndSendFax(_scanDirectory, _outDirectory);
           
            //-=================================-
            
            //FileSystemWatcher watcher = new FileSystemWatcher();
            //watcher.Path = _scanDirectory;
            //watcher.Filter = "file";
            //watcher.Created += new FileSystemEventHandler(watcher_Created);
            //watcher.EnableRaisingEvents = true;


             this.timer = new System.Timers.Timer(_scanInterval);  // 30000 milliseconds = 30 seconds

                this.timer.AutoReset = true;
                this.timer.Elapsed += new System.Timers.ElapsedEventHandler(this.timer_Elapsed);
                this.timer.Start(); 
             }
             catch (Exception exception)
             {
                 
                 log(exception);
             }
        }

        private void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            GetDocumentsAndSendFax(_scanDirectory, _outDirectory);
        }


          void watcher_Created(object sender, FileSystemEventArgs e)
        {
            GetDocumentsAndSendFax(_scanDirectory, _outDirectory);

            //Will this instance of p stick around until the timer within it is finished?
        }
          private void log(string text)
        {
            if (InvokeRequired)
            {
                this.Invoke(new MethodInvoker(delegate
                {
                    logTextBox.Text += "\r\n" + DateTime.Now.ToString("dd/MM/yy HH:mm:ss ") + text;
                }));
            }
            else
            {
                logTextBox.Text += "\r\n" + DateTime.Now.ToString("dd/MM/yy HH:mm:ss ") + text;
            }
            
        }
          private void log(Exception exception)
        {
            Logger.WriteLogFile(exception);
            if (InvokeRequired)
            {
                this.Invoke(new MethodInvoker(delegate
                {
                    logTextBox.Text += "\r\n" + DateTime.Now.ToString("dd/MM/yy HH:mm:ss ") + exception.ToString();
                }));

            }
            else
            {
                logTextBox.Text += "\r\n" + DateTime.Now.ToString("dd/MM/yy HH:mm:ss ") + exception.ToString();
            }
            
        }
        private   void GetDocumentsAndSendFax(string directoryPath, string outputPath)
        {
            try
            {
                log("Scanning Directory");
                if (running)
                {
                    log("Scan is in progress. Waiting for it to end");
                    return;
                }
                running = true;
                // create the output dir
                DateTime now = DateTime.Now;
                string yearMonth = now.Year.ToString() + @"\" + now.Month.ToString("00");
                // string newOutputPath = outputPath + yearMonth + @"\";

                outputPath += yearMonth + @"\";

                //get all Tiff files in dir, and process them
                string[] fileEntries = Directory.GetFiles(directoryPath);
                if (fileEntries != null)
                {
                    Directory.CreateDirectory(outputPath.TrimEnd('\\'));
                     Directory.CreateDirectory(_scanDirectory.TrimEnd('\\') + "\\Missing Fax Number");
                    //  eventLog1.WriteEntry("Monitoring the System", EventLogEntryType.Information, eventId++);


                }
                foreach (string fileName in fileEntries)
                {
                    try
                    {
                        if (Path.GetExtension(fileName).ToUpper() == ".TIF" &&
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
                        log(ex);
                        running = false;

                    }
                }
            }
            catch
                (System.Exception ex)
            {

                MessageBox.Show(ex.ToString());
                log(ex);
            }
            finally
            {
                running = false;
            }
        }

        public void ProcessScanFile(string filePath, string outputDirPath)
        {
            //container name is the file name
            string containerName = Path.GetFileNameWithoutExtension(filePath);
            U_CONTAINER_USER receivedContainer =
                dal.FindBy<U_CONTAINER_USER>(
                    cu =>
                    cu.U_CONTAINER.NAME == containerName.Replace("_","/")  && cu.U_STATUS != "X" &&
                    cu.U_CLINIC != null).SingleOrDefault();//&& cu.U_FAX_SEND_ON == null
            if (receivedContainer != null)
            {
                

                log("Processing file: '"+filePath.ToString()+"'");
                if (receivedContainer.U_CLINIC1.U_CLINIC_USER.U_FAX_NBR != null)
                {

                    if (Fax(filePath, receivedContainer.U_CLINIC1.U_CLINIC_USER.U_FAX_NBR,
                            receivedContainer.U_CLINIC1.NAME, receivedContainer))
                    {
                        var newFilePath = MoveToNewLocation(filePath, outputDirPath);
                    }
                    else
                    {
                        var newFilePath = MoveToNewLocation(filePath, _errorDirectory);
                    }
                }
                else
                {
                    var newFilePath = MoveToNewLocation(filePath, _scanDirectory + "\\Missing Fax Number\\");
                }
                //רוצים? שישלמו
                //if (receivedContainer.U_CLINIC1.U_CLINIC_USER.U_EMAIL_ADDRESS != null && receivedContainer.U_CLINIC1.U_CLINIC_USER.U_EMAIL_ADDRESS.Contains('@'))
                //{

                //    Email(newFilePath, receivedContainer.U_CLINIC1.U_CLINIC_USER.U_EMAIL_ADDRESS, receivedContainer.U_CLINIC1.NAME, receivedContainer);
                //}
               
            }
        }

     
        private   void Email(string scanFile, string emailAddress, string recipientName,U_CONTAINER_USER container)
        {
            var oApp = new Microsoft.Office.Interop.Outlook.Application();
            MailItem oMailItem = (MailItem)oApp.CreateItem(OlItemType.olMailItem);
            oMailItem.To = emailAddress;
            oMailItem.Attachments.Add(scanFile, OlAttachmentType.olByValue, 1, scanFile);
            oMailItem.Subject = "צידנית #" + Path.GetFileNameWithoutExtension(scanFile) + " התקבלה בפתולאב";
            ((ItemEvents_10_Event)oMailItem).Send += (MailService_Send);
            ((ItemEvents_10_Event)oMailItem).Close += (ThisAddIn_Close);
            oMailItem.Display(true);
            container.U_SEND_ON = dal.GetSysdate();
            dal.SaveChanges();
        }
          void MailService_Send(ref bool Cancel)
        {
            // SentFromOutlook = true;
        }



          void ThisAddIn_Close(ref bool Cancel)
        {

        }
        private bool Fax(string scanFile, string faxnumber, string recipientName,U_CONTAINER_USER container)
        {

            //FAXCOMEXLib.FaxDocument faxDoc = new FaxDocument();
            //FAXCOMLib.FaxDoc faxDoc = new FaxDocClass();
            //faxServer = new FAXCOMEXLib.FaxServer();

            //:todo : try to use a global fax server
            try
            {
            //FaxServer faxServer = new FaxServer();
            FaxDoc faxDoc = null;
            
                faxDoc = faxServer.CreateDocument(scanFile);
                
          
                faxDoc.FaxNumber = faxnumber;
                faxDoc.RecipientName = recipientName;
                faxDoc.DisplayName = recipientName;
                 
                //faxDoc.CoverpageName ="COVER";
                //faxDoc.SendCoverpage = 1;
                ////faxDoc.ServerCoverpage = 2;
                //faxDoc.CoverpageSubject = "Container:" + Path.GetFileNameWithoutExtension(scanFile).Replace("_", "/");

                //faxDoc.CoverpageNote = (container.U_RECEIVED_ON ?? DateTime.Now).ToString(@"dd\/MM\/yyyy");
                faxDoc.Send();
                log( "Container:'" + container.U_CONTAINER.NAME + "' was sent to fax server.");
                container.U_FAX_SEND_ON = dal.GetSysdate();
                dal.SaveChanges();
                return true;


                //var jobId = faxDoc.ConnectedSubmit(faxServer);
            }
            catch (System.Exception e)
            {

                log(e);
                log( "Error faxing Container:'" + container.U_CONTAINER.NAME + "' !");
                return false;

            }

        }
        private  string MoveToNewLocation(string filePath, string outputDirPath)
        {
            string newFilePath = outputDirPath + Path.GetFileName(filePath);
            try
            {
                int i = 1;
                string originalFileNameWithoutExtention = Path.GetFileNameWithoutExtension(filePath);
                while (File.Exists(newFilePath))
                {
                    newFilePath = outputDirPath + originalFileNameWithoutExtention + "(" + i.ToString() + ")" +
                                  Path.GetExtension(filePath);
                    i++;
                }
                File.Move(filePath, newFilePath);
            }
            catch (Exception ex)
            {
                log(ex);
            }
            return newFilePath;
        }

        private bool GetScanDirectoyFromPhrase()
        {
            bool result = false;
            try
            {
                PHRASE_HEADER systemParams = dal.GetPhraseByName("System Parameters");
                result =
                    systemParams.PhraseEntriesDictonary.TryGetValue("Shipment Scan Directory", out _scanDirectory)
                    &&
                    systemParams.PhraseEntriesDictonary.TryGetValue("Shipment Out Directory", out _outDirectory)
                    &&
                    systemParams.PhraseEntriesDictonary.TryGetValue("Shipment Error Directory", out _errorDirectory);
                     ;
                _scanDirectory += _scanDirectory.EndsWith(@"\") ? "" : @"\";
                _outDirectory += _outDirectory.EndsWith(@"\") ? "" : @"\";
                _errorDirectory += _errorDirectory.EndsWith(@"\") ? "" : @"\";

            }
            catch (System.Exception ex)
            {

                
                log(ex);
                // if (debug) MessageBox.Show(@"Error: Could not find entry  ""Shipment Scan Directory"" in Phrase ""System Parameters""");
                log(new System.Exception(@"Error: Could not find entry  ""Shipment Scan Directory"" or ""Shipment Out Directory""  or ""Shipment Error Directory""  in Phrase ""System Parameters"""));
                return false;
            }

            return result;
        }

        private void logTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void FaxRecievedContainersapp_FormClosed(object sender, FormClosedEventArgs e)
        {
            faxServer.Disconnect();
            faxServer = null;
        }
    }
}
