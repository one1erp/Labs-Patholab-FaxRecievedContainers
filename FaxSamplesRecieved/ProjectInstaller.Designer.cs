using System.Collections;

namespace FaxRecievedContainers
{
    partial class ProjectInstaller
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.FaxRecievedContainersServiceProcessInstaller = new System.ServiceProcess.ServiceProcessInstaller();
            this.FaxRecievedContainersServiceInstaller = new System.ServiceProcess.ServiceInstaller();
            // 
            // FaxRecievedContainersServiceProcessInstaller
            // 
            this.FaxRecievedContainersServiceProcessInstaller.Account = System.ServiceProcess.ServiceAccount.LocalSystem;
            this.FaxRecievedContainersServiceProcessInstaller.Password = null;
            this.FaxRecievedContainersServiceProcessInstaller.Username = null;
            // 
            // FaxRecievedContainersServiceInstaller
            // 
            this.FaxRecievedContainersServiceInstaller.Description = "Service that Faxes  Recieved Containers Slips from a folder defined in nautilus p" +
    "arameters";
            this.FaxRecievedContainersServiceInstaller.DisplayName = "FaxRecievedContainersService";
            this.FaxRecievedContainersServiceInstaller.ServiceName = "FaxRecievedContainersService";
            this.FaxRecievedContainersServiceInstaller.StartType = System.ServiceProcess.ServiceStartMode.Automatic;
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.FaxRecievedContainersServiceProcessInstaller,
            this.FaxRecievedContainersServiceInstaller});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller FaxRecievedContainersServiceProcessInstaller;
        private System.ServiceProcess.ServiceInstaller FaxRecievedContainersServiceInstaller;

        protected override void OnBeforeInstall(IDictionary savedState)
        {
            string parameter = "MySource1\" \"MyLogFile1";
            Context.Parameters["assemblypath"] = "\"" + Context.Parameters["assemblypath"] + "\" \"" + parameter + "\"";
            base.OnBeforeInstall(savedState);
        }
    }
}