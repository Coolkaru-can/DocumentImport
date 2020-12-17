using System;
using System.Text;
using System.Windows.Forms;

namespace DocumentImport
{
    public partial class Monitor : Form
    {
        public Monitor()
        {
            InitializeComponent();
        }
        private void Monitor_Load(object sender, EventArgs e)
        {
            this.Text = JMSL.Framework.Divers.My.AppName() + " - Monitor mode";
            JMSL.Framework.Log.Manager.LogManager.LoggerObject.LogOccured += LoggerObject_LogOccured;
        }

        void LoggerObject_LogOccured(object sender, JMSL.Framework.Log.EventArgs.LogEventArgs e)
        {
            StringBuilder s = new StringBuilder(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            s.AppendFormat(" {0} {1}", e.EventType, e.Message);


            if (listBox1.InvokeRequired)
                listBox1.Invoke((MethodInvoker)(() => { listBox1.Items.Insert(0, s); }));
            else
                listBox1.Items.Insert(0, s);

        }

        private void Monitor_FormClosed(object sender, FormClosedEventArgs e)
        {
            JMSL.Framework.Log.Manager.LogManager.LoggerObject.LogOccured -= LoggerObject_LogOccured;
        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            MessageBox.Show((string)listBox1.SelectedItem);
        }

        private void Monitor_Shown(object sender, EventArgs e)
        {
            DocumentImportService.Start();
            //ComboxFolder.Rename(@"Méthode", "Approbation client", "Approbation et technique", "Technics and approvals");
            //ComboxFolder.Rename(@"Méthode", "Document approuvé par le client", "Documents approuvés", "Release document");
            /*ComboxFolder.AddTemplateFolderToAllProjects("Méthode/Approbation et technique", "Archives", "Archives");
            ComboxFolder.AddTemplateFolderToAllProjects("Méthode/Documents approuvés", "Archives", "Archives");
            ComboxFolder.AddTemplateFolderToAllProjects("Méthode/MCR", "Archives", "Archives");*/

            //ArchitectFolderTemplate.ImportFile();
        }
    }
}
