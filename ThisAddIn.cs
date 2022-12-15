using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Reflection;
using Microsoft.Office.Interop.Outlook;
using System.IO;

namespace ComplementoDesigner
{
    public partial class ThisAddIn
    {
        Inspectors inspectors;
        private void CrearTxt(string cuerpo, string fechaEnvio, string destinatario, string asunto)
        {
            string path = @"C:\Users\Usuario\Documents\";

            using (StreamWriter sw = new StreamWriter(Path.Combine(path, "Mails.txt"), true))
            {
                sw.WriteLine($"Fecha envio: {fechaEnvio} - Destinatario: {destinatario} - Asunto: {asunto} - Cuerpo: {cuerpo} ");
            }

        }
        void PersistenciaTxt(Inspector inspector)
        {
            Outlook.Application app;
            MailItem items;
            NameSpace ns;
            MAPIFolder inbox;

            Outlook.Application application = new Outlook.Application();
            app = application;
            ns = application.Session;
            inbox = ns.GetDefaultFolder(OlDefaultFolders.olFolderSentMail);
            items = inbox.Items.GetLast();

            string cuerpo = (items as Outlook.MailItem).Body;
            string fechaEnvio = DateTime.Now.ToString();
            string remitente = (items as Outlook.MailItem).Sender.Address;
            string destinatario = (items as Outlook.MailItem).To;
            string asunto = (items as Outlook.MailItem).Subject;

            CrearTxt(cuerpo, fechaEnvio, destinatario, asunto);
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector += new InspectorsEvents_NewInspectorEventHandler(PersistenciaTxt);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Nota: Outlook ya no genera este evento. Si tiene código que 
            //    se debe ejecutar cuando Outlook se apaga, consulte https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region Código generado por VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
