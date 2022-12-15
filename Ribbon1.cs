using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Compression;
using System.IO;
using System.Reflection;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;

namespace ComplementoDesigner
{
    public partial class Ribbon1
    {
        string strRutaZip = "";
        string nombre = "";
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            cargarCombo();
            txtTimeStamp.Text = DateTime.Now.ToString();
        }
        private void cargarCombo()
        {
            var listado = ConsumirApiAsync();

            List<Usuario> lista = listado.Result;


            foreach (Usuario usu in lista)
            {
                RibbonDropDownItem item = Factory.CreateRibbonDropDownItem();
                item.Tag = usu.name;
                item.Label = usu.name;
                dropDownProyecto.Items.Add(item);
            }
        }
        private async Task<List<Usuario>> ConsumirApiAsync()
        {
            List<Usuario> listado = new List<Usuario>();

            var client = new HttpClient();

            client.BaseAddress = new Uri("http://jsonplaceholder.typicode.com/users");

            var response = await client.GetAsync(client.BaseAddress);

            if (response.IsSuccessStatusCode)
            {
                var json = await response.Content.ReadAsStringAsync();
                var listaUsuarios = JsonConvert.DeserializeObject<List<Usuario>>(json);
                foreach (var item in listaUsuarios)
                {
                    listado.Add(item);
                }
            }

            return listado;
        }

        private void btnGenerarTemplate_Click(object sender, RibbonControlEventArgs e)
        {
            bool comprimido = ComprimirCarpeta();
            if(comprimido)
            {
                CrearMailNvo();
            }
        }
        private bool ComprimirCarpeta()
        {
            string ruta = "";
            long size = 0;

            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.ShowDialog();
                ruta = dialog.SelectedPath;
            }

            DirectoryInfo dir = new DirectoryInfo(ruta);

            size = dir.EnumerateFiles("*.*", SearchOption.AllDirectories).Sum(f => f.Length);
            //Validación de que la carpeta a comprimir no supere los 15MB 
            if (size > 15728640)
            {
                MessageBox.Show("El tamaño de la carpeta supera los 15MB", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            else
            {
                strRutaZip = ruta + ".zip";

                ZipFile.CreateFromDirectory(ruta, strRutaZip);

                nombre = Path.GetFileName(strRutaZip);

                return true;
            }
        }
        private void CrearMailNvo()
        {
            var ol = new Outlook.Application();

            MailItem mail = ol.CreateItem(OlItemType.olMailItem) as MailItem;
            //Mensaje predefinido por codigo
            mail.Body = "Buenas tardes Cliente";
            //Tomar el item que se selecciono desde el combo
            string nomPro = dropDownProyecto.SelectedItem.ToString();
            //Tomar la hora actural
            string timestamp = DateTime.Now.ToString();
            //Añadiendo etiqueta con nombre de proyecto y hora actual
            mail.Body += "TAG - " + nomPro + " _ " +  timestamp + " - TAG";
            //Adjuntando carpeta comprimida
            //(strRutaZip es un campo de la clase Ribbon.cs y se le asigna su valor en el momento de comprimir la carpeta seleccionada)
            mail.Attachments.Add(strRutaZip, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
            //Se abre la ventana de mail nuevo.
            mail.Display();

        }
        private void btnObtenerDatosMail_Click(object sender, RibbonControlEventArgs e)
        {
            GetDatosMail();
        }
        private void GetDatosMail()
        {
            MailItem mail = (MailItem)Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];

            bool resultTag = mail.Body.Contains("TAG -");

            if (!resultTag)
            {
                MessageBox.Show("Proyecto no encontrado!");
            }
            else
            {
                int index = mail.Body.IndexOf("TAG -");

                string cadenaConTag = mail.Body.Substring(index);

                string titProSinUnTag = cadenaConTag.Substring(6);
                int indexFinTitulo = titProSinUnTag.IndexOf("_");
                string tituloProyectoFinal = titProSinUnTag.Substring(0, (indexFinTitulo - 1));

                int indexIniTs = cadenaConTag.IndexOf("_");
                string tsSinunTag = cadenaConTag.Substring(indexIniTs + 1);
                int indexFinTs = tsSinunTag.IndexOf("-");
                string tsFinal = tsSinunTag.Substring(0, (indexFinTs - 1));

                txtTimeStamp.Text = tsFinal;

                string remitente = mail.Sender.Address;

                txtRemitente.Text = remitente;

                var listado = ConsumirApiAsync();

                List<Usuario> lista = listado.Result;
                int indexPro = 0;
                foreach (var item in lista)
                {
                    if (item.name == tituloProyectoFinal)
                    {
                        indexPro = lista.IndexOf(item);
                    }
                }

                dropDownProyecto.SelectedItemIndex = indexPro;
            }
        }
        private void btnRegistrarMailRecibido_Click(object sender, RibbonControlEventArgs e)
        {
            RegitrarMailRecibidoTxt();
        }
        private void RegitrarMailRecibidoTxt()
        {
            MailItem mail = (MailItem)Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];

            bool resultTag = mail.Body.Contains("TAG -");

            if (!resultTag)
            {
                MessageBox.Show("Datos de Proyecto y TimeStamp no se encuentran");
            }
            else
            {
                string path = @"C:\Users\Usuario\Documents\";

                int index = mail.Body.IndexOf("TAG -");
                string cadenaConTag = mail.Body.Substring(index);
                string titProSinUnTag = cadenaConTag.Substring(6);
                int indexFinTitulo = titProSinUnTag.IndexOf("_");
                string tituloProyectoFinal = titProSinUnTag.Substring(0, (indexFinTitulo - 1));

                string remitente = mail.Sender.Address;

                int indexIniTs = cadenaConTag.IndexOf("_");
                string tsSinunTag = cadenaConTag.Substring(indexIniTs + 1);
                int indexFinTs = tsSinunTag.IndexOf("-");
                string tsFinal = tsSinunTag.Substring(0, (indexFinTs - 1));

                string timeStampActual = DateTime.Now.ToString();

                using (StreamWriter sw = new StreamWriter(Path.Combine(path, "Mails.txt"), true))
                {
                    sw.WriteLine($"Proyecto: {tituloProyectoFinal} - Remitente: {remitente} - " +
                        $"TimeStamp de creacion: {tsFinal} - TimeStampActual: {timeStampActual}");
                }
                //FALTA MENSAJE DE CONFIRMACION
            }
        }
        private void btnGuardarMailRecibido_Click(object sender, RibbonControlEventArgs e)
        {
            GuardarMailRecibido();
        }
        private void GuardarMailRecibido()
        {
            MailItem mail = (MailItem)Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
            int index = mail.Body.IndexOf("TAG -");
            string cadenaConTag = mail.Body.Substring(index);
            string titProSinUnTag = cadenaConTag.Substring(6);
            int indexFinTitulo = titProSinUnTag.IndexOf("_");
            string tituloProyectoFinal = titProSinUnTag.Substring(0, (indexFinTitulo - 1));

            int indexIniTs = cadenaConTag.IndexOf("_");
            string tsSinunTag = cadenaConTag.Substring(indexIniTs + 1);
            int indexFinTs = tsSinunTag.IndexOf("-");
            string tsFinal = tsSinunTag.Substring(0, (indexFinTs - 1));

            string tsConvertido = tsFinal.Replace("/", "-");

            string fi = tsConvertido.Replace(":", ".");

            string ubicacion = $@"C:\Users\Usuario\Documents\{tituloProyectoFinal}{fi}.msg";

            mail.SaveAs(ubicacion);
            MessageBox.Show("Mail guardado con exito!");
        }
    }
}
