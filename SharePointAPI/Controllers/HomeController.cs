using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Web;
using System.Web.Mvc;
using Microsoft.SharePoint.Client;

namespace SharePointAPI.Controllers
{
    public class HomeController : Controller
    {
        // GET: Index
        public ActionResult Index()
        {
            return View();
        }

        // Método para subir un archivo a SharePoint y devolver su URL
        [HttpPost]
        public ActionResult UploadFile(HttpPostedFileBase file, string documentLibraryName)
        {
            // Credenciales de SharePoint
            string userName = "gap\\SPInstall";
            string password = "LRb1@uRg.Bft+aNzth7n";
            string siteUrl = "http://10.223.163.100/sites/zonasegura";
            //string documentLibraryName = "Administracion";

            // Inicializar la URL del archivo
            string fileUrl = string.Empty;

            // Validar que se haya seleccionado un archivo
            if (file != null && file.ContentLength > 0)
            {
                // Llamar al servicio para subir el archivo a SharePoint y obtener la URL
                using (var fileStream = file.InputStream)
                {
                    fileUrl = "http://10.223.163.100/sites/zonasegura/" + UploadFileAndGetUrl(file.FileName, fileStream, documentLibraryName, userName, password, siteUrl);
                }
            }
            else
            {
                ViewBag.Message = "Por favor seleccione un archivo.";
            }

            // Devolver la URL del archivo
            return Content(fileUrl);
        }


        // Método para subir el archivo utilizando el servicio existente y obtener la URL del archivo
        private string UploadFileAndGetUrl(string fileName, Stream fileStream, string libraryName, string userName, string password, string siteUrl)
        {
            var credentials = new NetworkCredential(userName, password);
            using (var context = new ClientContext(siteUrl))
            {
                context.Credentials = credentials;
                var list = context.Web.Lists.GetByTitle(libraryName);
                context.Load(list.RootFolder);
                context.ExecuteQuery();

                var fileUrl = $"{list.RootFolder.ServerRelativeUrl}/{fileName}";

                // Convertir el contenido del archivo a un arreglo de bytes
                byte[] fileBytes;
                using (var memoryStream = new MemoryStream())
                {
                    fileStream.CopyTo(memoryStream);
                    fileBytes = memoryStream.ToArray();
                }

                // Crear la información del archivo con el arreglo de bytes
                var fileCreationInfo = new FileCreationInformation
                {
                    Content = fileBytes,
                    Url = fileUrl,
                    Overwrite = true
                };

                // Agregar el archivo al directorio raíz de la lista
                var uploadFile = list.RootFolder.Files.Add(fileCreationInfo);
                context.ExecuteQuery();

                return fileUrl;
            }
        }

        // Método para obtener los archivos de la biblioteca en SharePoint
        public ActionResult GetFiles()
        {
            // Credenciales de SharePoint
            string userName = "gap\\SPInstall";
            string password = "LRb1@uRg.Bft+aNzth7n";
            string siteUrl = "http://10.223.163.100/sites/zonasegura";
            string documentLibraryName = "Administracion";

            // Lista para almacenar las URLs de los archivos
            List<string> fileUrls = new List<string>();

            // Consultar los archivos en la biblioteca
            var credentials = new NetworkCredential(userName, password);
            using (var context = new ClientContext(siteUrl))
            {
                context.Credentials = credentials;
                var list = context.Web.Lists.GetByTitle(documentLibraryName);
                var files = list.RootFolder.Files;
                context.Load(files);
                context.ExecuteQuery();

                // Obtener las URLs de los archivos
                foreach (var file in files)
                {
                    fileUrls.Add("http://10.223.163.100"+file.ServerRelativeUrl);
                }
            }

            // Pasar las URLs de los archivos a la vista
            ViewBag.FileUrls = fileUrls;

            return View();
        }
    }
}
