using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
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
        public ActionResult UploadFile(HttpPostedFileBase file, string documentLibraryName, string folderName, string tipo = null)
        {
            // Credenciales de SharePoint
            string userName = ConfigurationManager.AppSettings["SharePointUserName"];
            string password = ConfigurationManager.AppSettings["SharePointPassword"];
            string siteUrl = ConfigurationManager.AppSettings["SharePointSiteUrl"];
            //string documentLibraryName = "Administracion";

            // Inicializar la URL del archivo
            string fileUrl = string.Empty;

            // Validar que se haya seleccionado un archivo
            if (file != null && file.ContentLength > 0)
            {
                // Llamar al servicio para subir el archivo a SharePoint y obtener la URL
                using (var fileStream = file.InputStream)
                {
                    fileUrl = UploadFileAndGetUrl(file.FileName, fileStream, documentLibraryName, folderName, userName, password, siteUrl, tipo);
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
        private string UploadFileAndGetUrl(string fileName, Stream fileStream, string libraryName, string folderName, string userName, string password, string siteUrl, string tipo)
        {
            var credentials = new NetworkCredential(userName, password);
            using (var context = new ClientContext(siteUrl))
            {
                context.Credentials = credentials;
                var list = context.Web.Lists.GetByTitle(libraryName);
                context.Load(list.RootFolder);
                context.ExecuteQuery();

                Folder folder;
                if (tipo == "informe")
                {
                    // Si el tipo es "informe", subir directamente a la biblioteca de documentos
                    folder = list.RootFolder;
                }
                else
                {
                    // Verificar si la carpeta existe
                    FolderCollection folders = list.RootFolder.Folders;
                    context.Load(folders, fl => fl.Include(ct => ct.Name)
                                                  .Where(ct => ct.Name == folderName));
                    context.ExecuteQuery();

                    if (folders.Count > 0)
                    {
                        folder = context.Web.GetFolderByServerRelativeUrl($"{list.RootFolder.ServerRelativeUrl}/{folderName}");
                        context.Load(folder);
                        context.ExecuteQuery();
                    }
                    else
                    {
                        // Si la carpeta no existe, crearla
                        folder = list.RootFolder.Folders.Add(folderName);
                        context.Load(folder);
                        context.ExecuteQuery();
                    }
                }

                var fileUrl = $"{folder.ServerRelativeUrl}/{fileName}";

                // Crear la información del archivo con el stream
                var fileCreationInfo = new FileCreationInformation
                {
                    ContentStream = fileStream,
                    Url = fileUrl,
                    Overwrite = true
                };
                // Agregar el archivo a la carpeta o biblioteca de documentos
                var uploadFile = folder.Files.Add(fileCreationInfo);
                context.ExecuteQuery();

                return fileUrl;
            }
        }



        // Método para obtener los archivos de la biblioteca en SharePoint
        public ActionResult GetFiles()
        {
            // Credenciales de SharePoint
            string userName = ConfigurationManager.AppSettings["SharePointUserName"];
            string password = ConfigurationManager.AppSettings["SharePointPassword"];
            string siteUrl = ConfigurationManager.AppSettings["SharePointSiteUrl"];
            string documentLibraryName = "Adquisiciones";

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
                    fileUrls.Add("http://10.223.163.100" + file.ServerRelativeUrl);
                }
            }

            // Pasar las URLs de los archivos a la vista
            ViewBag.FileUrls = fileUrls;

            return View();
        }

        // Método para obtener un archivo desde SharePoint y devolverlo al cliente
        public ActionResult GetFile(string fileName, string documentLibraryName, string folderName, string tipo = null)
        {
            try
            {
                // Credenciales de SharePoint
                string userName = ConfigurationManager.AppSettings["SharePointUserName"];
                string password = ConfigurationManager.AppSettings["SharePointPassword"];
                string siteUrl = ConfigurationManager.AppSettings["SharePointSiteUrl"];

                // Obtener el archivo desde SharePoint
                var credentials = new NetworkCredential(userName, password);
                using (var context = new ClientContext(siteUrl))
                {
                    context.Credentials = credentials;
                    var list = context.Web.Lists.GetByTitle(documentLibraryName);
                    context.Load(list.RootFolder);
                    context.ExecuteQuery();
                    Microsoft.SharePoint.Client.File file = null;
                    if (tipo == "informe")
                    {
                        Folder folder;
                        try
                        {
                            folder = context.Web.GetFolderByServerRelativeUrl($"{list.RootFolder.ServerRelativeUrl}/");
                            context.Load(folder);
                            context.ExecuteQuery();
                        }
                        catch (ServerException ex)
                        {
                            // Si la carpeta no existe, devolver un error al cliente
                            return Content(ex.Message);
                        }

                        var files = folder.Files;
                        context.Load(files);
                        context.ExecuteQuery();
                        // Buscar el archivo en la carpeta                                                        
                        foreach (var f in files)
                        {
                            if (f.Name == fileName)
                            {
                                file = f;
                                break;
                            }
                        }
                    }
                    else
                    {
                        // Obtener la carpeta
                        Folder folder;
                        try
                        {
                            folder = context.Web.GetFolderByServerRelativeUrl($"{list.RootFolder.ServerRelativeUrl}/{folderName}");
                            context.Load(folder);
                            context.ExecuteQuery();
                        }
                        catch (ServerException ex)
                        {
                            // Si la carpeta no existe, devolver un error al cliente
                            return Content(ex.Message);
                        }

                        var files = folder.Files;
                        context.Load(files);
                        context.ExecuteQuery();
                        // Buscar el archivo en la carpeta                                                        
                        foreach (var f in files)
                        {
                            if (f.Name == fileName)
                            {
                                file = f;
                                break;
                            }
                        }
                    }

                    // Si no se encontró el archivo, devolver un error al cliente
                    if (file == null)
                    {
                        return HttpNotFound($"El archivo '{fileName}' no se encontró.");
                    }

                    // Obtener el archivo desde SharePoint y devolverlo al cliente
                    var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(context, file.ServerRelativeUrl);
                    return File(fileInfo.Stream, "application/octet-stream", fileName);
                }
            }
            catch (Exception e)
            {
                return Content(e.Message);
            }
        }


        // Método para obtener un archivo desde SharePoint y devolverlo al cliente        
        public ActionResult GetFile2(string fileName, string documentLibraryName)
        {
            try
            {
                // Credenciales de SharePoint
                string userName = ConfigurationManager.AppSettings["SharePointUserName"];
                string password = ConfigurationManager.AppSettings["SharePointPassword"];
                string siteUrl = ConfigurationManager.AppSettings["SharePointSiteUrl"];

                // Obtener el archivo desde SharePoint
                var credentials = new NetworkCredential(userName, password);
                using (var context = new ClientContext(siteUrl))
                {
                    context.Credentials = credentials;
                    var list = context.Web.Lists.GetByTitle(documentLibraryName);
                    var files = list.RootFolder.Files;
                    context.Load(files);
                    context.ExecuteQuery();

                    // Verificar si la colección de archivos no está vacía
                    if (files != null && files.Count > 0)
                    {
                        // Iterar sobre cada archivo en la colección de archivos
                        foreach (var file in files)
                        {
                            // Verificar si el nombre del archivo coincide con el nombre especificado
                            if (file.Name == fileName)
                            {
                                // Obtener el archivo desde SharePoint y devolverlo al cliente
                                var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(context, file.ServerRelativeUrl);
                                return File(fileInfo.Stream, "application/octet-stream", fileName);
                            }
                        }

                        // Si no se encontró el archivo, devolver un error al cliente
                        return HttpNotFound();
                    }
                    else
                    {
                        // Si la colección de archivos está vacía, devolver un mensaje de error al cliente
                        return Content("No se encontraron archivos en la biblioteca.");
                    }
                }
            }
            catch (Exception e)
            {
                return Content(e.Message);
            }
        }

        public ActionResult DeleteFile(string fileName, string documentLibraryName)
        {
            try
            {
                // Credenciales de SharePoint
                string userName = ConfigurationManager.AppSettings["SharePointUserName"];
                string password = ConfigurationManager.AppSettings["SharePointPassword"];
                string siteUrl = ConfigurationManager.AppSettings["SharePointSiteUrl"];

                // Eliminar el archivo de SharePoint
                var credentials = new NetworkCredential(userName, password);
                using (var context = new ClientContext(siteUrl))
                {
                    context.Credentials = credentials;
                    var list = context.Web.Lists.GetByTitle(documentLibraryName);
                    var files = list.RootFolder.Files;
                    context.Load(files);
                    context.ExecuteQuery();
                    // Verificar si la colección de archivos no está vacía
                    if (files != null && files.Count > 0)
                    {
                        // Iterar sobre cada archivo en la colección de archivos
                        foreach (var file in files)
                        {
                            // Verificar si el nombre del archivo coincide con el nombre especificado
                            if (file.Name == fileName)
                            {
                                file.DeleteObject();
                                context.ExecuteQuery();
                                return Content($"El archivo '{fileName}' se ha eliminado correctamente.");
                            }
                        }

                        // Si no se encontró el archivo, devolver un error al cliente
                        return HttpNotFound();
                    }
                    else
                    {
                        // Si la colección de archivos está vacía, devolver un mensaje de error al cliente
                        return Content("No se encontraron archivos en la biblioteca.");
                    }
                }
            }
            catch (Exception e)
            {
                // Devolver un mensaje de error
                return Content($"Error al eliminar el archivo '{fileName}': {e.Message}");
            }
        }





        // Endpoint para verificar la conectividad
        [HttpGet]
        public ActionResult TestConnection()
        {
            // Puedes devolver un mensaje simple
            return Content("API is up and running!");

            // O puedes devolver un código de estado HTTP 200 OK
            // return new HttpStatusCodeResult(HttpStatusCode.OK);
        }
    }
}
