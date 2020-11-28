using FiletoDatabase.Models;
using GemBox.Document;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web;
using System.Web.Mvc;

namespace FiletoDatabase.Controllers
{
    public class FileUploadController : Controller
    {
        string desktopPathHtml, desktopPathDocx;
        public FileUploadController()
        {
            desktopPathHtml = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Output.html";
            desktopPathDocx = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Output.docx";
        }
        // GET: FileUpload
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]  
        public ActionResult Index(HttpPostedFileBase file)
        {
            dbDoctoHtmlEntities dbModel = new dbDoctoHtmlEntities();
            var path = Path.Combine(Server.MapPath("~/Files"),Path.GetFileName(file.FileName));
            file.SaveAs(path);
            WordToHtml(path,desktopPathHtml);

            var temp = System.IO.File.ReadAllText(desktopPathHtml);
            dbModel.FilePath.Add(new FilePath { Source = temp,ID=Guid.NewGuid() });
            dbModel.SaveChanges();

            HtmlToWord(path,desktopPathDocx);
            return RedirectToAction(nameof(Index));
        }

        public void HtmlToWord(string path,string output)
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            ComponentInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;
            DocumentModel document = DocumentModel.Load(path);
            Section section = document.Sections[0];
            PageSetup pageSetup = section.PageSetup;
            PageMargins pageMargins = pageSetup.PageMargins;
            pageMargins.Top = pageMargins.Bottom = pageMargins.Left = pageMargins.Right = 0;
            document.Save(output);
        }

        public void WordToHtml(string path,string output)
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            ComponentInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;
            var document = DocumentModel.Load(path);
            var saveOptions = new HtmlSaveOptions()
            {
                HtmlType = HtmlType.Html,
                EmbedImages = true,
                UseSemanticElements = true
            };
            var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            document.Save(output, saveOptions);
        }
        public ActionResult GetUploadedFile()
        {
            dbDoctoHtmlEntities dbModel = new dbDoctoHtmlEntities();
            //Editorden gelen string(html) verisi
            var temp = System.IO.File.ReadAllText(desktopPathHtml);
            return View(dbModel.FilePath.ToList());
        }
    }
}