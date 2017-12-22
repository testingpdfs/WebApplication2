using GemBox.Document;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WebApplication2.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult PrintCV(Guid? id, string docuform)
        {
          var  documentform = "pdf";
          object cv = "cv informaton";
            var document = GetProfileDocument(cv, documentform);
            if (documentform == "doc")
            {
                document.Save(this.Response, "cv name" + ".docx");
            }
            else
            {
               
                document.Save(this.Response, "cv name" + ".pdf");
            }
            return View();
        }

    private DocumentModel GetProfileDocument(object cv, string docuform)
    {
            var document = CVDocument.LoadDocumentWithHeader(cv, Server.MapPath("/helpers/CV_Template.docx"), docuform);
            var items = new List<Inline>();
            var items2 = new List<Inline>();
            //Check for null in each to make sure they don't crash if empty.  Check for 0 in each to skip if not null, but not filled with information.
            CVDocument.AddProfile(cv, items, docuform, document);
            document.Sections[0].Blocks.Add(new Paragraph(document,
            items.ToArray()
        ));
            ////Competence Section
            CVDocument.AddCompetence(cv, items2, docuform, document, "db");

            return document;
        }
    }
}