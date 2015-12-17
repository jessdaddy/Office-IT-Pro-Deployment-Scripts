using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Xml.Linq;
using System.Net;
using System.IO;
using Microsoft.Web; 

namespace SelfService.Controllers
{
    public class SelfServiceController : Controller
    {
        //
        // GET: /SelfService/

        public ActionResult Index()
        {
            return View();
        }

        //
        //Get: /SelfService/Languages
        public string Languages()
        {
            return ConfigurationManager.AppSettings["Language"];
        }

        //
        //Get: /SelfService/Products
        public string Products()
        {
            return ConfigurationManager.AppSettings["Product"];
        }

        //
        //Get: /SelfService/Versions
        public string Versions()
        {
            return ConfigurationManager.AppSettings["Version"];
        }


        //
        // GET: /SelfService/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /SelfService/Create

        public ActionResult Create()
        {
            return View();
        }

        //
        // POST: /SelfService/Create

        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
                // TODO: Add insert logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /SelfService/Edit/5

        public ActionResult Edit(int id)
        {
            return View();
        }

        //
        // POST: /SelfService/Edit/5

        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /SelfService/Delete/5

        public ActionResult Delete(int id)
        {
            return View();
        }

        //
        // POST: /SelfService/Delete/5

        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }


        private void addLanguages(XDocument doc, List<string> languageList){

            foreach (string language in languageList)
            {
                doc.Root.Element("Add").Element("Product").Add(
                    new XElement("Language",
                        new XAttribute("ID", language)));
            } 
        }

        private void createQueryString(string xmlPath, string setupPath)
        {
            var baseURL = Request.Url.GetLeftPart(UriPartial.Authority);
            var builder = new UriBuilder(baseURL); 
          

            var query = HttpUtility.ParseQueryString(string.Empty);
            query["xmlPath"] = xmlPath;
            query["setupPath"] = setupPath;

            builder.Query = query.ToString();

            Response.Redirect("index.cshtml?xml=" + xmlPath + "&installer=" + setupPath);


        }


        public ActionResult generateXML(string buildName, List<string> languageList, string uiLanguage)
        {
            string result;

            try
            {            
                string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
                var oldXML = XDocument.Load(currentDirectory + "Content\\XML_Build_Files\\Base_Files\\" + buildName + ".xml");
                XDocument newXML = new XDocument(oldXML);

                newXML.Root.Element("Add").Element("Product").Add(
                  new XElement("Language",
                      new XAttribute("ID", uiLanguage)));


                if(languageList.Count > 1)
                {
                    addLanguages(newXML, languageList);
                }

                string fileName = Guid.NewGuid().ToString() + ".xml";
                string savePath = currentDirectory + "Content\\XML_Build_Files\\Generated_Files\\" + fileName;
                newXML.Save(savePath);

                string xmlPath = Request.Url.GetLeftPart(UriPartial.Authority) + HttpRuntime.AppDomainAppVirtualPath + "Content/XML_Build_Files/Generated_Files/" + fileName ;
                string exePath = Request.Url.GetLeftPart(UriPartial.Authority) + HttpRuntime.AppDomainAppVirtualPath +  "Content/ODT_Launcher.application";
                string setupPath = Request.Url.GetLeftPart(UriPartial.Authority) + HttpRuntime.AppDomainAppVirtualPath + "Content/Office2016Setup.exe";


                return Json(new { xml = xmlPath, exe = exePath, setup = setupPath });
            }
            catch(Exception e)
            {
                Response.StatusCode = 500;
                if (e.Message.Contains("Could not find file"))
                {
                    result = "Base configuration file does not exist for " + buildName;
                }
                else
                {
                    result = e.Message;
                }

                return Json(new { message = result });
            }

        }
    }
}
