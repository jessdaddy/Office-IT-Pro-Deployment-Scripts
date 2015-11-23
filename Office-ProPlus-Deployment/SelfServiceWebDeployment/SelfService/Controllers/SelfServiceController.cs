using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.OpenIdConnect;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SelfService.Models;
using SelfService.Utils;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IdentityModel.Tokens;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace SelfService.Controllers
{
    //[Authorize]
    public class SelfServiceController : Controller
    {

        private const string TenantIdClaimType = "http://schemas.microsoft.com/identity/claims/tenantid";
        private static readonly string clientId = ConfigurationManager.AppSettings["ida:ClientId"];
        private static readonly string appKey = ConfigurationManager.AppSettings["ida:AppKey"];
        private readonly string graphResourceId = ConfigurationManager.AppSettings["ida:GraphUrl"];

        private readonly string graphUserUrl = "https://graph.windows.net/{0}/me?api-version=" +
                                               ConfigurationManager.AppSettings["ida:GraphApiVersion"];

        //
        // GET: /SelfService/

        public async Task<ActionResult> Index()
        {
            if (!Request.IsAuthenticated) {
                HttpContext.GetOwinContext()
                    .Authentication.Challenge(new AuthenticationProperties { RedirectUri = "/SelfService" },
                        OpenIdConnectAuthenticationDefaults.AuthenticationType);
                
            }

            return View();
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
    }
}
