using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using CPUFramework;
using ContactManagerBizObjects;
using PagedList;

namespace ContactManagerCsharp2_7_2021.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index(string currentFilter,string searchString,int? page)
        {
            if (searchString != null)
            {
                page = 1;
            }
            else
            {
                searchString = currentFilter;
            }

            ViewBag.CurrentFilter = searchString;
            BizContact contactobj = new BizContact();
            List<BizContact> lst = contactobj.ListofContacts();
            if (!String.IsNullOrEmpty(searchString))
            {
                lst = contactobj.Search(searchString);
            }
            int pageSize = 10;
            int pageNumber = (page ?? 1);
            return View(lst.ToPagedList(pageNumber,pageSize));
        }

        public ActionResult Edit(int id)
        {
            BizContact contactobj = new BizContact();
            contactobj.Load(id);
            return View(contactobj);
        }

        [HttpPost]
        public ActionResult Edit(BizContact contactobj)
        {
            try
            {
                contactobj.Save();
                return RedirectToAction("Index");
            }
            catch(CPUException ex)
            {
                ViewBag.ErrorMessage = ex.Message;
                return View(contactobj);
            }
        }
        public ActionResult Delete(int id)
        {
            
            BizContact contactobj = new BizContact();
            contactobj.Load(id);
            return View(contactobj);
        }

        [HttpPost]
        public ActionResult Delete(BizContact contactobj)
        {
            long contactid = contactobj.PrimaryKeyValue;
            try
            {
                contactobj.Load(contactid);
                contactobj.Delete();
                return RedirectToAction("Index");
            }
            catch (CPUException ex)
            {
                contactobj.Load(contactid);
                ViewBag.ErrorMessage = ex.Message;
                return View(contactobj);
            }


        }
    }
}