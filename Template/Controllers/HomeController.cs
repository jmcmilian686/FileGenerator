using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using FileGenerator.Domain.Abstract;
using FileGenerator.Domain.Entities;

namespace FileGenerator.Controllers
{
    
    public class HomeController : Controller
    {
        
        private IFieldsRepository fieldsRepo;
        private IDataFieldRepository datafieldRepo;
        private ILFileRepository docRepo;
        private IStructRepository structRepo;
        private IStructFieldRepository structFieldRepo;



        public HomeController(IFieldsRepository fieldRepository, IDataFieldRepository datafieldRepository, ILFileRepository docRepository, IStructRepository structRepository, IStructFieldRepository structFieldRepository)
        {
            this.fieldsRepo = fieldRepository;
            this.datafieldRepo = datafieldRepository;
            this.docRepo = docRepository;
            this.structRepo = structRepository;
            this.structFieldRepo = structFieldRepository;

        }




        // GET: Home
        public ActionResult Index()
        {

            ViewBag.Fields = fieldsRepo.Fields.Count();
            ViewBag.Data = datafieldRepo.DataFields.Count();
            ViewBag.Document = docRepo.LFiles.Count();
            ViewBag.Struct = structRepo.Structs.Count();

            return View();
        }

       

        // GET: Home/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: Home/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Home/Create
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

        // GET: Home/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: Home/Edit/5
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

        // GET: Home/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: Home/Delete/5
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
