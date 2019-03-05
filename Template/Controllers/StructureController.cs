using FileGenerator.Domain.Abstract;
using FileGenerator.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Template.Controllers
{
    public class StructureController : Controller
    {
        private IStructRepository structRepo;


        public StructureController(IStructRepository structRepository)
        {
            this.structRepo = structRepository;

        }
        // GET: Structure
        public ActionResult Index()
        {
            

            return View();
        }


        public ActionResult Details()
        {
            IEnumerable<Struct> filSt = structRepo.Structs.ToList();

            return PartialView("_Table", filSt);

        }

        [HttpGet]
        public ActionResult Create()
        {


            Struct model = new Struct();
            return PartialView("_Create", model);


        }

        [HttpPost]
        public ActionResult Create(Struct structF)
        {
            try
            {

               

                if (ModelState.IsValid)
                {


                    if (structF.ID > 0)
                    {
                        ViewBag.Action = "Edited";
                    }
                    else
                    {
                        ViewBag.Action = "Created";
                    }

                    structRepo.SaveStruct(structF);
                    return PartialView("_StructEditSuccess");
                }
                else
                {
                    return PartialView("_Create", structF);
                }

            }
            catch (Exception ex)
            {
                ModelState.AddModelError("Server side error occurred", ex.Message.ToString());
                return PartialView("_Create", structF);
            }
        }

        // GET: Field/Edit/5
        public ActionResult Edit(int id)
        {
            try
            {
                if (id >= 0)
                {
                    Struct structF = structRepo.Structs.Where(f => f.ID == id).FirstOrDefault();
                    if (structF != null)
                    {
                        return PartialView("_Create", structF);
                    }

                }
                return PartialView("_Create");

            }
            catch (Exception ex)
            {

                return View(ex.Message);

            }


        }

        // GET: Field/Delete/5
        [HttpPost]
        public ActionResult Delete(int id)
        {
            try
            {
                if (id >= 0)
                {
                    Struct structF = structRepo.Structs.Where(f => f.ID == id).FirstOrDefault();
                    if (structF != null)
                    {
                        structRepo.DeleteStruct(structF.ID);
                    }

                }
                return Json("ok");

            }
            catch (Exception ex)
            {

                return View(ex.Message);

            }
        }
    }
}