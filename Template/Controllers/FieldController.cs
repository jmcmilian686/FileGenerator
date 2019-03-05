using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using FileGenerator.Domain.Entities;
using FileGenerator.Domain.Abstract;
using FileGenerator.Domain.Concrete;
using System.Text.RegularExpressions;

namespace Template.Controllers
{
    public class FieldController : Controller
    {
        private IFieldsRepository fieldsRepo;


        public FieldController(IFieldsRepository fieldRepository)
        {
            this.fieldsRepo = fieldRepository;

        }


        // GET: Field
        public ActionResult Index()
        {
           
           

            return View();
        }

        // GET: Field/Details/5
        public ActionResult Details()
        {
            IEnumerable<Field> fields = fieldsRepo.Fields.ToList();
            return PartialView("_Table",fields);
        }

        // GET: Field/Create
      
        [HttpGet]
        public ActionResult Create()
        {
            

                Field model = new Field();
                return PartialView("_Create", model);
         
            
        }

        // POST: Field/Create
        [HttpPost]
        public ActionResult Create(Field field)
        {
            try
            {

                if (field != null) {

                    switch (field.Field_Type)
                    {
                        case "Number":

                            if (field.DefaultV != null) {

                                if (!Regex.IsMatch(field.DefaultV, @"^\d+$")) {
                                    ModelState.AddModelError("Default", "The Default value must be a number");
                                }
                            }

                            break;
                        case "Float":
                            if (field.DefaultV != null)
                            {

                                if (!Regex.IsMatch(field.DefaultV, @"^\d+$"))
                                {
                                    ModelState.AddModelError("Default", "The Default value must be a float value");
                                }
                            }

                            break;

                        default:
                            break;
                    }

                }

                if (ModelState.IsValid)
                {

                    
                    if (field.ID > 0)
                    {
                        ViewBag.Action = "Edited";
                    }
                    else {
                        ViewBag.Action = "Created";
                    }

                    fieldsRepo.SaveField(field);
                    return PartialView("_FieldEditSuccess");
                }
                else {
                    return PartialView("_Create",field);
                }

            }
            catch (Exception ex)
            {
                ModelState.AddModelError("Server side error occurred", ex.Message.ToString());
                return PartialView("_Create",field);
            }
        }

        // GET: Field/Edit/5
        public ActionResult Edit(int id)
        {
            try
            {
                if (id >= 0) {
                    Field field = fieldsRepo.Fields.Where(f => f.ID == id).FirstOrDefault();
                    if (field != null) {
                        return PartialView("_Create", field);
                    }

                }
                return PartialView("_Create");

            }
            catch(Exception ex) {

                return View(ex.Message);

            }

            
        }

        // POST: Field/Edit/5
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


        // GET: Field/Delete/5
        [HttpPost]
        public ActionResult Delete(int id)
        {
            try
            {
                if (id >= 0)
                {
                    Field field = fieldsRepo.Fields.Where(f => f.ID == id).FirstOrDefault();
                    if (field != null)
                    {
                        fieldsRepo.DeleteField(field.ID);
                    }

                }
                return Json("ok");

            }
            catch (Exception ex)
            {

                return View(ex.Message);

            }
        }




        // POST: Field/Delete/5
       
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
