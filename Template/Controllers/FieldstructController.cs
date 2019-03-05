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
    public class FieldstructController : Controller
    {

        private IStructFieldRepository strFieldsRepo;
        public IStructRepository structRepo;
        public IFieldsRepository fieldRepo;


        public FieldstructController(IStructFieldRepository sTfieldRepository, 
                                     IStructRepository structReposit,
                                     IFieldsRepository fieldReposit
            )
        {
            this.strFieldsRepo = sTfieldRepository;
            this.structRepo = structReposit;
            this.fieldRepo = fieldReposit;
        }
        // GET: Fieldstruct
        public ActionResult Index()
        {
            return View();
        }
        // GET: Field/Details/5
        public ActionResult Details()
        {
            IEnumerable<StructField> sTfields = strFieldsRepo.StructFields.ToList();

            
            return PartialView("_Table", sTfields);
        }

        [HttpGet]
        public ActionResult Create()
        {

            
            StructField model = new StructField();
            /*IEnumerable<Struct> strcts = new List<Struct>();
            strcts = structRepo.Structs.AsEnumerable();
            ViewBag.StructName = strcts;
            IEnumerable<Field> fields = new List<Field>();
            fields = fieldRepo.Fields.AsEnumerable();
            ViewBag.FieldName = fields;*/

            ViewBag.StructName = structRepo.Structs.AsEnumerable();
            ViewBag.FieldName = fieldRepo.Fields.AsEnumerable();

            return PartialView("_Create", model);


        }

        // POST: Field/Create
        [HttpPost]
        public ActionResult Create(StructField sTfield)
        {
            ViewBag.StructName = structRepo.Structs.AsEnumerable();
            ViewBag.FieldName = fieldRepo.Fields.AsEnumerable();
            try
            {
                    if (sTfield.ID > 0)
                    {
                         ViewBag.Action = "Edited";
                    }
                    else
                    {
                        //if there is a field for the same struct with the same order
                        if (strFieldsRepo.StructFields.Where(p => p.StructID == sTfield.StructID && p.Field_Order == sTfield.Field_Order).FirstOrDefault() != null)
                        {
                            ModelState.AddModelError("Field_Order", "The File Order must be unique, here is already a field at given possition");
                        };

                         ViewBag.Action = "Created";
                    }

               


                if (ModelState.IsValid)
                {
                   


                    strFieldsRepo.SaveStructField(sTfield);
                    return PartialView("_FSEditSuccess");
                }
                else
                {
                    return PartialView("_Create", sTfield);
                }

            }
            catch (Exception ex)
            {
                ModelState.AddModelError("Server side error occurred", ex.Message.ToString());
                return PartialView("_Create", sTfield);
            }
        }

        [HttpPost]
        public ActionResult Edit(int id)
        {
            ViewBag.StructName = structRepo.Structs.AsEnumerable();
            ViewBag.FieldName = fieldRepo.Fields.AsEnumerable();
            try
            {
                if (id >= 0)
                {
                    StructField sfield = strFieldsRepo.StructFields.Where(f => f.ID == id).FirstOrDefault();
                    if (sfield != null)
                    {
                        return PartialView("_Create", sfield);
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
                    StructField Sfield = strFieldsRepo.StructFields.Where(f => f.ID == id).FirstOrDefault();
                    if (Sfield != null)
                    {
                        strFieldsRepo.DeleteStructField(Sfield.ID);
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