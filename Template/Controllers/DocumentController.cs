using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using FileGenerator.Domain.Abstract;
using FileGenerator.Domain.Entities;
using FileGenerator.Models;

namespace Template.Controllers
{
    public class DocumentController : Controller
    {
        private IStructFieldRepository strFieldsRepo;
        public IStructRepository structRepo;
        public IFieldsRepository fieldRepo;
        public ILFileRepository lfileRepo;


        public DocumentController(IStructFieldRepository sTfieldRepository,
                                     IStructRepository structReposit,
                                     IFieldsRepository fieldReposit,
                                     ILFileRepository lfileReposit
            )
        {
            this.strFieldsRepo = sTfieldRepository;
            this.structRepo = structReposit;
            this.fieldRepo = fieldReposit;
            this.lfileRepo = lfileReposit;
        }
        // GET: Document
        public ActionResult Index(int? id=0)
        {
            List<LFileViewModel> model = new List<LFileViewModel>();

            List<LFile> lfiles = lfileRepo.LFiles.ToList();

            foreach (var n in lfiles) { // we iterate over all the documents

                LFileViewModel mod = new LFileViewModel
                {
                    Document = n
                };

                List<Struct> structs = structRepo.Structs.Where(s=>s.LFile_ID== n.LFile_ID).ToList(); // and all the structures linked to every document

                AddRowViewModel  rowview = new AddRowViewModel();
                List<StructPosViewModel> strpos = new List<StructPosViewModel>();

                foreach (var st in structs) { //we create the list of StructPosViewModel for each doc
                   
                    StructPosViewModel strpvm = new StructPosViewModel
                    {
                        ID = st.ID,
                        Name = st.St_Name,
                        Description = st.St_Description,
                        Order = (int)st.Order_In_Doc
                    };

                    strpos.Add(strpvm);


                };

                rowview.NewStId = 0;
                rowview.Posit = 0;
                rowview.StructurePos = strpos;


               
                mod.DocStructs = rowview;
                model.Add(mod);
            }
            if (id != 0)
            {

                ViewBag.Element = id;
            }
            else if (model.Count() > 0)
            {

                ViewBag.Element = model[0].Document.LFile_ID;
            }
            else {

                ViewBag.Element = 0;
            }
            ViewBag.Action = "index";
            return View(model);
        }

        [HttpGet]
        public ActionResult Create()
        {

            LFileViewModel lfvmodel = new LFileViewModel();

           AddRowViewModel spmodel = new AddRowViewModel();

            List<StructPosViewModel> listStr = new List<StructPosViewModel>();

            spmodel.StructurePos = listStr;

            lfvmodel.DocStructs = spmodel;
           

           

            return PartialView("_Create", lfvmodel);


        }


        [HttpPost]
        public ActionResult Create(LFileViewModel inbModel)
        {
            try
            {
                if (ModelState.IsValid)
                {

                    if (inbModel.Document.LFile_ID > 0)
                    {     //if the document already exist

                        lfileRepo.SaveLFile(inbModel.Document);
                        List<Struct> strInDoc = structRepo.Structs.Where(p => p.LFile_ID == inbModel.Document.LFile_ID).ToList();

                        IEnumerable<StructPosViewModel> st_inp = inbModel.DocStructs.StructurePos.AsEnumerable();

                        foreach (var p in strInDoc)
                        {

                           

                            var element = st_inp.Where(k => k.ID == p.ID).FirstOrDefault();

                            if (element != null)
                            {
                                p.Order_In_Doc = element.Order;
                                structRepo.SaveStruct(p);

                            }
                            else {

                                p.LFile_ID = 0;
                                p.Order_In_Doc = 0;
                                structRepo.SaveStruct(p);

                            }

                        }

                        foreach (var inp_element in inbModel.DocStructs.StructurePos) {

                            Struct new_value = structRepo.Structs.Where(k => k.ID == inp_element.ID).FirstOrDefault();

                            if (new_value.LFile_ID != inbModel.Document.LFile_ID) {

                                new_value.LFile_ID = inbModel.Document.LFile_ID;
                                structRepo.SaveStruct(new_value);

                            }


                        }
                    }
                    else {// document is new

                        lfileRepo.SaveLFile(inbModel.Document);

                        foreach (var n in inbModel.DocStructs.StructurePos) {

                            Struct str = structRepo.Structs.Where(k => k.ID == n.ID).FirstOrDefault();

                            str.LFile_ID = inbModel.Document.LFile_ID;
                            str.Order_In_Doc = n.Order;
                            structRepo.SaveStruct(str);
                        }


                    }


                    if (inbModel.Document.LFile_ID > 0)
                    {
                        ViewBag.Action = "Edited";
                    }
                    else
                    {
                        ViewBag.Action = "Created";
                    }

                    return PartialView("_DocumEditSuccess");
                }
                else
                {
                   
                    return PartialView("_Create", inbModel);
                }

            }
            catch(Exception ex) {

                return new HttpStatusCodeResult(404, "Error in cloud - GetPLUInfo" + ex.Message);


            }
           


        }

        public ActionResult Edit(int id) {

            LFileViewModel model = new LFileViewModel();
            LFile doc = lfileRepo.LFiles.Where(p => p.LFile_ID == id).FirstOrDefault();

            model.Document = doc;

            List<Struct> structs = structRepo.Structs.Where(s => s.LFile_ID == doc.LFile_ID).ToList(); // and all the structures linked to every document

            AddRowViewModel rowview = new AddRowViewModel();
            List<StructPosViewModel> strpos = new List<StructPosViewModel>();

            foreach (var st in structs)
            { //we create the list of StructPosViewModel for each doc

                StructPosViewModel strpvm = new StructPosViewModel
                {
                    ID = st.ID,
                    Name = st.St_Name,
                    Description = st.St_Description,
                    Order = (int)st.Order_In_Doc
                };

                strpos.Add(strpvm);
                
            };
            ViewBag.Action = "edit";
            ViewBag.FieldName = structRepo.Structs.AsEnumerable();
            rowview.NewStId = 0;
            rowview.Posit = 0;
            rowview.StructurePos = strpos;
            model.DocStructs = rowview;

            return PartialView("_Create", model);

        }

        [HttpGet]
        public ActionResult CreatePos()
        {

            

            AddRowViewModel spmodel = new AddRowViewModel();

            List<StructPosViewModel> listStr = new List<StructPosViewModel>();
            spmodel.StructurePos = listStr;

            ViewBag.FieldName = structRepo.Structs.AsEnumerable();



            return PartialView("_StrPos", spmodel);


        }


        

        [HttpPost]
        public ActionResult AddPos(AddRowViewModel addModel)
        {
            //Eliminate from the dropdown list the elements already in the table.

            IEnumerable<Struct> viewBL = structRepo.Structs.AsQueryable();
           

            IEnumerable<StructPosViewModel> structures = addModel.StructurePos.AsEnumerable();
            List<Struct> laux = viewBL.ToList();
            /*
            if (structures != null)
            {

                foreach (StructPosViewModel n in structures)
                {
                    var element = viewBL.Where(p => p.ID == n.ID).FirstOrDefault();
                    if (element != null)
                    {

                        laux.Remove(element);

                    }

                }
                ViewBag.FieldName = laux.AsEnumerable();

            }
            else {

                ViewBag.FieldName = viewBL.AsEnumerable();
            }*/

            ViewBag.FieldName = viewBL.AsEnumerable();




            if ((addModel.StructurePos!=null)&&(addModel.StructurePos.Count() > 0)) // If there are structures in the model
            {

                bool found = false;

                foreach (var n in addModel.StructurePos)
                {
                    if (n.Order == addModel.Posit)
                    {
                        found = true;
                    }
                }



                

                if (found) // if we found a structure with the same Order
                {
                    ModelState.AddModelError("position", "The File Order must be unique, there is already a field at given position");

                    return PartialView("_StrPos", addModel);
                }
               

            }
            

                StructPosViewModel newModel = new StructPosViewModel();

            if (addModel.NewStId == 0)
            {
                addModel.NewStId = 0;
                ModelState.AddModelError("strId", "You must select a Structure");

                return PartialView("_StrPos", addModel);

            }
            else if (addModel.Posit == 0)
            {

                addModel.Posit = 0;
                ModelState.AddModelError("position", "You must select a position");

                return PartialView("_StrPos", addModel);

            }
            else {

                 newModel.Order = (int)addModel.Posit;
                Struct Structure = structRepo.Structs.Where(p => p.ID == addModel.NewStId).FirstOrDefault();
                newModel.ID = Structure.ID;
                newModel.Name = Structure.St_Name;
                newModel.Description = Structure.St_Description;

                List<StructPosViewModel> list = new List<StructPosViewModel>();

                if (addModel.StructurePos != null) {

                    list = addModel.StructurePos.ToList();
               
                }

                list.Insert(0, newModel);

                addModel.StructurePos = list;
                addModel.NewStId = 0;
                addModel.Posit = 0;


                return PartialView("_StrPos", addModel);

            }

            
        }
        // Post: Document/Delete/5
        [HttpPost]
        public ActionResult Delete(int id)
        {
            try
            {
                if (id >= 0)
                {
                    List<Struct> structs = structRepo.Structs.Where(p => p.LFile_ID == id).ToList();

                    foreach (var st in structs) {
                        st.LFile_ID = 0;
                      //  structRepo.SaveStruct(st);
                    }

                    lfileRepo.DeleteLFile(id);

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