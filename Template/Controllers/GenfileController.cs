using FileGenerator.Domain.Abstract;
using FileGenerator.Domain.Entities;
using FileGenerator.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Syncfusion.XlsIO;
using System.Data;
using System.IO;
using System.Net;
using ExcelDataReader;


namespace Template.Controllers
{
    public class GenfileController : Controller
    {
        private IFieldsRepository fieldsRepo;
        private IDataFieldRepository datafieldRepo;
        private ILFileRepository docRepo;
        private IStructRepository structRepo;
        private IStructFieldRepository structFieldRepo;

        public int Level2 = 0;


        public class InValue
        {

            public int Field { get; set; }

        }

        public class ExcellExp
        {

            public string Value { get; set; }

        }

        public struct Document
        {
            public List<Dictionary<string, string>> Doc;
        }

        public class VPercent
        {
            public string Name { get; set; }
            public string Value { get; set; }
            public int Counter { get; set; }
        }

        public IEnumerable<string> ExceptValues = new List<string>() { "ORDER_NO", "SHIP_TO_ADDRESS LINE2", "BILL TO_ADDRESS_LINE2", "SHIP_TO_ZIP_EXT", "BILL TO_ZIP_EXT" };

        public GenfileController(IFieldsRepository fieldRepository, IDataFieldRepository datafieldRepository, ILFileRepository docRepository, IStructRepository structRepository, IStructFieldRepository structFieldRepository)
        {
            this.fieldsRepo = fieldRepository;
            this.datafieldRepo = datafieldRepository;
            this.docRepo = docRepository;
            this.structRepo = structRepository;
            this.structFieldRepo = structFieldRepository;

        }
        // GET: Genfile
        public ActionResult Index()
        {
            ViewBag.Organization = datafieldRepo.DataFields.Where(p => p.Field.Field_Name == "ORGANIZATION_ID").ToList();

            ViewBag.Items = fieldsRepo.Fields.ToList();

            ViewBag.Bussiness = datafieldRepo.DataFields.Where(p => p.Field.Field_Name == "BUSINESS_UNIT").ToList();

            ViewBag.OrderT = datafieldRepo.DataFields.Where(p => p.Field.Field_Name == "ORDER_TYPE").ToList();

            ViewBag.Document = docRepo.LFiles.ToList();



            return View("Generate");
        }

       

        //Combo box show ajax
        [HttpPost]
        public ActionResult Comboshow(InValue paramets)
        {

            List<DataField> fieldsval = new List<DataField>();

            fieldsval = datafieldRepo.DataFields.Where(p => p.FieldID == paramets.Field).ToList();



            return Json(fieldsval);

        }

        [HttpPost]
        public ActionResult Nameshow(InValue mod)
        {
            int param = Convert.ToInt32(mod.Field);

            Field val = new Field();

            val = fieldsRepo.Fields.Where(p => p.ID == param).FirstOrDefault();



            return Json(val.Field_Name);

        }


        //=================Create File 2 generate================================================================

        [HttpPost]
        public ActionResult Generate(FileGenViewModel model)
        {
            LFile document = docRepo.LFiles.Where(p => p.LFile_ID == model.DocID).FirstOrDefault();

            List<Struct> structs = structRepo.Structs.Where(s => s.LFile_ID == model.DocID).ToList();

            List<Document> document_out = new List<Document>(); // creating the document output list
            List<string> supra_duplicated;

            //getting filter values 

            List<Elements> filter = model.FiltValues;

            List<VPercent> filtPerc = new List<VPercent>();

            if (filter != null) {


                foreach (var fl in filter)
                {

                    if (Convert.ToInt32(fl.Percent) < 100)
                    {

                        VPercent val = new VPercent
                        {
                            Name = fl.Name,
                            Value = fl.Val,
                            Counter = (int)Math.Truncate(Convert.ToDecimal((Convert.ToInt32(fl.Percent) * 0.01) * (model.NDocs * model.NBatch)))
                        };

                        filtPerc.Add(val);
                    }

                }
            }



            int itnum = 0;
            Level2 = 0;
            //iterating over batch number

            for (int l = 0; l < model.NBatch; l++) {

                supra_duplicated = new List<string>();

                for (int i = 0; i < model.NDocs; i++)
                {  // iterating over the number of documents
                    int order_line = 0;
                    itnum++;
                    


                    Document new_doc = new Document();

                    List<Dictionary<string, string>> doc = new List<Dictionary<string, string>>();

                    new_doc.Doc = doc;


                    List<string> duplicated = new List<string>();

                    foreach (var st in structs)
                    {  // checking the structs


                        List<StructField> structfields = structFieldRepo.StructFields.Where(s => s.StructID == st.ID).OrderBy(k => k.Field_Order).ToList();

                        Dictionary<string, string> str_formed = new Dictionary<string, string>();

                        foreach (var m in structfields)
                        {

                            str_formed.Add(m.Field.Field_Name, "");// filling the created structure with the document Struc-Fields

                        };



                        if (st.Order_In_Doc == 1) // if structure is document header 
                        {
                            if (str_formed.ContainsKey("RECORD TYPE"))
                            {
                                str_formed["RECORD TYPE"] = "H";
                            }

                            

                            if (model.FiltValues != null)
                            {

                                foreach (var par in model.FiltValues)
                                { // searching in the params from input

                                    if (par.Type == "text")
                                    {

                                        if (str_formed.ContainsKey(par.Name))
                                        {

                                            str_formed[par.Name] = par.Val;

                                        }


                                    }
                                    else
                                    {

                                        int value1 = Convert.ToInt32(par.Val);

                                        DataField namefilter = datafieldRepo.DataFields.Where(m => m.ID == value1).FirstOrDefault();

                                        if ((namefilter != null) && (str_formed.ContainsKey(namefilter.Field.Field_Name)))
                                        { // if param is in header is assigned the incoming value

                                            str_formed[namefilter.Field.Field_Name] = par.Name;

                                        }
                                    }
                                }
                               // FillFromDBLinked(ref str_formed);
                            }



                            
                            FillFromDB(ref str_formed, itnum, ref duplicated);
                            UpdCounters(ref filtPerc);
                            UpdFilter(ref filter, ref filtPerc, ref supra_duplicated);

                            new_doc.Doc.Add(str_formed);// adding the header to the document



                        }

                        else if ((st.Order_In_Doc < document.Num_Struct) && (st.Order_In_Doc > 1) && st.Multiple)
                        {// if struct is detail

                            if (model.NDets == 0)
                            {


                                model.NDets = 1;

                            }

                            var detnumb = model.NDets;
                            Random rnd = new Random();


                            if (model.Max)
                            {

                                detnumb = rnd.Next(1, model.NDets);

                            }


                            for (int j = 0; j < detnumb; j++) // iterating over the number of details
                            {
                                Dictionary<string, string> new_detail = new Dictionary<string, string>(str_formed);

                                if (new_detail.ContainsKey("RECORD TYPE"))
                                {
                                    new_detail["RECORD TYPE"] = "D";
                                }

                                if (new_detail.ContainsKey("ORDER LINE NO"))
                                {
                                    order_line++;
                                    new_detail["ORDER LINE NO"] = order_line.ToString();

                                }

                                if (new_doc.Doc.Count() > 0)
                                { // if there is already a header

                                    var header = new_doc.Doc[0];

                                    foreach (var word in str_formed)
                                    {  // checking header and filling detail with header values 

                                        if (header.ContainsKey(word.Key) && (new_detail[word.Key] == ""))
                                        {

                                            new_detail[word.Key] = header[word.Key];

                                        }

                                    }



                                }

                                if (model.FiltValues != null)
                                {

                                    foreach (var par in model.FiltValues)
                                    { // searching in the params from input

                                        if (str_formed.ContainsKey(par.Name))
                                        { // if param in detail is assigned the incoming value

                                            new_detail[par.Name] = par.Val;

                                        }

                                    }
                                }
                               // FillFromDBLinked(ref new_detail);

                                FillFromDB(ref new_detail, itnum, ref supra_duplicated);

                                new_doc.Doc.Add(new_detail);// adding the new detail to the document


                            }

                        }
                        else
                        { // if struct is Comment

                            Dictionary<string, string> new_comment = new Dictionary<string, string>(str_formed);

                            if (new_comment.ContainsKey("RECORD TYPE"))
                            {
                                new_comment["RECORD TYPE"] = "C";
                            }
                            if (new_comment.ContainsKey("ORDER LINE NO"))
                            {
                                new_comment["ORDER LINE NO"] = order_line.ToString();
                            }

                            if (new_doc.Doc.Count() > 0)
                            { // if there is already a header

                                var header = new_doc.Doc[0];

                                foreach (var word in str_formed)
                                {  // checking header and filling comment with header values 

                                    if (header.ContainsKey(word.Key) && (new_comment[word.Key] == ""))
                                    {

                                        new_comment[word.Key] = header[word.Key];

                                    }

                                }



                            }
                            if (new_doc.Doc.Count() > 1)
                            { // if there is already a detail

                                var detail = new_doc.Doc[1];

                                foreach (var word in str_formed)
                                {  // checking header and filling comment with header values 

                                    if (detail.ContainsKey(word.Key) && (new_comment[word.Key] == ""))
                                    {

                                        new_comment[word.Key] = detail[word.Key];

                                    }

                                }



                            }


                            if (model.FiltValues != null)
                            {

                                foreach (var par in model.FiltValues)
                                { // searching in the params from input

                                    if (str_formed.ContainsKey(par.Name))
                                    { // if param in detail is assigned the incoming value

                                        new_comment[par.Name] = par.Val;

                                    }

                                }
                            }
                           // FillFromDBLinked(ref new_comment);
                            FillFromDB(ref new_comment, itnum, ref duplicated);

                            new_doc.Doc.Add(new_comment);// adding the new detail to the document

                        }




                    }

                    document_out.Add(new_doc);

                }

                supra_duplicated = null;
            }


            // end of document iteration


            // generating excel
            int counter = 0;



           

            //excel generation 2
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Set the default application version as Excel 2016.
                excelEngine.Excel.DefaultVersion = ExcelVersion.Excel2013;
                //Create a workbook with a worksheet.
                IWorkbook workbook = excelEngine.Excel.Workbooks.Create(1);

                //Access first worksheet from the workbook instance.
                IWorksheet worksheet = workbook.Worksheets[0];
               
                int rownum = 0;
                int detnum = 0;
                foreach (var docm in document_out)
                {
                    counter++;
                    string cell = "";
                    foreach (var row in docm.Doc)
                    {




                        string[] list = new string[row.Count()];

                        int aux = 0;

                        if ((row.ContainsKey("RECORD TYPE") && (row["RECORD TYPE"] != "D") && (detnum > 0)))
                        {
                            detnum = 0;


                        }

                        if (detnum < 1)
                        {



                            foreach (var field in row)
                            {

                                list[aux] = field.Key;
                                aux++;
                            }
                            
                            rownum++;
                            

                            worksheet.ImportArray(list,rownum, 1, false);
                            



                            aux = 0;

                        }

                        if ((row.ContainsKey("RECORD TYPE") && (row["RECORD TYPE"] == "D")))
                        {
                            detnum++;


                        }
                        else
                        {
                            detnum = 0;
                        }

                        foreach (var field in row)
                        {

                            list[aux] = field.Value;
                            aux++;
                        }

                        rownum++;
                        cell = "M" + rownum;

                        if (list.Count() > 12) {
                            string test_elem = list[12];

                            if (test_elem.Substring(0, 1) == "0")
                            {
                               
                                string mask = "";
                                for (int k = 0; k < test_elem.Length; k++) {
                                    mask += "0";
                                }
                                worksheet.Range[cell].NumberFormat = mask;
                            }
                            else
                            {
                                worksheet.Range[cell].NumberFormat = "0";
                            }
                        }
                        

                       
                        
                        worksheet.ImportArray(list, rownum, 1, false);
                        worksheet.AutofitRow(rownum);


                    }

                    //rownum ++2;

                }

                string fileName = "";

                if (model.FileName != null)
                {

                    fileName = "C:\\tkfile\\" + model.FileName + counter + ".xlsx";

                }
                else
                {

                    fileName = "C:\\tkfile\\" + counter + ".xlsx";

                }


                workbook.SaveAs(fileName);
                workbook.Close();
                excelEngine.Dispose();
            }

            //end of excel generation 2

            return PartialView("_GenViewSuccPartial");

        }


        //================= End Create File 2 generate================================================================

        public void FillFromDBLinked(ref Dictionary<string, string> inDict)
        {

            Dictionary<string, string> auxDicc = new Dictionary<string, string>(inDict);

            foreach (var word in inDict)
            {

                   

                    DataField data_field = datafieldRepo.DataFields.Where(p => p.Data == word.Value && p.Field.Field_Name == word.Key&&p.Data!="").FirstOrDefault();

                    if ((data_field!=null)&&(data_field.Field.FLevel > 1) && (data_field.Link_S != null) &&(Level2 == 0))
                    {
                        Level2 = Convert.ToInt32(data_field.Link_S);
                    }

                if ((data_field != null) && (data_field.Link_P != null))
                {

                    List<DataField> linked_vals = datafieldRepo.DataFields.Where(m => m.Link_P == data_field.Link_P && m.Link_S == data_field.Link_S).ToList();// if there are linked values we retrieve them

                    if (linked_vals.Count() > 0)
                    { // try to find the linked values in the dictionary

                        foreach (var k in linked_vals)
                        {
                            if (inDict.ContainsKey(k.Field.Field_Name) && (auxDicc[k.Field.Field_Name] == ""))
                            { // if Key found and no value added

                                auxDicc[k.Field.Field_Name] = k.Data;

                            }


                        }

                    }

                } else if ((data_field != null) && (data_field.Link_S != null)) {

                    List<DataField> linked_vals = datafieldRepo.DataFields.Where(m => m.Link_S == data_field.Link_S).ToList();// if there are linked values we retrieve them
                    if (linked_vals.Count() > 0)
                    { // try to find the linked values in the dictionary

                        foreach (var k in linked_vals)
                        {
                            if (inDict.ContainsKey(k.Field.Field_Name) && (auxDicc[k.Field.Field_Name] == ""))
                            { // if Key found and no value added

                                auxDicc[k.Field.Field_Name] = k.Data;

                            }


                        }

                    }

                }

               

               

            }

            inDict = auxDicc;
        }


        public void FillFromDB(ref Dictionary<string, string> inDict, int iteration, ref List<string> duplicated)
        {

            Dictionary<string, string> auxDicc = new Dictionary<string, string>(inDict);

            

            foreach (var word in inDict)
            {

                DataField data_field = datafieldRepo.DataFields.Where(p => p.Data == word.Value && p.Field.Field_Name == word.Key && p.Data != "").FirstOrDefault();

                if ((data_field != null) && (data_field.Field.FLevel > 1) && (data_field.Link_S != null) && (Level2 == 0))
                {
                    Level2 = Convert.ToInt32(data_field.Link_S);
                }

                if ((word.Value == "") && (auxDicc[word.Key] == "")&&(!ExceptValues.Contains(word.Key)))
                {

                    var rand = new Random();
                    var distC = inDict["BUSINESS_UNIT"];
                   
                    IEnumerable<DataField> insValue = datafieldRepo.DataFields.Where(p => p.Field.Field_Name == word.Key).AsEnumerable();
                    int amount = insValue.Count();
                    Field field = fieldsRepo.Fields.Where(p => p.Field_Name == word.Key).FirstOrDefault();

                    

                    if (amount > 0)
                    {
                        var Slink = insValue.FirstOrDefault().Link_S;
                        if ((field.FLevel != 1)&&(Slink != null) &&(Level2==0)) {

                            Level2 = Convert.ToInt32(Slink);

                        }
                        if (field.UniqueV)
                        {



                            Boolean flag = false;

                            List<DataField> foundData = new List<DataField>();

                            List<DataField> newData = new List<DataField>();

                            int linking = 0;

                            if (Level2 != 0) {

                                newData = insValue.Where(s => s.Link_P == Level2).ToList();

                            }

                            if (newData.Count() > 0)
                            {
                                amount = newData.Count();

                                while (!flag)
                                {


                                    DataField valor = newData.Skip(rand.Next(0, amount)).First();


                                    if (!duplicated.Contains(valor.Data))
                                    {

                                        duplicated.Add(valor.Data);
                                        auxDicc[word.Key] = valor.Data;
                                        linking = Convert.ToInt32(valor.Link_S);
                                        flag = true;

                                    }
                                    if (!foundData.Contains(valor))
                                    {

                                        foundData.Add(valor);
                                    }
                                    if (foundData.Count() == newData.Count())
                                    {

                                        flag = true;

                                    }


                                }

                               

                            }
                            else {


                                while (!flag)
                                {


                                    DataField valor = insValue.Skip(rand.Next(0, amount)).First();


                                    if (!duplicated.Contains(valor.Data))
                                    {

                                        duplicated.Add(valor.Data);
                                        auxDicc[word.Key] = valor.Data;
                                        flag = true;
                                        linking = Convert.ToInt32(valor.Link_S);

                                    }
                                    if (!foundData.Contains(valor))
                                    {

                                        foundData.Add(valor);
                                    }
                                    if (foundData.Count() == insValue.Count())
                                    {

                                        flag = true;

                                    }


                                }


                            }
                            FillFromDBLinked2(ref auxDicc, linking);

                          //  FillFromDBLinked(ref auxDicc);

                        }
                        else
                        {

                            Boolean flag = false;
                            int iterator = 0;

                            while (!flag && iterator < amount) {
                                DataField valor = insValue.Skip(rand.Next(0, amount)).First();

                                if (!duplicated.Contains(valor.Data))
                                {
                                    auxDicc[word.Key] = valor.Data;
                                    if ((valor.Link_S != null)&&(valor.Field.FLevel ==1))
                                    {
                                        FillFromDBLinked2(ref auxDicc, (int)valor.Link_S);
                                    }
                                    flag = true;
                                }
                                iterator++;
                            }

                            


                           
                           
                        }




                    }
                    else
                    {

                        Field field2 = fieldsRepo.Fields.Where(k => k.Field_Name == word.Key).FirstOrDefault();

                        /*  if (field2.Field_Type == "String")
                          {

                             // auxDicc[word.Key] = word.Key.Substring(0, 2) + iteration;

                          }
                          else
                          {
                            //  Random rnd = new Random();
                             // int month = rnd.Next(1, 900);
                             // auxDicc[word.Key] = month.ToString();

                          }*/
                        auxDicc[word.Key] = "";

                    }


                }
                else if (word.Key == "ORDER_NO")
                {

                    if (inDict.ContainsKey("BUSINESS_UNIT") && (inDict["BUSINESS_UNIT"] != ""))
                    {
                        auxDicc[word.Key] = inDict["BUSINESS_UNIT"].Substring(0, 4) + "-" + iteration;

                    }
                    else
                    {
                        auxDicc[word.Key] = word.Value;


                    }
                    
                }
                
                inDict = auxDicc;
               
            }




        }

        public void UpdFilter(ref List<Elements> fil,ref List<VPercent> filtPerc, ref List<string> supra_duplicated) {

            if (filtPerc.Count > 0)
            {

                List<VPercent> aux =  new List<VPercent>(filtPerc);
                


                foreach (var el in aux)
                {
                    if (el.Counter == 0)
                    {
                        supra_duplicated.Add(el.Value);
                        var erem = fil.Find(p => p.Name == el.Name);
                        fil.Remove(erem);
                        int index = filtPerc.FindIndex(p => p.Name == el.Name);
                        filtPerc.RemoveAt(index);

                    }
                }

            }



        }

        public void UpdCounters(ref List<VPercent> filtPc)
        {

            if (filtPc.Count > 0) {

                var aux = filtPc;

                foreach (var el in aux) {
                    if (el.Counter > 0) {
                        int index = filtPc.FindIndex(p => p.Name == el.Name);
                        filtPc[index].Counter--;

                    }
                }

            }



        }

        //Fill from db linked 2

        public void FillFromDBLinked2(ref Dictionary<string, string> inDict, int LinkS)
        {

          
            List<DataField> filter = datafieldRepo.DataFields.Where(p => p.Link_S == LinkS).ToList();

            foreach (var word in filter)
            {

                if (inDict.ContainsKey(word.Field.Field_Name) && inDict[word.Field.Field_Name] == "") {

                    inDict[word.Field.Field_Name] = word.Data;

                }

            }

           
        }


    }
}