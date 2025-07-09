using AddOnGs.Extenstion;
using AddOnGs.Models.OpenAPI.Response;
using AddOnGs.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using WolfApprove.API2.Extension;
using static AddOnGs.Models.OpenAPI.Request.ExtCreateMemorandumRequest;
using static AddOnGs.Models.OpenAPI.Response.ExtCreateMemorandumResponse;
using WolfApprove.Model.CustomClass;
using WolfApprove.Model;
using System.Data.Entity;
using Newtonsoft.Json.Linq;
using WolfApprove.API2.Controllers.Utils;
using System.Text.RegularExpressions;

namespace AddOnGs.Service
{
    public class AutoCreated
    {
        private readonly HttpClient _client;
        private readonly string openAPIUrl;
        private NameValueCollection config;

        public AutoCreated(string conn)
        {
            config = Ext.GetAppSetting();
            _client = new HttpClient();
            openAPIUrl = config["OpenAPIUrl"] ?? "";
            _client.BaseAddress = new Uri(openAPIUrl);
        }
        public async Task<bool> AutoGenerate01(AddonFormModel.Form_Model model)
        {
            try
            {
                using (var context = DBContext.OpenConnection(model.memoPage.memoDetail.connectionString))
                {
                    var memoMain = await context.TRNMemoes.Where(e => e.MemoId == model.memoPage.memoDetail.memoid).FirstOrDefaultAsync();
                    if (memoMain.StatusName == "Completed")
                    {
                        Ext.WriteLogFile("Start AutoGenerate SE-F-SHE-020-01");

                        var advancefourm = context.TRNMemoForms.Where(e => e.MemoId == model.memoPage.memoDetail.memoid).ToList();
                        var tbrowcount = advancefourm.Where(e => e.obj_type == "tb").FirstOrDefault();
                        var tbData = advancefourm.Where(e => e.obj_type == "tb").ToList();
                        var Alldata = advancefourm.Where(e => e.obj_type != "tb").ToList();

                        List<CustomTRNFormModel> allList = new List<CustomTRNFormModel>();
                        for (int i = 1; i <= tbrowcount.row_count; i++)
                        {
                            var labellist = new CustomTRNFormModel();
                            var getrowindex = tbData.Where(x => x.row_index == i).ToList();
                            foreach (var itemvalue in getrowindex)
                            {
                                if (i == itemvalue.row_index)
                                {
                                    if (itemvalue.col_label == "LEAD AUDITOR")
                                    {
                                        labellist.LEADAUDITOR = itemvalue.col_value;
                                    }
                                    if (itemvalue.col_label == "AUDITEES")
                                    {
                                        labellist.AUDITEES = itemvalue.col_value;
                                    }
                                }

                            }
                            allList.Add(labellist);
                        }
                        foreach (var dorpvalue in allList)
                        {
                            foreach (var getdata in Alldata)
                            {
                                if (getdata.obj_label == "ครั้งที่ Audit")
                                {
                                    dorpvalue.Audittime = getdata.obj_value;
                                }
                                if (getdata.obj_label == "ปี Audit")
                                {
                                    dorpvalue.YearAudit = getdata.obj_value;
                                }
                                if (getdata.obj_label == "วันที่เริ่ม Audit")
                                {
                                    dorpvalue.Auditstartdate = getdata.obj_value;
                                }
                                if (getdata.obj_label == "วันที่สิ้นสุด Audit")
                                {
                                    dorpvalue.Auditenddate = getdata.obj_value;
                                }

                            }
                            var emps = context.MSTEmployees.Where(e => dorpvalue.LEADAUDITOR.Contains(e.NameEn)).ToList();
                            if (emps.Any())
                            {
                                var documentCode = "SE-F-SHE-020-01";
                                // Fix Url บนเครื่องตัวเองไปก่อน 
                                var getTemplate = await _client.GetAPI($"api/v1/template?documentCode={documentCode}");
                                if (getTemplate.IsSuccessStatusCode)
                                {
                                    var template = await getTemplate.ReadContentAs<ExtGetTemplateResponse>();
                                    var form = template.form;
                                    foreach (var item in template.form)
                                    {
                                        if (item.Template.Label == "ครั้งที่ Audit") item.Data = new Models.OpenAPI.Data { Value = dorpvalue.Audittime };
                                        if (item.Template.Label == "ปี Audit") item.Data = new Models.OpenAPI.Data { Value = dorpvalue.YearAudit };
                                        if (item.Template.Label == "วันที่เริ่ม Audit") item.Data = new Models.OpenAPI.Data { Value = dorpvalue.Auditstartdate };
                                        if (item.Template.Label == "วันที่สิ้นสุด Audit") item.Data = new Models.OpenAPI.Data { Value = dorpvalue.Auditenddate };
                                        if (item.Template.Label == "พื้นที่ตรวจ") item.Data = new Models.OpenAPI.Data { Value = dorpvalue.AUDITEES };
                                        if (item.Template.Label == "ตาราง AUDIT")
                                        {
                                            var getdatatb = tbData.Where(x => x.col_label == "LEAD AUDITOR").ToList();

                                            // สร้าง JSON string
                                            string tablevalue = $"[{string.Join(",", getdatatb.Select(x => x.row_value))}]";

                                            // แปลง JSON string เป็น List<List<Models.OpenAPI.Data>>
                                            var dataList = JsonConvert.DeserializeObject<List<List<Models.OpenAPI.Data>>>(tablevalue);

                                            // กำหนดค่าให้ item.Data
                                            item.Data = new Models.OpenAPI.Data
                                            {
                                                Row = dataList
                                            };
                                        }
                                    }
                                    var jsonString = JsonConvert.SerializeObject(template.form); // แปลงเป็น JSON string
                                    var json = JsonConvert.DeserializeObject(jsonString);
                                    foreach (var emp in emps)
                                    {
                                        // Mapping Data to MAdvance Forum
                                        var modelCreateMemo = new CreateMemorandumRequest
                                        {
                                            Subject = $"พื่นที่ตรวจ " + dorpvalue.AUDITEES + "ครั้งที่ Audit " + dorpvalue.Audittime + " ปี Audit " + dorpvalue.YearAudit + "",
                                            Form = template.form,
                                            UserPrincipalName = emp.Email,
                                            DocumentCode = documentCode,
                                            Action = "draft",
                                        };
                                        var response = await _client.PostAPI($"api/v1/memorandum", modelCreateMemo);
                                        if (response.IsSuccessStatusCode)
                                        {
                                            var data = await response.ReadContentAs<ExtCreateMemorandumResult>();
                                            Ext.WriteLogFile("Generate SOURCING-V Success !");
                                        }
                                    }


                                }
                                else
                                {
                                    Ext.WriteLogFile("Someting Wrong when call Get Template");
                                    return false;
                                }
                            }
                        }
                        return true;

                    }
                    else
                    {
                        // don't process
                        return true;
                    }
                }

            }
            catch (Exception ex)
            {
                Ext.ErrorLog(ex, "AutoGenerate01");
                return false;
            }
        }


        public async Task<bool> AutoGenerate021(AddonFormModel.Form_Model model)
        {
            try
            {
                using (var context = DBContext.OpenConnection(model.memoPage.memoDetail.connectionString))
                {
                    var memoMain = await context.TRNMemoes.Where(e => e.MemoId == model.memoPage.memoDetail.memoid).FirstOrDefaultAsync();
                    if (memoMain.StatusName == "Completed")
                    {
                        Ext.WriteLogFile("Start AutoGenerate SE-F-SHE-021");

                        var advancefourm = context.TRNMemoForms.Where(e => e.MemoId == model.memoPage.memoDetail.memoid).ToList();
                        var tbrowcount = advancefourm.Where(e => e.obj_type == "tb").FirstOrDefault();
                        var tbData = advancefourm.Where(e => e.obj_type == "tb").ToList();
                        var Alldata = advancefourm.Where(e => e.obj_type != "tb").ToList();

                        CustomTRNFormModel allList = new CustomTRNFormModel();


                        foreach (var getdata in Alldata)
                        {
                            if (getdata.obj_label == "ครั้งที่ Audit")
                            {
                                allList.Audittime = getdata.obj_value;
                            }
                            if (getdata.obj_label == "ปี Audit")
                            {
                                allList.YearAudit = getdata.obj_value;
                            }
                            if (getdata.obj_label == "วันที่ตรวจ")
                            {
                                allList.AuditDate = getdata.obj_value;
                            }
                            if (getdata.obj_label == "เวลาที่ตรวจ")
                            {
                                allList.Inspectiontime = getdata.obj_value;
                            }
                            if (getdata.obj_label == "พื้นที่ตรวจ")
                            {
                                allList.AuditArea = getdata.obj_value;
                            }

                        }
                        var emps = context.MSTEmployees.Where(e => memoMain.RNameEn.Contains(e.NameEn)).ToList();
                        if (emps.Any())
                        {
                            var documentCode = "SE-F-SHE-021";
                            // Fix Url บนเครื่องตัวเองไปก่อน 
                            var getTemplate = await _client.GetAPI($"api/v1/template?documentCode={documentCode}");
                            if (getTemplate.IsSuccessStatusCode)
                            {
                                var template = await getTemplate.ReadContentAs<ExtGetTemplateResponse>();
                                var form = template.form;
                                foreach (var item in template.form)
                                {
                                    if (item.Template.Label == "ครั้งที่ Audit") item.Data = new Models.OpenAPI.Data { Value = allList.Audittime };
                                    if (item.Template.Label == "ปี Audit") item.Data = new Models.OpenAPI.Data { Value = allList.YearAudit };
                                    if (item.Template.Label == "วันที่ตรวจ") item.Data = new Models.OpenAPI.Data { Value = allList.AuditDate };
                                    if (item.Template.Label == "เวลาที่ตรวจ") item.Data = new Models.OpenAPI.Data { Value = allList.Inspectiontime };
                                    if (item.Template.Label == "พื้นที่ตรวจ") item.Data = new Models.OpenAPI.Data { Value = allList.AuditArea };
                                    if (item.Template.Label == "Audit Team")
                                    {
                                        var getdatatb = tbData.Where(x => x.col_label == "LEAD AUDITOR" && x.col_value == memoMain.RNameEn).ToList();

                                        JArray jsonArray = new JArray(getdatatb.Select(x => JArray.Parse(x.row_value)));
                                        var gettoken = JToken.FromObject(jsonArray);
                                        var childrenTokens = gettoken.First().Children();
                                        var firstTwoTokens = childrenTokens.Take(2).ToList();
                                        JArray newJsonArray = new JArray(firstTwoTokens);
                                        string resultJson = JsonConvert.SerializeObject(newJsonArray, Formatting.Indented);
                                        string tablevalue = $"[{resultJson}]";
                                        var dataList = JsonConvert.DeserializeObject<List<List<Models.OpenAPI.Data>>>(tablevalue);

                                        item.Data = new Models.OpenAPI.Data
                                        {
                                            Row = dataList
                                        };
                                    }
                                }
                                var jsonString = JsonConvert.SerializeObject(template.form);
                                var json = JsonConvert.DeserializeObject(jsonString);
                                foreach (var emp in emps)
                                {
                                    // Mapping Data to MAdvance Forum
                                    var modelCreateMemo = new CreateMemorandumRequest
                                    {
                                        Subject = $"Audit Checklist ISO 14001/ISO 45001  : พื่นที่ตรวจ " + allList.AUDITEES + "ครั้งที่ Audit " + allList.Audittime + " ปี Audit " + allList.YearAudit + "",
                                        Form = template.form,
                                        UserPrincipalName = emp.Email,
                                        DocumentCode = documentCode,
                                        Action = "draft",
                                    };
                                    var response = await _client.PostAPI($"api/v1/memorandum", modelCreateMemo);
                                    if (response.IsSuccessStatusCode)
                                    {
                                        var data = await response.ReadContentAs<ExtCreateMemorandumResult>();
                                        Ext.WriteLogFile("Generate SOURCING-V Success !");
                                    }
                                }


                            }
                            else
                            {
                                Ext.WriteLogFile("Someting Wrong when call Get Template");
                                return false;
                            }
                        }

                        return true;

                    }
                    else
                    {
                        // don't process
                        return true;
                    }
                }

            }
            catch (Exception ex)
            {
                Ext.ErrorLog(ex, "AutoGenerate021");
                return false;
            }
        }
        public async Task<bool> AutoGenerate009(AddonFormModel.Form_Model model)
        {
            try
            {
                using (var context = DBContext.OpenConnection(model.memoPage.memoDetail.connectionString))
                {
                    var memoMain = await context.TRNMemoes.Where(e => e.MemoId == model.memoPage.memoDetail.memoid).FirstOrDefaultAsync();
                    if (memoMain.StatusName == "Completed")
                    {
                        Ext.WriteLogFile("Start GetData F-QMS-008" + memoMain.MemoId);
                        #region F-QMS-008
                        string AuditReceived_Date = string.Empty;
                        string Part = string.Empty;
                        string AuditTime_Date = string.Empty;
                        string RoundNo = string.Empty;
                        string Area = string.Empty;
                        string Department = string.Empty;
                        List<object> List_Audit_Team = new List<object>();
                        List<object> List_Menu_Audit = new List<object>();
                        JObject jsonAdvanceForm = JsonUtils.createJsonObject(memoMain.MAdvancveForm);
                        JArray itemsArray = (JArray)jsonAdvanceForm["items"];
                        foreach (JObject jItems in itemsArray)
                        {

                            JArray jLayoutArray = (JArray)jItems["layout"];
                            if (jLayoutArray.Count >= 1)
                            {
                                JObject jTemplateL = (JObject)jLayoutArray[0]["template"];
                                JObject jData = (JObject)jLayoutArray[0]["data"];
                                if ((String)jTemplateL["label"] == "วันที่รับการ Audit")
                                {
                                    AuditReceived_Date = jData["value"].ToString();
                                }
                                if ((String)jTemplateL["label"] == "Audit Team")
                                {

                                    foreach (JArray row in jData["row"])
                                    {
                                        List<object> rowObject = new List<object>();
                                        foreach (JObject item in row)
                                        {
                                            if (item.TryGetValue("value", out JToken valueToken) && valueToken != null)
                                            {
                                                rowObject.Add(valueToken.ToString());
                                            }
                                        }
                                        if (rowObject.Count > 0)
                                        {
                                            List_Audit_Team.Add(rowObject);
                                        }
                                    }
                                }
                                if ((String)jTemplateL["label"] == "ส่วน")
                                {
                                    Part = jData["value"].ToString();
                                }
                                if ((String)jTemplateL["label"] == "รายการ AUDIT ข้อกำหนด")
                                {

                                    foreach (JArray row in jData["row"])
                                    {
                                        List<object> rowObject = new List<object>();
                                        foreach (JObject item in row)
                                        {
                                            if (item.TryGetValue("value", out JToken valueToken) && valueToken != null)
                                            {
                                                rowObject.Add(valueToken.ToString());
                                            }
                                        }
                                        if (rowObject.Count > 0)
                                        {
                                            List_Menu_Audit.Add(rowObject);
                                        }
                                    }
                                }
                                if (jLayoutArray.Count > 1)
                                {
                                    JObject jTemplateR = (JObject)jLayoutArray[1]["template"];
                                    JObject jData2 = (JObject)jLayoutArray[1]["data"];
                                    if ((String)jTemplateR["label"] == "เวลาที่รับการ Audit")
                                    {
                                        AuditTime_Date = jData2["value"].ToString();
                                    }
                                    if ((String)jTemplateR["label"] == "รอบที่")
                                    {
                                        RoundNo = jData2["value"].ToString();
                                    }
                                    if ((String)jTemplateR["label"] == "พื้นที่ตรวจ")
                                    {
                                        Area = jData2["value"].ToString();
                                    }
                                    if ((String)jTemplateR["label"] == "ฝ่าย")
                                    {
                                        Department = jData2["value"].ToString();
                                    }
                                }
                            }
                        }
                        Ext.WriteLogFile("End GetData F-QMS-008" + memoMain.MemoId);
                        #endregion
                        #region F-QMS-009
                        var emps = context.MSTEmployees.Where(e => memoMain.RNameEn.Contains(e.NameEn)).ToList();
                        if (emps.Any())
                        {
                            var documentCode = "F-QMS-009";
                            var getTemplate = await _client.GetAPI($"api/v1/template?documentCode={documentCode}");
                            if (getTemplate.IsSuccessStatusCode)
                            {
                                Ext.WriteLogFile("Start AutoGenerate F-QMS-009");
                                var template = await getTemplate.ReadContentAs<ExtGetTemplateResponse>();
                                var form = template.form;
                                foreach (var item in template.form)
                                {
                                    switch (item.Template.Label)
                                    {
                                        case "วันที่รับการตรวจ":
                                            item.Data = new Models.OpenAPI.Data { Value = AuditReceived_Date };
                                            break;
                                        case "เวลาที่รับการตรวจ":
                                            item.Data = new Models.OpenAPI.Data { Value = AuditTime_Date };
                                            break;
                                        case "ส่วน":
                                            item.Data = new Models.OpenAPI.Data { Value = Part };
                                            break;
                                        case "พื้นที่ตรวจ":
                                            item.Data = new Models.OpenAPI.Data { Value = Area };
                                            break;
                                        case "รอบที่":
                                            item.Data = new Models.OpenAPI.Data { Value = RoundNo };
                                            break;
                                        case "ฝ่าย":
                                            item.Data = new Models.OpenAPI.Data { Value = Department };
                                            break;
                                        case "Audit Team":
                                            if (List_Audit_Team.Count > 0)
                                            {
                                                var rowList = new List<List<AddOnGs.Models.OpenAPI.Data>>();
                                                for (int i = 0; i < List_Audit_Team.Count; i++)
                                                {
                                                    dynamic itema = List_Audit_Team[i];
                                                    string Name = itema[0];
                                                    string Position = itema[1];
                                                    var row = new List<AddOnGs.Models.OpenAPI.Data>
                                                    {
                                                        new AddOnGs.Models.OpenAPI.Data { Value = Name },
                                                        new AddOnGs.Models.OpenAPI.Data { Value = Position }
                                                    };
                                                    rowList.Add(row);
                                                }
                                                item.Data = new AddOnGs.Models.OpenAPI.Data { Row = rowList };
                                            }
                                            break;
                                    }
                                }
                                var jsonString = JsonConvert.SerializeObject(template.form);
                                var json = JsonConvert.DeserializeObject(jsonString);
                                foreach (var emp in emps)
                                {
                                    var modelCreateMemo = new CreateMemorandumRequest
                                    {
                                        Subject = $"Audit Checklist ISO 9001/IATF16949 : พื้นที่ตรวจ {Area} รอบที่ {RoundNo}",
                                        Form = template.form,
                                        UserPrincipalName = emp.Email,
                                        DocumentCode = documentCode,
                                        Action = "draft",
                                    };
                                    var response = await _client.PostAPI($"api/v1/memorandum", modelCreateMemo);
                                    if (response.IsSuccessStatusCode)
                                    {
                                        var data = await response.ReadContentAs<ExtCreateMemorandumResult>();
                                        Ext.WriteLogFile("Generate F-QMS-009 Success !");
                                    }
                                }
                            }
                            else
                            {
                                Ext.WriteLogFile("Someting Wrong when call Get Template");
                                Ext.WriteLogFile("----------------------------------------------------------------------------");
                                return false;
                            }
                        }
                        return true;
                        #endregion
                    }
                    else
                    {
                        // don't process
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Ext.ErrorLog(ex, "AutoGenerate009");
                return false;
            }
        }
        public async Task<bool> AutoGenerate010(AddonFormModel.Form_Model model)
        {
            try
            {
                using (var context = DBContext.OpenConnection(model.memoPage.memoDetail.connectionString))
                {
                    var memoMain = await context.TRNMemoes.Where(e => e.MemoId == model.memoPage.memoDetail.memoid).FirstOrDefaultAsync();
                    if (memoMain.StatusName == "Wait for Approve")
                    {
                        var empseq1 = context.TRNLineApproves.Where(x => x.MemoId == memoMain.MemoId && x.Seq == 1).Select(x => x.EmployeeId);
                        if (empseq1.Any(e => e.HasValue && e.Value == memoMain.PersonWaitingId))
                        {
                            #region F-QMS-009
                            Ext.WriteLogFile("Start GetData F-QMS-009" + memoMain.MemoId);
                            string Audit_Date = string.Empty;
                            string Audit_Time = string.Empty;
                            string Part = string.Empty;
                            string RoundNo = string.Empty;
                            string Area = string.Empty;
                            string Department = string.Empty;
                            List<object> List_Audit_Team = new List<object>();
                            List<object> Audit_Checklist = new List<object>();
                            JObject jsonAdvanceForm = JsonUtils.createJsonObject(memoMain.MAdvancveForm);
                            JArray itemsArray = (JArray)jsonAdvanceForm["items"];
                            foreach (JObject jItems in itemsArray)
                            {

                                JArray jLayoutArray = (JArray)jItems["layout"];
                                if (jLayoutArray.Count >= 1)
                                {
                                    JObject jTemplateL = (JObject)jLayoutArray[0]["template"];
                                    JObject jData = (JObject)jLayoutArray[0]["data"];
                                    if ((String)jTemplateL["label"] == "วันที่รับการตรวจ")
                                    {
                                        Audit_Date = jData["value"].ToString();
                                    }
                                    if ((String)jTemplateL["label"] == "เวลาที่รับการตรวจ")
                                    {
                                        Audit_Time = jData["value"].ToString();
                                    }
                                    if ((String)jTemplateL["label"] == "ส่วน")
                                    {
                                        Part = jData["value"].ToString();
                                    }
                                    if ((String)jTemplateL["label"] == "Audit Team")
                                    {

                                        foreach (JArray row in jData["row"])
                                        {
                                            List<object> rowObject = new List<object>();
                                            foreach (JObject item in row)
                                            {
                                                if (item.TryGetValue("value", out JToken valueToken) && valueToken != null)
                                                {
                                                    rowObject.Add(valueToken.ToString());
                                                }
                                            }
                                            if (rowObject.Count > 0)
                                            {
                                                List_Audit_Team.Add(rowObject);
                                            }
                                        }
                                    }
                                    if ((String)jTemplateL["label"] == "Audit Checklist")
                                    {

                                        foreach (JArray row in jData["row"])
                                        {
                                            List<object> rowObject = new List<object>();
                                            foreach (JObject item in row)
                                            {
                                                if (item.TryGetValue("value", out JToken valueToken) && valueToken != null)
                                                {
                                                    rowObject.Add(valueToken.ToString());
                                                }
                                            }
                                            if (rowObject.Count > 0)
                                            {
                                                Audit_Checklist.Add(rowObject);
                                            }
                                        }
                                    }
                                    if (jLayoutArray.Count > 1)
                                    {
                                        JObject jTemplateR = (JObject)jLayoutArray[1]["template"];
                                        JObject jData2 = (JObject)jLayoutArray[1]["data"];
                                        if ((String)jTemplateR["label"] == "พื้นที่ตรวจ")
                                        {
                                            Area = jData2["value"].ToString();
                                        }
                                        if ((String)jTemplateR["label"] == "รอบที่")
                                        {
                                            RoundNo = jData2["value"].ToString();
                                        }
                                        if ((String)jTemplateR["label"] == "ฝ่าย")
                                        {
                                            Department = jData2["value"].ToString();
                                        }
                                    }
                                }
                            }
                            Ext.WriteLogFile("End GetData F-QMS-009" + memoMain.MemoId);
                            #endregion
                            #region F-QMS-010
                            var emps = context.MSTEmployees.Where(e => memoMain.RNameEn.Contains(e.NameEn)).ToList();
                            if (emps.Any())
                            {
                                var documentCode = "F-QMS-010";
                                var getTemplate = await _client.GetAPI($"api/v1/template?documentCode={documentCode}");
                                if (getTemplate.IsSuccessStatusCode)
                                {
                                    Ext.WriteLogFile("Start AutoGenerate F-QMS-010");
                                    var template = await getTemplate.ReadContentAs<ExtGetTemplateResponse>();
                                    var form = template.form;
                                    foreach (var item in template.form)
                                    {
                                        switch (item.Template.Label)
                                        {
                                            case "วันที่ตรวจ":
                                                item.Data = new Models.OpenAPI.Data { Value = Audit_Date };
                                                break;
                                            case "เวลาที่ตรวจ":
                                                item.Data = new Models.OpenAPI.Data { Value = Audit_Time };
                                                break;
                                            case "ส่วน":
                                                item.Data = new Models.OpenAPI.Data { Value = Part };
                                                break;
                                            case "Audit Team":
                                                if (List_Audit_Team.Count > 0)
                                                {
                                                    var rowList = new List<List<AddOnGs.Models.OpenAPI.Data>>();
                                                    for (int i = 0; i < List_Audit_Team.Count; i++)
                                                    {
                                                        dynamic itema = List_Audit_Team[i];
                                                        string Name = itema[0];
                                                        string Position = itema[1];
                                                        var row = new List<AddOnGs.Models.OpenAPI.Data>
                                                        {
                                                            new AddOnGs.Models.OpenAPI.Data { Value = Name },
                                                            new AddOnGs.Models.OpenAPI.Data { Value = Position }
                                                        };
                                                        rowList.Add(row);
                                                    }
                                                    item.Data = new AddOnGs.Models.OpenAPI.Data { Row = rowList };
                                                }
                                                break;
                                            case "ข้อกำหนดที่ตรวจ":
                                                if (Audit_Checklist.Count > 0)
                                                {
                                                    string value = string.Empty;
                                                    for (int i = 0; i < Audit_Checklist.Count; i++)
                                                    {
                                                        dynamic itema = Audit_Checklist[i];
                                                        string Clause = itema[0];
                                                        if (i > 0) { value += ","; }
                                                        value += Clause;
                                                    }
                                                    item.Data = new Models.OpenAPI.Data { Value = value };
                                                }
                                                break;
                                            case "ตารางผลการตรวจ":
                                                if (Audit_Checklist.Count > 0)
                                                {
                                                    var rowList = new List<List<AddOnGs.Models.OpenAPI.Data>>();
                                                    for (int i = 0; i < Audit_Checklist.Count; i++)
                                                    {
                                                        dynamic items = Audit_Checklist[i];
                                                        string Nonconformity = items[3];
                                                        string Clause = items[0];
                                                        string NameClause = items[1];
                                                        string results = items[4];
                                                        if (!string.IsNullOrEmpty(results))
                                                        {
                                                            if (results == "NC")
                                                            {
                                                                results = "CAR";
                                                            }
                                                            else if (results == "OB.")
                                                            {
                                                                results = "OB";
                                                            }
                                                        }
                                                        if (results == "C")
                                                        {
                                                            continue;
                                                        }
                                                        var row = new List<AddOnGs.Models.OpenAPI.Data>
                                                            {
                                                                new AddOnGs.Models.OpenAPI.Data { Value = Nonconformity },
                                                                new AddOnGs.Models.OpenAPI.Data { Value = Clause },
                                                                new AddOnGs.Models.OpenAPI.Data { Value = NameClause },
                                                                new AddOnGs.Models.OpenAPI.Data { Value = results },
                                                            };
                                                        rowList.Add(row);

                                                    }
                                                    item.Data = new AddOnGs.Models.OpenAPI.Data { Row = rowList };
                                                }
                                                break;
                                            case "พื้นที่ตรวจ":
                                                item.Data = new Models.OpenAPI.Data { Value = Area };
                                                break;
                                            case "รอบที่":
                                                item.Data = new Models.OpenAPI.Data { Value = RoundNo };
                                                break;
                                            case "ฝ่าย":
                                                item.Data = new Models.OpenAPI.Data { Value = Department };
                                                break;
                                        }
                                    }
                                    var jsonString = JsonConvert.SerializeObject(template.form);
                                    var json = JsonConvert.DeserializeObject(jsonString);
                                    foreach (var emp in emps)
                                    {
                                        var modelCreateMemo = new CreateMemorandumRequest
                                        {
                                            Subject = $"Audit Report ISO 9001/IATF16949 : พื้นที่ตรวจ {Area} รอบที่ {RoundNo}",
                                            Form = template.form,
                                            UserPrincipalName = emp.Email,
                                            DocumentCode = documentCode,
                                            Action = "draft",
                                        };
                                        var response = await _client.PostAPI($"api/v1/memorandum", modelCreateMemo);
                                        if (response.IsSuccessStatusCode)
                                        {
                                            var data = await response.ReadContentAs<ExtCreateMemorandumResult>();
                                            Ext.WriteLogFile("Generate F-QMS-010 Success !");
                                        }
                                    }
                                }
                                else
                                {
                                    Ext.WriteLogFile("Someting Wrong when call Get Template");
                                    Ext.WriteLogFile("----------------------------------------------------------------------------");
                                    return false;
                                }
                            }
                            #endregion
                        }
                        return true;
                    }
                    else
                    {
                        // don't process
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Ext.ErrorLog(ex, "AutoGenerate010");
                return false;
            }
        }
        public async Task<bool> AutoGenerate004(AddonFormModel.Form_Model model)
        {
            try
            {
                using (var context = DBContext.OpenConnection(model.memoPage.memoDetail.connectionString))
                {
                    var memoMain = await context.TRNMemoes.Where(e => e.MemoId == model.memoPage.memoDetail.memoid).FirstOrDefaultAsync();
                    if (memoMain.StatusName == "Completed")
                    {

                        #region F-QMS-010
                        Ext.WriteLogFile("Start GetData F-QMS-010" + memoMain.MemoId);
                        string Audit_Date = string.Empty;
                        string Audit_Time = string.Empty;
                        string Part = string.Empty;
                        string Audit_Requirement = string.Empty;
                        string RoundNo = string.Empty;
                        string Area = string.Empty;
                        string Department = string.Empty;
                        string AreaCode = string.Empty;
                        List<object> List_Audit_Team = new List<object>();
                        List<object> Audit_Checklist = new List<object>();
                        JObject jsonAdvanceForm = JsonUtils.createJsonObject(memoMain.MAdvancveForm);
                        JArray itemsArray = (JArray)jsonAdvanceForm["items"];
                        foreach (JObject jItems in itemsArray)
                        {

                            JArray jLayoutArray = (JArray)jItems["layout"];
                            if (jLayoutArray.Count >= 1)
                            {
                                JObject jTemplateL = (JObject)jLayoutArray[0]["template"];
                                JObject jData = (JObject)jLayoutArray[0]["data"];
                                if ((String)jTemplateL["label"] == "วันที่ตรวจ")
                                {
                                    Audit_Date = jData["value"].ToString();
                                }
                                if ((String)jTemplateL["label"] == "เวลาที่ตรวจ")
                                {
                                    Audit_Time = jData["value"].ToString();
                                }
                                if ((String)jTemplateL["label"] == "ส่วน")
                                {
                                    Part = jData["value"].ToString();
                                }
                                if ((String)jTemplateL["label"] == "Audit Team")
                                {

                                    foreach (JArray row in jData["row"])
                                    {
                                        List<object> rowObject = new List<object>();
                                        foreach (JObject item in row)
                                        {
                                            if (item.TryGetValue("value", out JToken valueToken) && valueToken != null)
                                            {
                                                rowObject.Add(valueToken.ToString());
                                            }
                                        }
                                        if (rowObject.Count > 0)
                                        {
                                            List_Audit_Team.Add(rowObject);
                                        }
                                    }
                                }
                                if ((String)jTemplateL["label"] == "ข้อกำหนดที่ตรวจ")
                                {
                                    Audit_Requirement = jData["value"].ToString();
                                }
                                if ((String)jTemplateL["label"] == "ตารางผลการตรวจ")
                                {

                                    foreach (JArray row in jData["row"])
                                    {
                                        List<object> rowObject = new List<object>();
                                        foreach (JObject item in row)
                                        {
                                            if (item.TryGetValue("value", out JToken valueToken) && valueToken != null)
                                            {
                                                rowObject.Add(valueToken.ToString());
                                            }
                                        }
                                        if (rowObject.Count > 0)
                                        {
                                            Audit_Checklist.Add(rowObject);
                                        }
                                    }
                                }
                                if (jLayoutArray.Count > 1)
                                {
                                    JObject jTemplateR = (JObject)jLayoutArray[1]["template"];
                                    JObject jData2 = (JObject)jLayoutArray[1]["data"];
                                    if ((String)jTemplateR["label"] == "พื้นที่ตรวจ")
                                    {
                                        Area = jData2["value"].ToString();
                                    }
                                    if ((String)jTemplateR["label"] == "AreaCode")
                                    {
                                        AreaCode = jData2["value"].ToString();
                                    }
                                    if ((String)jTemplateR["label"] == "รอบที่")
                                    {
                                        RoundNo = jData2["value"].ToString();
                                    }
                                    if ((String)jTemplateR["label"] == "ฝ่าย")
                                    {
                                        Department = jData2["value"].ToString();
                                    }
                                }
                            }
                        }
                        Ext.WriteLogFile("End GetData F-QMS-010" + memoMain.MemoId);
                        #endregion
                        #region F-QMS-004
                        var emps = context.MSTEmployees.Where(e => memoMain.RNameEn.Contains(e.NameEn)).ToList();
                        if (emps.Any())
                        {
                            var documentCode = "F-QMS-004";
                            var templateid = context.MSTTemplates.Where(x => x.DocumentCode == documentCode).FirstOrDefault();
                            var getTemplate = await _client.GetAPI($"api/v1/template?documentCode={documentCode}");
                            if (getTemplate.IsSuccessStatusCode)
                            {
                                Ext.WriteLogFile("Start AutoGenerate F-QMS-004");
                                if (Audit_Checklist.Count > 0)
                                {
                                    for (int c = 0; c < Audit_Checklist.Count; c++)
                                    {
                                        dynamic items = Audit_Checklist[c];
                                        string results = items[3];
                                        string Clause = items[1];
                                        string NameClause = items[2];
                                        string DETAILS = items[0];
                                        string NONCONFORMITY_CLAUSE_NO = $"{Clause} : {NameClause}";
                                        if (!String.IsNullOrEmpty(results) && results == "CAR")
                                        {
                                            #region running
                                            //string AUDIT_AREA = string.Empty;
                                            //string pattern = @"\(([^)]+)\)";
                                            //Match match = Regex.Match(Area, pattern);
                                            //if (match.Success)
                                            //{
                                            //    AUDIT_AREA = match.Groups[1].Value;
                                            //}

                                            //string DOC_NO = string.Empty;
                                            //string Prefix = $"IQA-{RoundNo}-{AUDIT_AREA}-";
                                            string DOC_NO = string.Empty;
                                            string Prefix = $"IQA-{RoundNo}-{AreaCode}-";
                                            DOC_NO = InsertControlRunning_FQMS004(context, Prefix, templateid);
                                            #endregion
                                            var template = await getTemplate.ReadContentAs<ExtGetTemplateResponse>();
                                            var form = template.form;
                                            foreach (var item in template.form)
                                            {
                                                switch (item.Template.Label)
                                                {
                                                    case "วันที่ตรวจติดตาม":
                                                        item.Data = new Models.OpenAPI.Data { Value = Audit_Date };
                                                        break;
                                                    case "เวลาที่ตรวจติดตาม":
                                                        item.Data = new Models.OpenAPI.Data { Value = Audit_Time };
                                                        break;
                                                    case "ส่วน":
                                                        item.Data = new Models.OpenAPI.Data { Value = Part };
                                                        break;
                                                    case "หมายเลขลำดับเอกสาร":
                                                        item.Data = new Models.OpenAPI.Data { Value = DOC_NO };
                                                        break;
                                                    case "สิ่งที่ไม่เป็นไปตามข้อกำหนดที่":
                                                        item.Data = new Models.OpenAPI.Data { Value = NONCONFORMITY_CLAUSE_NO };
                                                        break;
                                                    case "ได้แก่":
                                                        item.Data = new Models.OpenAPI.Data { Value = DETAILS };
                                                        break;
                                                    case "ทีมผู้ตรวจติดตาม":
                                                        if (List_Audit_Team.Count > 0)
                                                        {
                                                            var rowList = new List<List<AddOnGs.Models.OpenAPI.Data>>();
                                                            for (int i = 0; i < List_Audit_Team.Count; i++)
                                                            {
                                                                dynamic itema = List_Audit_Team[i];
                                                                string Name = itema[0];
                                                                string Position = itema[1];
                                                                var row = new List<AddOnGs.Models.OpenAPI.Data>
                                                                {
                                                                    new AddOnGs.Models.OpenAPI.Data { Value = Name },
                                                                    new AddOnGs.Models.OpenAPI.Data { Value = Position }
                                                                };
                                                                rowList.Add(row);
                                                            }
                                                            item.Data = new AddOnGs.Models.OpenAPI.Data { Row = rowList };
                                                        }
                                                        break;
                                                    case "พื้นที่ตรวจติดตาม":
                                                        item.Data = new Models.OpenAPI.Data { Value = Area };
                                                        break;
                                                    case "การตรวจติดตาม":
                                                        item.Data = new Models.OpenAPI.Data { Value = "INTERNAL QUALITY AUDIT (IQA)" };
                                                        break;
                                                    case "รอบที่":
                                                        item.Data = new Models.OpenAPI.Data { Value = RoundNo };
                                                        break;
                                                    case "ฝ่าย":
                                                        item.Data = new Models.OpenAPI.Data { Value = Department };
                                                        break;
                                                    case "AreaCode":
                                                        item.Data = new Models.OpenAPI.Data { Value = AreaCode };
                                                        break;
                                                }
                                            }
                                            var jsonString = JsonConvert.SerializeObject(template.form);
                                            var json = JsonConvert.DeserializeObject(jsonString);
                                            foreach (var emp in emps)
                                            {
                                                var modelCreateMemo = new CreateMemorandumRequest
                                                {
                                                    Subject = $"CORRECTIVE ACTION REQUEST ISO 9001/IATF16949 : พื้นที่ตรวจ {Area} รอบที่ {RoundNo}",
                                                    Form = template.form,
                                                    UserPrincipalName = emp.Email, 
                                                    DocumentCode = documentCode,
                                                    Action = "draft"//,
                                                    //EFixCreator = "varintorn.a@gsbattery.co.th",
                                                    //EFixRequestor = emp.Email

                                                };
                                                var response = await _client.PostAPI($"api/v1/memorandum", modelCreateMemo);
                                                if (response.IsSuccessStatusCode)
                                                {
                                                    var data = await response.ReadContentAs<ExtCreateMemorandumResult>();
                                                    Ext.WriteLogFile("Generate F-QMS-004 Success !");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                Ext.WriteLogFile("Someting Wrong when call Get Template");
                                Ext.WriteLogFile("----------------------------------------------------------------------------");
                                return false;
                            }
                        }
                        return true;
                        #endregion
                    }
                    else
                    {
                        // don't process
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Ext.ErrorLog(ex, "AutoGenerate004");
                return false;
            }
        }
        public static string InsertControlRunning_FQMS004(WolfApproveModel db, string Prefix, MSTTemplate FQMS004)
        {
            try
            {
                var checkPrefix = db.TRNControlRunning.Where(x => x.Prefix == Prefix).OrderBy(r => r.Running).ToList();
                if (checkPrefix != null)
                {
                    TRNControlRunning objControlRunning = new TRNControlRunning();
                    objControlRunning.TemplateId = FQMS004.TemplateId ?? 0;
                    objControlRunning.Prefix = Prefix;
                    objControlRunning.Digit = 4;
                    objControlRunning.Running = (checkPrefix.LastOrDefault()?.Running ?? 0) + 1;
                    objControlRunning.CreateBy = "1";
                    objControlRunning.CreateDate = DateTime.Now;
                    objControlRunning.RunningNumber = $"{Prefix}{objControlRunning.Running.ToString().PadLeft(4, '0')}";
                    db.TRNControlRunning.Add(objControlRunning);
                    db.SaveChanges();
                    string Running = objControlRunning.RunningNumber;
                    return Running;
                }
                else
                {
                    TRNControlRunning objControlRunning = new TRNControlRunning();
                    objControlRunning.TemplateId = FQMS004.TemplateId ?? 0;
                    objControlRunning.Prefix = Prefix;
                    objControlRunning.Digit = 4;
                    objControlRunning.Running = 1;
                    objControlRunning.CreateBy = "1";
                    objControlRunning.CreateDate = DateTime.Now;
                    objControlRunning.RunningNumber = $"{Prefix}0001";
                    db.TRNControlRunning.Add(objControlRunning);
                    db.SaveChanges();
                    string Running = objControlRunning.RunningNumber;
                    return Running;
                }
                return string.Empty;
            }
            catch (Exception ex)
            {
                Ext.ErrorLog(ex, "InsertControlRunning_FQMS004");
                return string.Empty;
            }

        }
        public async Task<bool> AutoGenerate022(AddonFormModel.Form_Model model)
        {
            try
            {
                using (var context = DBContext.OpenConnection(model.memoPage.memoDetail.connectionString))
                {
                    var memoMain = await context.TRNMemoes.Where(e => e.MemoId == model.memoPage.memoDetail.memoid).FirstOrDefaultAsync();
                    if (memoMain.StatusName == "Wait for Approve")
                    {
                        Ext.WriteLogFile("Start AutoGenerate SE-F-SHE-022");
                        var empseq1 = context.TRNLineApproves.Where(x => x.MemoId == memoMain.MemoId && x.Seq == 1).Select(x => x.EmployeeId);
                        if (empseq1.Any(e => e.HasValue && e.Value == memoMain.PersonWaitingId))
                        {
                            var advancefourm = context.TRNMemoForms.Where(e => e.MemoId == model.memoPage.memoDetail.memoid).ToList();
                            var tbrowcount = advancefourm.Where(e => e.obj_type == "tb").FirstOrDefault();
                            var tbData = advancefourm.Where(e => e.obj_type == "tb").ToList();
                            var Alldata = advancefourm.Where(e => e.obj_type != "tb").ToList();

                            CustomTRNFormModel allList = new CustomTRNFormModel();

                            List<string> newvaluetb = new List<string>();
                            foreach (var getdata in Alldata)
                            {
                                if (getdata.obj_label == "ครั้งที่ Audit")
                                {
                                    allList.Audittime = getdata.obj_value;
                                }
                                if (getdata.obj_label == "ปี Audit")
                                {
                                    allList.YearAudit = getdata.obj_value;
                                }
                                if (getdata.obj_label == "วันที่ตรวจ")
                                {
                                    allList.AuditDate = getdata.obj_value;
                                }
                                if (getdata.obj_label == "เวลาที่ตรวจ")
                                {
                                    allList.Inspectiontime = getdata.obj_value;
                                }
                                if (getdata.obj_label == "พื้นที่ตรวจ")
                                {
                                    allList.AuditArea = getdata.obj_value;
                                }

                            }




                            var emps = context.MSTEmployees.Where(e => memoMain.RNameEn.Contains(e.NameEn)).ToList();

                            if (emps.Any())
                            {
                                var documentCode = "SE-F-SHE-022";
                                // Fix Url บนเครื่องตัวเองไปก่อน 
                                var getTemplate = await _client.GetAPI($"api/v1/template?documentCode={documentCode}");
                                if (getTemplate.IsSuccessStatusCode)
                                {
                                    var template = await getTemplate.ReadContentAs<ExtGetTemplateResponse>();
                                    var form = template.form;
                                    foreach (var item in template.form)
                                    {
                                        if (item.Template.Label == "ครั้งที่ Audit") item.Data = new Models.OpenAPI.Data { Value = allList.Audittime };
                                        if (item.Template.Label == "ปี Audit") item.Data = new Models.OpenAPI.Data { Value = allList.YearAudit };
                                        if (item.Template.Label == "วันที่ตรวจ") item.Data = new Models.OpenAPI.Data { Value = allList.AuditDate };
                                        if (item.Template.Label == "เวลาที่ตรวจ") item.Data = new Models.OpenAPI.Data { Value = allList.Inspectiontime };
                                        if (item.Template.Label == "พื้นที่ตรวจ") item.Data = new Models.OpenAPI.Data { Value = allList.AuditArea };
                                        if (item.Template.Label == "Audit Team")
                                        {
                                            var getdatatb = tbData.Where(x => x.col_label == "LEAD AUDITOR" && x.obj_label == "Audit Team").ToList();

                                            // สร้าง JSON string
                                            string tablevalue = $"[{string.Join(",", getdatatb.Select(x => x.row_value))}]";

                                            // แปลง JSON string เป็น List<List<Models.OpenAPI.Data>>
                                            var dataList = JsonConvert.DeserializeObject<List<List<Models.OpenAPI.Data>>>(tablevalue);

                                            // กำหนดค่าให้ item.Data
                                            item.Data = new Models.OpenAPI.Data
                                            {
                                                Row = dataList
                                            };
                                        }
                                        if (item.Template.Label == "ข้อกำหนดที่ตรวจ")
                                        {
                                            string Check = string.Join(",", context.TRNMemoForms.Where(e => e.MemoId == model.memoPage.memoDetail.memoid && e.obj_label == "Audit Checklist" && e.col_label == "ข้อกำหนด").Select(e => e.col_value).ToList());
                                            item.Data = new Models.OpenAPI.Data { Value = Check };
                                        }
                                        if (item.Template.Label == "ตารางผลการตรวจ")
                                        {
                                            var getdatatb = tbData.Where(x => x.col_label == "ผลการตัดสิน" && (x.col_value == "OFI" || x.col_value == "NC")).ToList();
                                            var orderMapping = new Dictionary<string, int>
                                        {
                                            { "ข้อบกพร่องที่พบ /ข้อเสนอและปรับปรุง", 0 },
                                            { "ข้อกำหนด", 1 },
                                            { "ชื่อข้อกำหนด", 2 },
                                            { "ผลตรวจ", 3 }
                                        };


                                            foreach (var checkdata in getdatatb)
                                            {
                                                var resultList = new List<Dictionary<string, string>>();
                                                if (checkdata.col_value == "OFI" || checkdata.col_value == "NC")
                                                {
                                                    var getrowindex = context.TRNMemoForms
                                                                            .Where(x => x.MemoId == model.memoPage.memoDetail.memoid
                                                                                        && x.row_index == checkdata.row_index && x.obj_label == "Audit Checklist")
                                                                            .ToList();

                                                    foreach (var addvalue in getrowindex)
                                                    {
                                                        var rowData = new Dictionary<string, string>();

                                                        if (addvalue.col_label == "หลักฐานการตรวจ")
                                                        {
                                                            if (item.Template.Attribute.Column[0].Template.Label == "ข้อบกพร่องที่พบ /ข้อเสนอและปรับปรุง")
                                                            {
                                                                rowData["label"] = "ข้อบกพร่องที่พบ /ข้อเสนอและปรับปรุง";
                                                                rowData["value"] = addvalue.col_value;
                                                            }
                                                        }
                                                        else if (addvalue.col_label == "ข้อกำหนด")
                                                        {
                                                            if (item.Template.Attribute.Column[1].Template.Label == addvalue.col_label)
                                                            {
                                                                rowData["label"] = "ข้อกำหนด";
                                                                rowData["value"] = addvalue.col_value;
                                                            }
                                                        }
                                                        else if (addvalue.col_label == "ชื่อข้อกำหนด")
                                                        {
                                                            if (item.Template.Attribute.Column[2].Template.Label == "ชื่อข้อกำหนด")
                                                            {
                                                                rowData["label"] = "ชื่อข้อกำหนด";
                                                                rowData["value"] = addvalue.col_value;
                                                            }
                                                        }
                                                        else if (addvalue.col_label == "ผลการตัดสิน")
                                                        {
                                                            if (item.Template.Attribute.Column[3].Template.Label == "ผลตรวจ")
                                                            {
                                                                rowData["label"] = "ผลตรวจ";
                                                                rowData["value"] = checkdata.col_value == "OFI" ? "OFI" : "CAR";
                                                            }
                                                        }

                                                        if (rowData.Count > 0)
                                                        {
                                                            resultList.Add(rowData);
                                                        }
                                                    }
                                                }
                                                resultList = resultList
                                                    .OrderBy(x => orderMapping.ContainsKey(x["label"]) ? orderMapping[x["label"]] : int.MaxValue)
                                                    .ToList();
                                                var finalResultList = resultList
                                                    .Select(x => new Dictionary<string, string> { { "value", x["value"] } })
                                                    .ToList();
                                                StringBuilder jsonBuilder = new StringBuilder();
                                                jsonBuilder.Append("[");

                                                foreach (var items in finalResultList)
                                                {
                                                    jsonBuilder.AppendFormat("{{\"value\":\"{0}\"}},", items["value"]);
                                                }
                                                if (jsonBuilder[jsonBuilder.Length - 1] == ',')
                                                {
                                                    jsonBuilder.Length--;
                                                }
                                                jsonBuilder.Append("]");

                                                string jsonResult = jsonBuilder.ToString();
                                                newvaluetb.Add(jsonResult);
                                            }


                                            string resultJson = string.Join(",", newvaluetb);
                                            string tablevalue = $"[{resultJson}]";
                                            var dataList = JsonConvert.DeserializeObject<List<List<Models.OpenAPI.Data>>>(tablevalue);

                                            item.Data = new Models.OpenAPI.Data
                                            {
                                                Row = dataList
                                            };
                                        }
                                    }
                                    var jsonString = JsonConvert.SerializeObject(template.form);
                                    var json = JsonConvert.DeserializeObject(jsonString);
                                    foreach (var emp in emps)
                                    {
                                        // Mapping Data to MAdvance Forum
                                        var modelCreateMemo = new CreateMemorandumRequest
                                        {
                                            Subject = $"Audit Report ISO 14001/ISO 45001 : พื่นที่ตรวจ " + allList.AUDITEES + "ครั้งที่ Audit " + allList.Audittime + " ปี Audit " + allList.YearAudit + "",
                                            Form = template.form,
                                            UserPrincipalName = emp.Email,
                                            DocumentCode = documentCode,
                                            Action = "draft",
                                        };
                                        var response = await _client.PostAPI($"api/v1/memorandum", modelCreateMemo);
                                        if (response.IsSuccessStatusCode)
                                        {
                                            var data = await response.ReadContentAs<ExtCreateMemorandumResult>();
                                            Ext.WriteLogFile("AutoGenerate SE-F-SHE-021 Success !");
                                        }
                                    }


                                }
                                else
                                {
                                    Ext.WriteLogFile("Someting Wrong when call Get Template");
                                    return false;
                                }
                            }
                        }
                        return true;
                    }
                    else
                    {
                        // don't process
                        return true;
                    }
                }

            }
            catch (Exception ex)
            {
                Ext.ErrorLog(ex, "AutoGenerate021");
                return false;
            }
        }
        public async Task<bool> AutoGenerateCar(AddonFormModel.Form_Model model)
        {
            try
            {
                using (var context = DBContext.OpenConnection(model.memoPage.memoDetail.connectionString))
                {
                    var memoMain = await context.TRNMemoes.Where(e => e.MemoId == model.memoPage.memoDetail.memoid).FirstOrDefaultAsync();
                    if (memoMain.StatusName == "Completed")
                    {
                        Ext.WriteLogFile("Start AutoGenerate SE-F-SHE-024");

                        var advancefourm = context.TRNMemoForms.Where(e => e.MemoId == model.memoPage.memoDetail.memoid).ToList();
                        var tbrowcount = advancefourm.Where(e => e.obj_type == "tb").FirstOrDefault();
                        var tbData = advancefourm.Where(e => e.obj_type == "tb").ToList();
                        var Alldata = advancefourm.Where(e => e.obj_type != "tb").ToList();
                        var Cardata = advancefourm.Where(e => e.obj_type == "tb" && e.col_value == "CAR").ToList();
                        CustomTRNFormModel allList = new CustomTRNFormModel();

                        List<string> newvaluetb = new List<string>();
                        foreach (var getdata in Alldata)
                        {
                            if (getdata.obj_label == "ครั้งที่ Audit")
                            {
                                allList.Audittime = getdata.obj_value;
                            }
                            if (getdata.obj_label == "ปี Audit")
                            {
                                allList.YearAudit = getdata.obj_value;
                            }
                            if (getdata.obj_label == "วันที่ตรวจ")
                            {
                                allList.AuditDate = getdata.obj_value;
                            }
                            if (getdata.obj_label == "เวลาที่ตรวจ")
                            {
                                allList.Inspectiontime = getdata.obj_value;
                            }
                            if (getdata.obj_label == "พื้นที่ตรวจ")
                            {
                                allList.AuditArea = getdata.obj_value;
                            }

                        }
                        foreach (var oprnCar in Cardata)
                        {
                            var documentCode = "SE-F-SHE-024";
                            var templateid = context.MSTTemplates.Where(x => x.DocumentCode == documentCode).FirstOrDefault();
                            string DOC_NO = string.Empty;
                            string Prefix = $"{allList.AuditArea}-{allList.Audittime}-{allList.YearAudit}-";
                            DOC_NO = InsertControlRunning_Car(context, Prefix, templateid);
                            var emps = context.MSTEmployees.Where(e => memoMain.RNameEn.Contains(e.NameEn)).ToList();
                            allList.CarNo = DOC_NO;
                            if (emps.Any())
                            {
                                var getTemplate = await _client.GetAPI($"api/v1/template?documentCode={documentCode}");
                                if (getTemplate.IsSuccessStatusCode)
                                {
                                    var template = await getTemplate.ReadContentAs<ExtGetTemplateResponse>();
                                    var form = template.form;
                                    foreach (var item in template.form)
                                    {
                                        if (item.Template.Label == "ครั้งที่ Audit") item.Data = new Models.OpenAPI.Data { Value = allList.Audittime };
                                        if (item.Template.Label == "ปี Audit") item.Data = new Models.OpenAPI.Data { Value = allList.YearAudit };
                                        if (item.Template.Label == "วันที่ตรวจ") item.Data = new Models.OpenAPI.Data { Value = allList.AuditDate };
                                        if (item.Template.Label == "วันที่ตรวจติดตาม") item.Data = new Models.OpenAPI.Data { Value = allList.Inspectiontime };
                                        if (item.Template.Label == "พื้นที่ตรวจ") item.Data = new Models.OpenAPI.Data { Value = allList.AuditArea };
                                        if (item.Template.Label == "หมายเลขลำดับเอกสาร") item.Data = new Models.OpenAPI.Data { Value = allList.CarNo };
                                        if (item.Template.Label == "การตรวจติดตาม(AUDIT)") item.Data = new Models.OpenAPI.Data { Value = "การตรวจติดตามภายใน" };
                                        var Cardatavalue = advancefourm.Where(x => x.obj_type == "tb" && x.obj_label == "ตารางผลการตรวจ" && x.col_label == "ข้อบกพร่องที่พบ /ข้อเสนอและปรับปรุง" && x.row_index == oprnCar.row_index).FirstOrDefault();
                                        if (item.Template.Label == "ได้แก่") item.Data = new Models.OpenAPI.Data { Value = Cardatavalue.col_value };
                                        if (item.Template.Label == "สิ่งที่ไม่เป็นไปตามข้อกำหนดที่")
                                        {
                                            string withoutrequirements = string.Empty;
                                            var Datatb1 = advancefourm.Where(e => e.obj_type == "tb" && e.obj_label == "ตารางผลการตรวจ" && e.row_index == oprnCar.row_index && e.col_label == "ข้อกำหนด").FirstOrDefault();
                                            var Datatb2 = advancefourm.Where(e => e.obj_type == "tb" && e.obj_label == "ตารางผลการตรวจ" && e.row_index == oprnCar.row_index && e.col_label == "ชื่อข้อกำหนด").FirstOrDefault();
                                            withoutrequirements += $"{Datatb1.col_value}: {Datatb2.col_value}";
                                            item.Data = new Models.OpenAPI.Data { Value = withoutrequirements };
                                        }

                                        if (item.Template.Label == "ทีมผู้ตรวจติดตาม")
                                        {
                                            var getdatatb = tbData.Where(x => x.col_label == "LEAD AUDITOR" && x.obj_label == "Audit Team").ToList();

                                            // สร้าง JSON string
                                            string tablevalue = $"[{string.Join(",", getdatatb.Select(x => x.row_value))}]";

                                            // แปลง JSON string เป็น List<List<Models.OpenAPI.Data>>
                                            var dataList = JsonConvert.DeserializeObject<List<List<Models.OpenAPI.Data>>>(tablevalue);

                                            // กำหนดค่าให้ item.Data
                                            item.Data = new Models.OpenAPI.Data
                                            {
                                                Row = dataList
                                            };
                                        }
                                    }
                                    var jsonString = JsonConvert.SerializeObject(template.form);
                                    var json = JsonConvert.DeserializeObject(jsonString);
                                    foreach (var emp in emps)
                                    {
                                        // Mapping Data to MAdvance Forum
                                        var modelCreateMemo = new CreateMemorandumRequest
                                        {
                                            Subject = $"CORRECTIVE ACTION REQUEST ISO 14001/ISO 45001 : พื่นที่ตรวจ " + allList.AUDITEES + "ครั้งที่ Audit " + allList.Audittime + " ปี Audit " + allList.YearAudit + "",
                                            Form = template.form,
                                            UserPrincipalName = emp.Email,
                                            DocumentCode = documentCode,
                                            Action = "draft",
                                        };
                                        var response = await _client.PostAPI($"api/v1/memorandum", modelCreateMemo);
                                        if (response.IsSuccessStatusCode)
                                        {
                                            var data = await response.ReadContentAs<ExtCreateMemorandumResult>();
                                            Ext.WriteLogFile("AutoGenerate SE-F-SHE-024 Success !");
                                        }
                                    }
                                }
                                else
                                {
                                    Ext.WriteLogFile("Someting Wrong when call Get Template");
                                    return false;
                                }
                            }
                        }
                        

                        return true;

                    }
                    else
                    {
                        // don't process
                        return true;
                    }
                }

            }
            catch (Exception ex)
            {
                Ext.ErrorLog(ex, "SE-F-SHE-024");
                return false;
            }
        }
        public static string InsertControlRunning_Car(WolfApproveModel db, string Prefix, MSTTemplate FQMS004)
        {
            try
            {
                var checkPrefix = db.TRNControlRunning.Where(x => x.Prefix == Prefix).OrderBy(r => r.Running).ToList();
                if (checkPrefix != null)
                {
                    TRNControlRunning objControlRunning = new TRNControlRunning();
                    objControlRunning.TemplateId = FQMS004.TemplateId ?? 0;
                    objControlRunning.Prefix = Prefix;
                    objControlRunning.Digit = 3;
                    objControlRunning.Running = (checkPrefix.LastOrDefault()?.Running ?? 0) + 1;
                    objControlRunning.CreateBy = "1";
                    objControlRunning.CreateDate = DateTime.Now;
                    objControlRunning.RunningNumber = $"{Prefix}{objControlRunning.Running.ToString().PadLeft(3, '0')}";
                    db.TRNControlRunning.Add(objControlRunning);
                    db.SaveChanges();
                    string Running = objControlRunning.RunningNumber;
                    return Running;
                }
                else
                {
                    TRNControlRunning objControlRunning = new TRNControlRunning();
                    objControlRunning.TemplateId = FQMS004.TemplateId ?? 0;
                    objControlRunning.Prefix = Prefix;
                    objControlRunning.Digit = 3;
                    objControlRunning.Running = 1;
                    objControlRunning.CreateBy = "1";
                    objControlRunning.CreateDate = DateTime.Now;
                    objControlRunning.RunningNumber = $"{Prefix}0001";
                    db.TRNControlRunning.Add(objControlRunning);
                    db.SaveChanges();
                    string Running = objControlRunning.RunningNumber;
                    return Running;
                }
                return string.Empty;
            }
            catch (Exception ex)
            {
                Ext.ErrorLog(ex, "InsertControlRunning_SHE0244");
                return string.Empty;
            }

        }
    }
}
