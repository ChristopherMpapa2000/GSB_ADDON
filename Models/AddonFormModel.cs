using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WolfApprove.Model.CustomClass;
using WolfApprove.Model;


namespace AddOnGs.Models
{
    public class AddonFormModel
    {
        [Serializable]
        public class Form_Model : CustomClass
        {
            //public WorklistPage worklistPage { get; set; }
            public Form_Model()
            {
                this.delegatePage = new DelegatePage();
                this.MemoAutoNumber = new MemoAutoNumber();
            }

            public MemoPage memoPage { get; set; }
            public string Type { get; set; }
            public string sAction { get; set; }
            public bool IsPreview { get; set; }
            public string UserPrincipalName { get; set; }
            public string ConnectionString { get; set; }

            public DelegatePage delegatePage { get; set; }

            public MemoAutoNumber MemoAutoNumber { get; set; }
            //public Boolean IsPreview { get; set; }
        }

        #region | Worklist Page |
        [Serializable]
        public class WorklistPage
        {
            public WorklistPage(string task)
            {
                var taskgroup = new Taskgroup();
                selectedTaskgroup = task;
            }
            public List<CustomWorklist> listWorklistDetail { get; set; }
            public string selectedTaskgroup { get; set; }
        }
        [Serializable]
        public class Taskgroup
        {
            public KeyValuePair<string, string> _todo = new KeyValuePair<string, string>("To Do List", "todo");
            public KeyValuePair<string, string> _inprocess = new KeyValuePair<string, string>("Inprocess", "inprocess");
            public KeyValuePair<string, string> _completed = new KeyValuePair<string, string>("Completed", "completed");
            //public SelectList taskgroup { get; set; }
            public Taskgroup()
            {
                var list = new List<KeyValuePair<string, string>>();
                list.Add(_todo);
                list.Add(_inprocess);
                list.Add(_completed);
            }
        }
        [Serializable]
        public class WorklistDetail : CustomClass
        {
            public ViewEmployee requestor { get; set; }

            public int? memoid { get; set; }
            public string waiting_for { get; set; }
            public string status { get; set; }
            public string document_no { get; set; }
            public string template_name { get; set; }
            public string request_date { get; set; }
            public string company_name { get; set; }
            public string subject { get; set; }
            public string amount { get; set; }
            public string modified_date { get; set; }
        }

        #endregion
        #region | Memo Page |
        [Serializable]
        public class MemoPage : CustomClass
        {
            public MemoPage()
            {
                //create new memo
                memoDetail = new MemoDetail();
                //memoDetail.actor = Extension.GetEmployeeDetail("T-0057");     
                memoDetail.creator = memoDetail.actor;
                memoDetail.requestor = memoDetail.actor;
                listApprovalDetails = new List<ApprovalDetail>();
                listFileAttachDetails = new List<FileAttachDetail>();
                listHistoryDetails = new List<HistoryDetail>();
                listFormName = new List<CustomTemplate>();
                listCompany = new List<CustomCompany>();
                listProject = new List<CustomProject>();
                listSignatureWording = new List<CustomMasterData>();
                // new code get history view master data
                HistoryView = new CustomMasterData();
                this.listLogic = new List<MSTTemplateLogic>();
                this.listUserPermission = new List<CustomUserPermission>();
                listRefDocDetails = new List<RefDocDetail>();
                listControlRunning = new List<CustomControlRunning>();


            }
            public MemoDetail memoDetail { get; set; }
            //public List<ApprovalDetail> listApprovalDetails { get; set; }
            public List<RefDocDetail> listRefDocDetails { get; set; }
            public List<ApprovalDetail> listApprovalDetails { get; set; }
            public List<HistoryDetail> listHistoryDetails { get; set; }
            public List<FileAttachDetail> listFileAttachDetails { get; set; }
            public List<ActionButtonDetail> listActionButtonDetails { get; set; }
            public List<CustomTemplate> listFormName { get; set; }
            public List<CustomCompany> listCompany { get; set; }
            public List<CustomProject> listProject { get; set; }
            public List<CustomMasterData> listSignatureWording { get; set; }
            // 06-04-2020
            // new code listlogic
            public List<MSTTemplateLogic> listLogic { get; set; }
            public List<ItemLogicloaddata> listLogicloaddata { get; set; }
            public List<CustomUserPermission> listUserPermission { get; set; }
            // new code get history view master data
            public CustomMasterData HistoryView { get; set; }

            public List<CustomControlRunning> listControlRunning { get; set; }
            public List<int> listDelegateToId { get; set; }
        }

        [Serializable]
        public class MemoAutoNumber
        {
            public int? TemplateId { get; set; }
            public string PreFix { get; set; }
            public int? Digit { get; set; }

            public string RunningNumber { get; set; }
            public string Connection { get; set; }
        }

        [Serializable]
        public class FileAttachDetail : CustomClass
        {
            public CustomViewEmployee actor { get; set; }

            public int? sequence { get; set; }
            public int attach_id { get; set; }
            public int? memo_id { get; set; }
            public string attach_file { get; set; }
            public string description { get; set; }
            public string attach_path { get; set; }
            public string attach_date { get; set; }

            public int? delegate_id { get; set; }
            public string modified_date { get; set; }
            public string modified_by { get; set; }
            public Boolean is_merge_pdf { get; set; }

        }
        [Serializable]
        public class HistoryDetail : CustomClass
        {
            public CustomViewEmployee actor { get; set; }

            public int? action_id { get; set; }
            public int memo_id { get; set; }
            public string action { get; set; }
            public string status { get; set; }
            public string comment { get; set; }

            public string action_date { get; set; }
            public int? signature_id { get; set; }
            public string platform { get; set; }

            public string ip_address { get; set; }
            public string list_file_path { get; set; }

            public int? actor_id { get; set; }
            public string actor_name_th { get; set; }
            public string actor_name_en { get; set; }
            public int? actor_position_id { get; set; }
            public string actor_position_name_th { get; set; }
            public string actor_position_name_en { get; set; }
            public int? actor_department_id { get; set; }
            public string actor_department_name_th { get; set; }
            public string actor_department_name_en { get; set; }

            public int? delegate_actor_id { get; set; }
            public string delegate_actor_name_th { get; set; }
            public string delegate_actor_name_en { get; set; }
            public int? delegate_actor_position_id { get; set; }
            public string delegate_actor_position_name_th { get; set; }
            public string delegate_actor_position_name_en { get; set; }
            public int? delegate_actor_department_id { get; set; }
            public string delegate_actor_department_name_th { get; set; }
            public string delegate_actor_department_name_en { get; set; }
            public string HAdvancveForm { get; set; }

        }
        [Serializable]
        public class MemoDetail : CustomClass
        {
            public MemoDetail()
            {
                document_no = "Auto Generate";
                status = "New Request";
            }

            public CustomViewEmployee creator { get; set; }
            public CustomViewEmployee requestor { get; set; }
            public CustomViewEmployee actor { get; set; }


            public int? memoid { get; set; }
            public int? current_approval_level { get; set; }
            public string waiting_for { get; set; }
            public string status { get; set; }
            public string document_no { get; set; }
            public int template_id { get; set; }
            public string template_name { get; set; }
            public string request_date { get; set; }
            public int? company_id { get; set; }
            public string company_name { get; set; }
            public string location { get; set; }
            public string to { get; set; }
            public string pass { get; set; }
            public string subject { get; set; }
            public int? project_id { get; set; }
            public string project { get; set; }
            public string template_desc { get; set; }
            public string costcenter { get; set; }
            public string amount { get; set; }
            public string wbs { get; set; }
            public string io { get; set; }
            public string comment { get; set; }
            public string document_set { get; set; }
            public string document_library { get; set; }
            public bool is_editable { get; set; }
            public int? department_id { get; set; }
            public bool is_public { get; set; }
            public string report_lang { get; set; }
            public string template_detail { get; set; }
            public bool auto_approve { get; set; }
            public string auto_approve_when { get; set; }
            public bool approver_can_edit { get; set; }
            public int? status_id { get; set; }
            public string created_date { get; set; }
            public string created_by { get; set; }
            public string modified_date { get; set; }
            public string modified_by { get; set; }
            public string last_action_by { get; set; }
            public int? last_status_id { get; set; }
            public string last_status_name { get; set; }
            public int? waiting_for_id { get; set; }
            public bool is_text_form { get; set; }
            public string template_code { get; set; }
            public string GroupTemplateName { get; set; }
            public string refrenece_doc { get; set; }
            public string TemplateApproveId { get; set; }

            public JObject Permission { get; set; }
            public bool IsCheckAccess { get; set; }
            public CustomViewEmployee actorCheckAccess { get; set; }
            public DateTime ModifiedDate { get; set; }
            public bool isProcess { get; set; }
            public string codesap { get; set; }
            public string branchsap { get; set; }
            public string currency { get; set; }
            public string wolfurl { get; set; }
        }
        [Serializable]
        public class ApprovalDetail : CustomClass
        {

            public CustomViewEmployee approver { get; set; }

            public int? lineid { get; set; }
            public int? memoid { get; set; }
            public int? sequence { get; set; }
            public int? emp_id { get; set; }
            public int? signature_id { get; set; }
            public string signature_th { get; set; }
            public string signature_en { get; set; }

            //public bool? is_deletable { get; set; }
            //public bool is_sequencable { get; set; }
            //public bool? is_editable { get; set; }
            //public bool? is_breakposca { get; set; }
            //public bool? is_breakdeptca { get; set; }

            public string modifiedby { get; set; }
            public string modifieddate { get; set; }

            public int? TemLineId { get; set; }
            public int? ApproveType { get; set; }
            public int? LineApproveSeq { get; set; }
            public bool? IsParallel { get; set; }
            public bool? IsApproveAll { get; set; }
            public int? ApproveSlot { get; set; }
            public bool? UserAction { get; set; }
            public string ColumnJsonCondValue { get; set; }
            public bool? IsSpecific { get; set; }
            public bool? IsSpecificLOC { get; set; }
            public int? ApproveLOC_MaxLevelId { get; set; }
            public decimal PositionLevel { get; set; }
        }

        [Serializable()]
        public class RefDocDetail : CustomClass
        {
            public int? memoRefdoc_id { get; set; }
            public int? memoid { get; set; }
            public int? sequence { get; set; }
            public int? doc_id { get; set; }
            public string doc_no { get; set; }
            public string document_no { get; set; }
            public int? template_ID { get; set; }
            public string template_Name { get; set; }
            public string template_Subject { get; set; }
            public string memoSubject { get; set; }
            public string createdby { get; set; }
            //public DateTime? createddate { get; set; }
            public decimal? amount { get; set; }
        }
        [Serializable]
        public class ActionButtonDetail : CustomClass
        {
            public int? sequence { get; set; }
            public string cssClass { get; set; }
            public string text { get; set; }
            public string commandName { get; set; }
            public string faClass { get; set; }
            public string style { get; set; }
            public bool isDelegate { get; set; }
            public int? delegateFormEmpID { get; set; }
            public int? delegateToEmpID { get; set; }
        }


        public class AttachFile : CustomClass
        {
            public byte[] file { get; set; }
            public string file_name { get; set; }
            public string file_desc { get; set; }
            public string document_lib { get; set; }
            public string document_set { get; set; }
            public CustomViewEmployee actor { get; set; }
        }
        #endregion

        #region *** Delegate ***
        [Serializable]
        public class DelegatePage : CustomClass
        {
            public int delegateId { get; set; }
            public CustomDelegate header { get; set; }
            public List<CustomDelegateDetail> line { get; set; }
            public List<FileAttachDetail> attachment { get; set; }
            public CustomViewEmployee actor { get; set; }
        }
        #endregion

        [Serializable]
        public class RefDocColumn
        {
            public string Template { get; set; }
            public string Key { get; set; }
            public string Value { get; set; }
            public string TypeControl { get; set; }

            public List<RefDocColumn> objTable { get; set; }

        }

        [Serializable]
        public class MemoDDL
        {
            public string Key { get; set; }
            public string Value { get; set; }
        }
        [Serializable]
        public class DashBoardDetail
        {
            public ViewEmployee emp_detail { get; set; }

            public int? todo_count { get; set; }
            public int? inprogress_count { get; set; }
            public int? completed_count { get; set; }

        }
        public class HeaderItem
        {
            public string Item { get; set; }
            public IList<Data> Data { get; set; }
        }
        public class Data
        {
            public string KEY { get; set; }
            public string VALUE { get; set; }
            public string TYPE { get; set; }
            public string ItemCheck { get; set; }
        }
        public class SaveApprovals : CustomClass
        {
            public int? MemoId { get; set; }
            public int? current_approval_level { get; set; }
            public List<ApprovalDetail> ListApprovalDetails { get; set; }
        }
    }
}
