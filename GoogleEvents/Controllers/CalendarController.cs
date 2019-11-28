using DHTMLX.Common;
using DHTMLX.Scheduler;
using DHTMLX.Scheduler.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using GoogleEvents.Models;
using Google.Apis.Calendar.v3;
using Microsoft.AspNet.Identity;
using Google.Apis.Auth.OAuth2;
using System.IO;
using System.Threading;
using Google.Apis.Util.Store;
using Google.Apis.Services;
using Google.Apis.Calendar.v3.Data;
using System.Web.UI.WebControls;
using System.Web.UI;
using System.Collections;
using System.Data;
using OfficeOpenXml;
using OfficeOpenXml.Table;
namespace GoogleEvents.Controllers
{
    [Authorize]
    public class CalendarController : Controller
    {
        static string[] Scopes = { CalendarService.Scope.Calendar };
        string userName;

        public ActionResult Index()
        {
            this.userName = User.Identity.GetUserName();
            getCredentials();
            var sched = new DHXScheduler(this);
            sched.Skin = DHXScheduler.Skins.Material;
            sched.LoadData = true;
            sched.EnableDataprocessor = true;
            sched.InitialDate = new DateTime();
            
            // 'PDF' extension is required
            sched.Extensions.Add(SchedulerExtensions.Extension.PDF);
            return View(sched);
        }
        //public IList GetEventsFromdb()
        //{
        //    var user = User.Identity.GetUserName();
        //    var eventsList = new CalendarContext().Events.
        //        Where(e => e.user.Equals(user)).
        //    Select(e => new { e.id, e.text, e.start_date, e.end_date }).ToList();
        //    return eventsList;
        //}


          public ContentResult Data()
        {
            return (new SchedulerAjaxData(
                 GetEvents()
                .Select(e => new { e.gid, e.text, e.start_date, e.end_date })
                )
            );
        }
        IEnumerable<schedulerEvent> GetEvents()
        {
            var service = getService();
            // Define parameters of request.
            EventsResource.ListRequest request = service.Events.List("primary");
            request.TimeMin = new DateTime(2019, 01, 01);
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.MaxResults = 9999;

            Events events = request.Execute();

            List<schedulerEvent> eventsList = new List<schedulerEvent>();
            foreach (var e in events.Items)
            {
                eventsList.Add(new schedulerEvent
                {
                    gid= e.Id,
                    text = e.Summary ?? string.Empty,
                    start_date = (DateTime)e.Start.DateTime,
                    end_date = (DateTime)e.End.DateTime
                });
            }
            return eventsList;
            }
        public Google.Apis.Calendar.v3.Data.Event transformEvent(schedulerEvent changedEvent)
        {
            Event googleCalendarEvent = new Event();
            googleCalendarEvent.Start = new Google.Apis.Calendar.v3.Data.EventDateTime { DateTime = changedEvent.start_date.ToUniversalTime() };
            googleCalendarEvent.End = new Google.Apis.Calendar.v3.Data.EventDateTime { DateTime = changedEvent.end_date.ToUniversalTime() };
            
            googleCalendarEvent.Summary = changedEvent.text;
            return googleCalendarEvent;

        }

        public ContentResult Save(int? id, FormCollection actionValues)
        {
            var action = new DataAction(actionValues);
            var changedEvent = DHXEventsHelper.Bind<schedulerEvent>(actionValues);
            var entities = new CalendarContext();
            var service = getService();
            try
            {
                switch (action.Type)
                {
                    case DataActionTypes.Insert:
                        //System.Diagnostics.Debug.WriteLine(changedEvent.text);
                        Event ge = service.Events.Insert(transformEvent(changedEvent), "primary").Execute();
                        changedEvent.user = User.Identity.GetUserName();
                        changedEvent.gid = ge.Id;
                        entities.Events.Add(changedEvent);
                        break;
                    case DataActionTypes.Delete:
                       // System.Diagnostics.Debug.WriteLine(changedEvent.gid);
                        service.Events.Delete("primary", changedEvent.gid).Execute();
                        changedEvent = entities.Events.FirstOrDefault(ev => ev.gid.Equals(changedEvent.gid));
                        entities.Events.Remove(changedEvent);
                        break;
                    default:// "update"
                        System.Diagnostics.Debug.WriteLine(changedEvent.gid);
                        var target = entities.Events.Single(e => e.gid.Equals(changedEvent.gid));
                        DHXEventsHelper.Update(target, changedEvent, new List<string> { "id" });
                        target = entities.Events.Single(e => e.gid.Equals(changedEvent.gid));
                        var toUpdate = service.Events.Get("primary", target.gid).Execute();
                        toUpdate.Summary = target.text;
                        toUpdate.Start = new Google.Apis.Calendar.v3.Data.EventDateTime
                        {
                            DateTime = target.start_date.ToUniversalTime()
                        };
                        toUpdate.End = new Google.Apis.Calendar.v3.Data.EventDateTime
                        {
                            DateTime = target.end_date.ToUniversalTime()
                        };
                        service.Events.Update(toUpdate, "primary", toUpdate.Id).Execute();
                        break;
                }
                entities.SaveChanges();
                action.TargetId = changedEvent.id;
            }
            catch (Exception)
            {
                action.Type = DataActionTypes.Error;
            }
            
              return (new AjaxSaveResponse(action));
        }

        private UserCredential getCredentials()
        {
            //System.Diagnostics.Debug.WriteLine(GetEvents());
            var userName = User.Identity.GetUserName();
            UserCredential credential;
            using (var stream =
                new FileStream(System.AppDomain.CurrentDomain.BaseDirectory + "/App_Data/credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = System.AppDomain.CurrentDomain.BaseDirectory + "tokens/" + userName + "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                System.Diagnostics.Debug.WriteLine("Credential file saved to: " + credPath);
            }
            return credential;
        }
        private Google.Apis.Calendar.v3.CalendarService getService()
        {
            try
            {  // Create Google Calendar API service.
                var credential = getCredentials();
                var service = new Google.Apis.Calendar.v3.CalendarService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Calendar"
                });

                return service;
            }
            catch (Exception)
            {
                return null;
            }
        }
        public ActionResult ExcelExport()
        {
            CalendarContext db = new CalendarContext();
            var user = User.Identity.GetUserName();
            var eventsList = new CalendarContext().Events.
                Where(e => e.user.Equals(user)).
            Select(e => new { e.id, e.text, e.start_date, e.end_date }).ToList();

            try
            {

                DataTable Dt = new DataTable();
                Dt.Columns.Add("ID", typeof(string));
                Dt.Columns.Add("Text", typeof(string));
                Dt.Columns.Add("Start", typeof(string));
                Dt.Columns.Add("End", typeof(string));

                foreach (var data in eventsList)
                {
                    DataRow row = Dt.NewRow();
                    row[0] = data.id;
                    row[1] = data.text;
                    row[2] = data.start_date.ToString();
                    row[3] = data.end_date.ToString();
                    Dt.Rows.Add(row);

                }

                var memoryStream = new MemoryStream();
                using (var excelPackage = new ExcelPackage(memoryStream))
                {
                    var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
                    worksheet.Cells["A1"].LoadFromDataTable(Dt, true, TableStyles.None);
                    worksheet.Cells["A1:AN1"].Style.Font.Bold = true;
                    worksheet.DefaultRowHeight = 18;


                    worksheet.Column(2).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    worksheet.Column(6).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet.Column(7).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet.DefaultColWidth = 20;
                    worksheet.Column(2).AutoFit();

                    Session["DownloadExcel_FileManager"] = excelPackage.GetAsByteArray();
                    return Json("", JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        public ActionResult Download()
        {

            if (Session["DownloadExcel_FileManager"] != null)
            {
                byte[] data = Session["DownloadExcel_FileManager"] as byte[];
                return File(data, "application/octet-stream", "Events.xlsx");
            }
            else
            {
                return new EmptyResult();
            }
        }
    }
}