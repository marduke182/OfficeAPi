using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using Microsoft.Office.Interop.MSProject;
using System.Reflection;

namespace MvcApplication5.Controllers
{
    public class ValuesController : ApiController
    {
        // GET api/values
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        
        public HttpResponseMessage Post()
        {
            return Request.CreateResponse(HttpStatusCode.OK);

        }
        
        
        
        [HttpPost]
        [ActionName("Upload")]
        public async Task<ProjectApi.Models.Project> Upload()
        {
            

            // Check if the request contains multipart/form-data.
            if (!Request.Content.IsMimeMultipartContent())
            {
                throw new HttpResponseException(HttpStatusCode.UnsupportedMediaType);
            }

            string root = HttpContext.Current.Server.MapPath("~/App_Data");
            var provider = new MultipartFormDataStreamProvider(root);

            try
            {
                 //Read the form data.
                await Request.Content.ReadAsMultipartAsync(provider);
                       //This illustrates how to get the file names.
                foreach (MultipartFileData file in provider.FileData)
                {

                                    object readOnly = false;

                                    Microsoft.Office.Interop.MSProject.PjMergeType merge = Microsoft.Office.Interop.MSProject.PjMergeType.pjDoNotMerge;

                                    Microsoft.Office.Interop.MSProject.PjPoolOpen pool = Microsoft.Office.Interop.MSProject.PjPoolOpen.pjDoNotOpenPool;

                                    object ignoreReadOnlyRecommended = false;


                                    Application projectApp = new Application();
                                    projectApp.FileOpen(file.LocalFileName, readOnly, merge, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                pool, Missing.Value, Missing.Value, ignoreReadOnlyRecommended, Missing.Value);
                                    Project proj = projectApp.ActiveProject;
                ProjectApi.Models.Project myproj = new ProjectApi.Models.Project();
                                    myproj.Id = proj.ID;
                myproj.Name = proj.Name;

                                    List<ProjectApi.Models.Resource> listproj = new List<ProjectApi.Models.Resource>();

                                    foreach (Microsoft.Office.Interop.MSProject.Resource resource in proj.Resources)
                                    {
                                        ProjectApi.Models.Resource myresource = new ProjectApi.Models.Resource();
                                        myresource.Id = resource.UniqueID;
                                        myresource.Name = resource.Name;
                                        myresource.Work = resource.Work;
                                        listproj.Add(myresource);
                                    }
                                    myproj.Resources = listproj;

                                    List<ProjectApi.Models.Task> listtask = new List<ProjectApi.Models.Task>();

                                    foreach (Microsoft.Office.Interop.MSProject.Task task in proj.Tasks)
                                    {
                                        ProjectApi.Models.Task mytask = new ProjectApi.Models.Task();

                                        mytask.Id = task.UniqueID;
                                        mytask.Name = task.Name;
                                        mytask.Start = task.Start;
                                        mytask.Finish= task.Finish;
                                        mytask.Work= task.Work;
                                        Resources resources = task.Resources;
                                        Microsoft.Office.Interop.MSProject.Tasks children = task.OutlineChildren;
                                        int count = children.Count;
                                        Microsoft.Office.Interop.MSProject.Task parent = task.OutlineParent;
                                        foreach (Microsoft.Office.Interop.MSProject.Resource resource in resources)
                                        {
                                            string name_resource = resource.Name;
                                            double work_resource = resource.Work;
                                        }
                                        listtask.Add(mytask);
                                    }

                                    myproj.task = listtask;

                                    Calendar calendar = proj.Calendar;
                                    string calendar_name = calendar.Name;
                                    WeekDays weekDays = calendar.WeekDays;
                                    WorkWeeks wws = calendar.WorkWeeks;
                                    foreach (Microsoft.Office.Interop.MSProject.WorkWeek ww in wws)
                                    {
                                        WorkWeekDays name_resource = ww.WeekDays;
                                        foreach (Microsoft.Office.Interop.MSProject.WorkWeekDay wwd in name_resource)
                                        {
                                            string tmp = wwd.Name;
                                        }
                                    }

                                    Exceptions exc = calendar.Exceptions;
                                    foreach (Microsoft.Office.Interop.MSProject.Exception exce in exc)
                                    {
                                        int name_resource = exce.DaysOfWeek;
                                        string tmp = exce.Name;
                                        int tm2p = exce.Occurrences;
                                        PjMonth tmp3 = exce.Month;
                                        int tmp4 = exce.MonthDay;
                                        PjExceptionItem tmp5 = exce.MonthItem;
                                        PjExceptionPosition tmp6 = exce.MonthPosition;
                                        int tmp7 = exce.Period;
                                        dynamic start = exce.Start;
                                        int tmp8 = exce.DaysOfWeek;
                                    }

                                    projectApp.Application.FileSaveAs(file.LocalFileName);
                                    projectApp.Application.Quit(PjSaveType.pjDoNotSave);

                    return myproj;
                }
                return null;
            }
            catch (System.Exception e)
            {
                Trace.WriteLine(e);
                throw new HttpResponseException(HttpStatusCode.Conflict);
            }
        }
        // GET api/values/5
        public string Get(int id)
        {
            return "value";
        }

        

        // PUT api/values/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/values/5
        public void Delete(int id)
        {
        }
    }
}