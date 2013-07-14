using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ProjectApi.Models
{
    public class Project
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public IEnumerable<ProjectApi.Models.Resource> Resources { get; set; }
        public IEnumerable<ProjectApi.Models.Task> task { get; set; }

    }
}