using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ProjectApi.Models
{
    public class Task
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public double Work { get; set; }
        public dynamic Start { get; set; }
        public dynamic Finish { get; set; }
        public IEnumerable<ProjectApi.Models.Resource> Resources { get; set; }

    }
}