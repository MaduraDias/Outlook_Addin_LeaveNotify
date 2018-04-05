using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MockApi.Models;

namespace MockApi.Controllers
{
    [Produces("application/json")]
    [Route("api/[Controller]/")]
    public class EmployeeController : Controller
    {
        [HttpGet]
        [Route("{id}/LeaveNotifyEmailList")]
        public LeaveNotifyData LeaveNotifyEmailList([FromRoute]string id)
        {
            return new LeaveNotifyData()
            {
                 To = "manager@test.com",
                 CC = new string[] {"resourceManager@test.com"
                                 ,"accountManager@test.com"
                                 ,"pc@test.com"}
            };

        }
    }

}
