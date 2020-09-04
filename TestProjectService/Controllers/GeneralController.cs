using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace TestProjectService.Controllers
{
    [Route("")]
    [ApiController]
    public class GeneralController : ControllerBase
    {

        [HttpGet("test")]
        public async Task<IActionResult> Test()
        {
            var logic = new Logic();
            logic.DoStuff();
            return Ok();
        }


    }
}
