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

        [HttpPost("test")]
        public async Task<IActionResult> Test([FromForm] IFormFile excelFile)
        {
            var logic = new Logic();
            logic.DoStuff(excelFile);
            return Ok();
        }
    }
}
