using ApiExcelToDB.HNX_UPCOM;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ApiExcelToDB.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class HnxUpcomController : ControllerBase
    {
        [HttpGet]
        public IActionResult Get(string name)
        {
           if (!string.IsNullOrEmpty(name))
            {
               
                    IConfiguration configuration = new ConfigurationBuilder()
                   .SetBasePath(Directory.GetCurrentDirectory())
                   .AddJsonFile("appsettings.json")
                   .Build();
                    ReadFileUpcom fileUpcom = new ReadFileUpcom(configuration,name);
                return Ok($"Read file HNX + UPCOM Ok,.../ {name}!");
            }
            else
            {

                return Ok($"Chưa read file");
            }
           // Xử lý logic dựa trên tham số name ở đây

        }
    }
}
