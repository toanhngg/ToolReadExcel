using ApiExcelToDB.HNX_UPCOM;
using ApiExcelToDB.HOSE;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System.IO;

namespace ApiExcelToDB.Controllers
{
    [ApiController]
    [Route("api/[controller]")]

    public class HsxController : Controller
    {
        [HttpGet]
        public IActionResult Get(string name)
        {
            if (!string.IsNullOrEmpty(name))
            {
                ReadExcelHose run = new ReadExcelHose(name);
                return Ok($"Read file HSX Ok,.../ {name}!");
            }
            else
            {
                return Ok($"Chưa read file");
            }
        
         
        }
    }
}
