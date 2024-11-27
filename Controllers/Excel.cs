using System;
using Microsoft.AspNetCore.Mvc;
using reportesApi.Services;
using reportesApi.Models;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Hosting;
using System.IO;
using Microsoft.Extensions.Logging;

namespace reportesApi.Controllers
{
    [Route("api")]
    public class ExcelController: ControllerBase
    {
        private readonly ExcelService _ExcelService;
        private readonly ILogger<ExcelController> _logger;

        public ExcelController(ExcelService ExcelService, ILogger<ExcelController> logger)
        {
            _ExcelService = ExcelService;
            _logger = logger;
        }

    }
}