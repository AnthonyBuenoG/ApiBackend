using System;
using Microsoft.AspNetCore.Mvc;
using reportesApi.Services;
using reportesApi.Utilities;
using Microsoft.AspNetCore.Authorization;
using reportesApi.Models;
using Microsoft.Extensions.Logging;
using System.Net;
using reportesApi.Helpers;
using Newtonsoft.Json;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Microsoft.AspNetCore.Hosting;
using reportesApi.Models.Compras;
using ClosedXML.Excel;

namespace reportesApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly ExcelService _excelService;
        private readonly ILogger<ExcelController> _logger;

        public ExcelController(ExcelService excelService, ILogger<ExcelController> logger)
        {
            _excelService = excelService;
            _logger = logger;
        }

   

        [HttpGet("ExportarRecetas")]
        public IActionResult ExportarRecetas(DateTime? FechaInicio = null, DateTime? FechaFin = null)
        {
            var objectResponse = Helper.GetStructResponse();

            try
            {
                // Llamar al método para obtener las recetas según las fechas proporcionadas
                var recetas = _excelService.GetRecetasPorRangoDeFechas(FechaInicio ?? DateTime.MinValue, FechaFin ?? DateTime.MaxValue);

                if (recetas == null || recetas.Count == 0)
                {
                    objectResponse.StatusCode = (int)HttpStatusCode.NoContent;
                    objectResponse.message = "No hay datos disponibles para exportar.";
                    return new JsonResult(objectResponse);
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Recetas");

                    // Encabezado con el rango de fechas
                    var fechaInicioText = FechaInicio.HasValue ? FechaInicio.Value.ToString("yyyy-MM-dd") : "Inicio";
                    var fechaFinText = FechaFin.HasValue ? FechaFin.Value.ToString("yyyy-MM-dd") : "Fin";
                    worksheet.Cell(1, 1).Value = $"Reporte de recetas {fechaInicioText} al {fechaFinText}";
                    worksheet.Range(1, 1, 1, 4).Merge().Style
                        .Font.SetBold(true)
                        .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                    // Encabezados de columnas
                    var headers = new string[] { "ID Receta", "Nombre Receta", "Fecha Creación", "Usuario Registra" };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        worksheet.Cell(2, i + 1).Value = headers[i];
                        worksheet.Cell(2, i + 1).Style.Font.Bold = true;
                        worksheet.Cell(2, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }

                    // Llenar los datos de las recetas en el archivo Excel
                    int currentRow = 3;
                    foreach (var receta in recetas)
                    {
                        worksheet.Cell(currentRow, 1).Value = receta.RecetaID;
                        worksheet.Cell(currentRow, 2).Value = receta.NombreReceta;
                        worksheet.Cell(currentRow, 3).Value = receta.FechaCreacion;
                        worksheet.Cell(currentRow, 4).Value = receta.UsuarioRegistra;
                        currentRow++;
                    }

                    // Ajustar el tamaño de las columnas
                    worksheet.Columns().AdjustToContents();

                    // Guardar el archivo Excel en un MemoryStream
                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        var content = stream.ToArray();

                        // Devolver el archivo Excel como respuesta de descarga
                        return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Recetas.xlsx");
                    }
                }
            }
            catch (Exception ex)
            {
                // Manejo de excepciones
                objectResponse.StatusCode = (int)HttpStatusCode.InternalServerError;
                objectResponse.message = $"Error: {ex.Message}";
                return new JsonResult(objectResponse);
            }
        }
    }
}
