using System;
using Microsoft.AspNetCore.Mvc;
using reportesApi.Services;
using Microsoft.Extensions.Logging;
using ClosedXML.Excel;
using System.IO;
using System.Net;

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

        // Método para exportar recetas por rango de fechas
        [HttpGet("ExportarRecetas")]
        public IActionResult ExportarRecetas(DateTime? fechaInicio = null, DateTime? fechaFin = null, bool download = false)
        {
            try
            {
                var recetas = _excelService.GetRecetasPorRangoDeFechas(fechaInicio ?? DateTime.MinValue, fechaFin ?? DateTime.MaxValue);

                if (recetas == null || recetas.Count == 0)
                {
                    return NoContent();
                }

                if (download)
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Recetas");
                        worksheet.Cell(1, 1).Value = $"Reporte de recetas {(fechaInicio.HasValue ? fechaInicio.Value.ToString("yyyy-MM-dd") : "Inicio")} al {(fechaFin.HasValue ? fechaFin.Value.ToString("yyyy-MM-dd") : "Fin")}";
                        worksheet.Range(1, 1, 1, 7).Merge().Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        worksheet.Cell(2, 1).InsertTable(recetas);
                        worksheet.Columns().AdjustToContents();

                        using (var stream = new MemoryStream())
                        {
                            workbook.SaveAs(stream);
                            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ReporteRecetas.xlsx");
                        }
                    }
                }

                return Ok(recetas);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error al exportar recetas: {ex.Message}");
                return StatusCode((int)HttpStatusCode.InternalServerError, $"Error: {ex.Message}");
            }
        }

        // Método para exportar el contenido de una receta por ID
        [HttpGet("ExportarContenidoReceta")]
        public IActionResult ExportarContenidoReceta(int idReceta, bool download = false)
        {
            try
            {
                var contenidoReceta = _excelService.GetContenidoRecetaPorId(idReceta);

                if (contenidoReceta == null || contenidoReceta.Count == 0)
                {
                    return NoContent();
                }

                if (download)
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("ContenidoReceta");
                        worksheet.Cell(1, 1).Value = $"Contenido de receta con ID {idReceta}";
                        worksheet.Range(1, 1, 1, 7).Merge().Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        worksheet.Cell(2, 1).InsertTable(contenidoReceta);
                        worksheet.Columns().AdjustToContents();

                        using (var stream = new MemoryStream())
                        {
                            workbook.SaveAs(stream);
                            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ContenidoReceta.xlsx");
                        }
                    }
                }

                return Ok(contenidoReceta);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error al exportar contenido de receta: {ex.Message}");
                return StatusCode((int)HttpStatusCode.InternalServerError, $"Error: {ex.Message}");
            }
        }

        // Método para exportar traspasos de entrada
        [HttpGet("ExportarTraspasosEntrada")]
        public IActionResult ExportarTraspasosEntrada(DateTime? fechaInicio = null, DateTime? fechaFin = null, int almacenDestino = 0, bool download = false)
        {
            try
            {
                var traspasosEntrada = _excelService.GetTraspasosEntrada(fechaInicio ?? DateTime.MinValue, fechaFin ?? DateTime.MaxValue, almacenDestino);

                if (traspasosEntrada == null || traspasosEntrada.Count == 0)
                {
                    return NoContent();
                }

                if (download)
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("TraspasosEntrada");
                        worksheet.Cell(1, 1).Value = $"Traspasos de entrada {(fechaInicio.HasValue ? fechaInicio.Value.ToString("yyyy-MM-dd") : "Inicio")} al {(fechaFin.HasValue ? fechaFin.Value.ToString("yyyy-MM-dd") : "Fin")}";
                        worksheet.Range(1, 1, 1, 14).Merge().Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        worksheet.Cell(2, 1).InsertTable(traspasosEntrada);
                        worksheet.Columns().AdjustToContents();

                        using (var stream = new MemoryStream())
                        {
                            workbook.SaveAs(stream);
                            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "TraspasosEntrada.xlsx");
                        }
                    }
                }

                return Ok(traspasosEntrada);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error al exportar traspasos de entrada: {ex.Message}");
                return StatusCode((int)HttpStatusCode.InternalServerError, $"Error: {ex.Message}");
            }
        }

        // Método para exportar traspasos de salida
        [HttpGet("ExportarTraspasosSalida")]
        public IActionResult ExportarTraspasosSalida(DateTime? fechaInicio = null, DateTime? fechaFin = null, int almacenOrigen = 0, bool download = false)
        {
            try
            {
                var traspasosSalida = _excelService.GetTraspasosSalida(fechaInicio ?? DateTime.MinValue, fechaFin ?? DateTime.MaxValue, almacenOrigen);

                if (traspasosSalida == null || traspasosSalida.Count == 0)
                {
                    return NoContent();
                }

                if (download)
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("TraspasosSalida");
                        worksheet.Cell(1, 1).Value = $"Traspasos de salida {(fechaInicio.HasValue ? fechaInicio.Value.ToString("yyyy-MM-dd") : "Inicio")} al {(fechaFin.HasValue ? fechaFin.Value.ToString("yyyy-MM-dd") : "Fin")}";
                        worksheet.Range(1, 1, 1, 14).Merge().Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        worksheet.Cell(2, 1).InsertTable(traspasosSalida);
                        worksheet.Columns().AdjustToContents();

                        using (var stream = new MemoryStream())
                        {
                            workbook.SaveAs(stream);
                            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "TraspasosSalida.xlsx");
                        }
                    }
                }

                return Ok(traspasosSalida);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error al exportar traspasos de salida: {ex.Message}");
                return StatusCode((int)HttpStatusCode.InternalServerError, $"Error: {ex.Message}");
            }
        }
    }
}
