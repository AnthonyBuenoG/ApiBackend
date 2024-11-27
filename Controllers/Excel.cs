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
   
    [Route("api")]
    public class ExcelController: ControllerBase
    {
   
        private readonly ExcelService _ExcelService;
        private readonly ILogger<ExcelController> _logger;
  
        private readonly IJwtAuthenticationService _authService;
        private readonly IWebHostEnvironment _hostingEnvironment;
        

        Encrypt enc = new Encrypt();

        public ExcelController(ExcelService ExcelService, ILogger<ExcelController> logger, IJwtAuthenticationService authService) {
            _ExcelService = ExcelService;
            _logger = logger;
       
            _authService = authService;
            // Configura la ruta base donde se almacenan los archivos.
            // Asegúrate de ajustar la ruta según tu estructura de directorios.

            
            
        }
   [HttpGet("ExportarRecetasFechas")]
        public IActionResult ExportarRecetasRangoFechas([FromQuery] DateTime fechaInicio, [FromQuery] DateTime fechaFin)
        {
            try
            {
                var recetas = _ExcelService.GetRecetasPorRangoDeFechas(fechaInicio, fechaFin);

                // Crear archivo Excel
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Recetas");
                    worksheet.Cell(1, 1).Value = "RecetaID";
                    worksheet.Cell(1, 2).Value = "NombreReceta";
                    worksheet.Cell(1, 3).Value = "FechaCreacion";
                    worksheet.Cell(1, 4).Value = "UsuarioRegistra";
                    worksheet.Cell(1, 5).Value = "Insumo";
                    worksheet.Cell(1, 6).Value = "DescripcionInsumo";
                    worksheet.Cell(1, 7).Value = "Cantidad";

                    var row = 2;
                    foreach (var receta in recetas)
                    {
                        worksheet.Cell(row, 1).Value = receta.RecetaID;
                        worksheet.Cell(row, 2).Value = receta.NombreReceta;
                        worksheet.Cell(row, 3).Value = receta.FechaCreacion;
                        worksheet.Cell(row, 4).Value = receta.UsuarioRegistra;
                        worksheet.Cell(row, 5).Value = receta.Insumo;
                        worksheet.Cell(row, 6).Value = receta.DescripcionInsumo;
                        worksheet.Cell(row, 7).Value = receta.Cantidad;
                        row++;
                    }

                    var fileName = "Recetas_Rango_Fechas.xlsx";
                    using (var memoryStream = new MemoryStream())
                    {
                        workbook.SaveAs(memoryStream);
                        memoryStream.Seek(0, SeekOrigin.Begin);
                        return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error al exportar las recetas: {ex.Message}");
                return StatusCode(500, "Error interno del servidor.");
            }
        }

        [HttpGet("ExportarRecetaId")]
        public IActionResult ExportarContenidoReceta([FromQuery] int idReceta)
        {
            try
            {
                var contenidoReceta = _ExcelService.GetContenidoRecetaPorId(idReceta);

                // Crear archivo Excel
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Contenido Receta");
                    worksheet.Cell(1, 1).Value = "RecetaID";
                    worksheet.Cell(1, 2).Value = "NombreReceta";
                    worksheet.Cell(1, 3).Value = "FechaCreacion";
                    worksheet.Cell(1, 4).Value = "UsuarioRegistra";
                    worksheet.Cell(1, 5).Value = "Insumo";
                    worksheet.Cell(1, 6).Value = "DescripcionInsumo";
                    worksheet.Cell(1, 7).Value = "Cantidad";

                    var row = 2;
                    foreach (var contenido in contenidoReceta)
                    {
                        worksheet.Cell(row, 1).Value = contenido.RecetaID;
                        worksheet.Cell(row, 2).Value = contenido.NombreReceta;
                        worksheet.Cell(row, 3).Value = contenido.FechaCreacion;
                        worksheet.Cell(row, 4).Value = contenido.UsuarioRegistra;
                        worksheet.Cell(row, 5).Value = contenido.Insumo;
                        worksheet.Cell(row, 6).Value = contenido.DescripcionInsumo;
                        worksheet.Cell(row, 7).Value = contenido.Cantidad;
                        row++;
                    }

                    var fileName = "Contenido_Receta";
                    using (var memoryStream = new MemoryStream())
                    {
                        workbook.SaveAs(memoryStream);
                        memoryStream.Seek(0, SeekOrigin.Begin);
                        return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error al exportar el contenido de la receta: {ex.Message}");
                return StatusCode(500, "Error interno del servidor.");
            }
        }

        [HttpGet("ExportarTraspasosEntrada")]
        public IActionResult ExportarTraspasosEntrada([FromQuery] DateTime fechaInicio, [FromQuery] DateTime fechaFin, [FromQuery] int almacenDestino)
        {
            try
            {
                var traspasosEntrada = _ExcelService.GetTraspasosEntrada(fechaInicio, fechaFin, almacenDestino);

                // Crear archivo Excel
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Traspasos Entrada");
                    worksheet.Cell(1, 1).Value = "IdTRSP";
                    worksheet.Cell(1, 2).Value = "AlmacenOrigen";
                    worksheet.Cell(1, 3).Value = "NombreAlmacenOrigen";
                    worksheet.Cell(1, 4).Value = "AlmacenDestino";
                    worksheet.Cell(1, 5).Value = "NombreAlmacenDestino";
                    worksheet.Cell(1, 6).Value = "IdInsumo";
                    worksheet.Cell(1, 7).Value = "DescripcionInsumo";
                    worksheet.Cell(1, 8).Value = "FechaEntrada";
                    worksheet.Cell(1, 9).Value = "FechaSalida";
                    worksheet.Cell(1, 10).Value = "Cantidad";
                    worksheet.Cell(1, 11).Value = "TipoMovimiento";
                    worksheet.Cell(1, 12).Value = "Descripcion";
                    worksheet.Cell(1, 13).Value = "No_Folio";
                    worksheet.Cell(1, 14).Value = "FechaRegistro";
                    worksheet.Cell(1, 15).Value = "Estatus";
                    worksheet.Cell(1, 16).Value = "UsuarioRegistra";

                    var row = 2;
                    foreach (var traspaso in traspasosEntrada)
                    {
                        worksheet.Cell(row, 1).Value = traspaso.IdTRSP;
                        worksheet.Cell(row, 2).Value = traspaso.AlmacenOrigen;
                        worksheet.Cell(row, 3).Value = traspaso.NombreAlmacenOrigen;
                        worksheet.Cell(row, 4).Value = traspaso.AlmacenDestino;
                        worksheet.Cell(row, 5).Value = traspaso.NombreAlmacenDestino;
                        worksheet.Cell(row, 6).Value = traspaso.IdInsumo;
                        worksheet.Cell(row, 7).Value = traspaso.DescripcionInsumo;
                        worksheet.Cell(row, 8).Value = traspaso.FechaEntrada;
                        worksheet.Cell(row, 9).Value = traspaso.FechaSalida;
                        worksheet.Cell(row, 10).Value = traspaso.Cantidad;
                        worksheet.Cell(row, 11).Value = traspaso.TipoMovimiento;
                        worksheet.Cell(row, 12).Value = traspaso.Descripcion;
                        worksheet.Cell(row, 13).Value = traspaso.NoFolio;
                        worksheet.Cell(row, 14).Value = traspaso.FechaRegistro;
                        worksheet.Cell(row, 15).Value = traspaso.Estatus;
                        worksheet.Cell(row, 16).Value = traspaso.UsuarioRegistra;
                        row++;
                    }

                    var fileName = "Traspasos_Entrada";
                    using (var memoryStream = new MemoryStream())
                    {
                        workbook.SaveAs(memoryStream);
                        memoryStream.Seek(0, SeekOrigin.Begin);
                        return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error al exportar los traspasos de entrada: {ex.Message}");
                return StatusCode(500, "Error interno del servidor.");
            }
        }

        [HttpGet("ExportarTraspasosSalida")]
        public IActionResult ExportarTraspasosSalida([FromQuery] DateTime fechaInicio, [FromQuery] DateTime fechaFin, [FromQuery] int almacenOrigen)
        {
            try
            {
                var traspasosSalida = _ExcelService.GetTraspasosSalida(fechaInicio, fechaFin, almacenOrigen);

                // Crear archivo Excel
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Traspasos Salida");
                    worksheet.Cell(1, 1).Value = "IdTRSP";
                    worksheet.Cell(1, 2).Value = "AlmacenOrigen";
                    worksheet.Cell(1, 3).Value = "NombreAlmacenOrigen";
                    worksheet.Cell(1, 4).Value = "AlmacenDestino";
                    worksheet.Cell(1, 5).Value = "NombreAlmacenDestino";
                    worksheet.Cell(1, 6).Value = "IdInsumo";
                    worksheet.Cell(1, 7).Value = "DescripcionInsumo";
                    worksheet.Cell(1, 8).Value = "FechaEntrada";
                    worksheet.Cell(1, 9).Value = "FechaSalida";
                    worksheet.Cell(1, 10).Value = "Cantidad";
                    worksheet.Cell(1, 11).Value = "TipoMovimiento";
                    worksheet.Cell(1, 12).Value = "Descripcion";
                    worksheet.Cell(1, 13).Value = "No_Folio";
                    worksheet.Cell(1, 14).Value = "FechaRegistro";
                    worksheet.Cell(1, 15).Value = "Estatus";
                    worksheet.Cell(1, 16).Value = "UsuarioRegistra";

                    var row = 2;
                    foreach (var traspaso in traspasosSalida)
                    {
                        worksheet.Cell(row, 1).Value = traspaso.IdTRSP;
                        worksheet.Cell(row, 2).Value = traspaso.AlmacenOrigen;
                        worksheet.Cell(row, 3).Value = traspaso.NombreAlmacenOrigen;
                        worksheet.Cell(row, 4).Value = traspaso.AlmacenDestino;
                        worksheet.Cell(row, 5).Value = traspaso.NombreAlmacenDestino;
                        worksheet.Cell(row, 6).Value = traspaso.IdInsumo;
                        worksheet.Cell(row, 7).Value = traspaso.DescripcionInsumo;
                        worksheet.Cell(row, 8).Value = traspaso.FechaEntrada;
                        worksheet.Cell(row, 9).Value = traspaso.FechaSalida;
                        worksheet.Cell(row, 10).Value = traspaso.Cantidad;
                        worksheet.Cell(row, 11).Value = traspaso.TipoMovimiento;
                        worksheet.Cell(row, 12).Value = traspaso.Descripcion;
                        worksheet.Cell(row, 13).Value = traspaso.NoFolio;
                        worksheet.Cell(row, 14).Value = traspaso.FechaRegistro;
                        worksheet.Cell(row, 15).Value = traspaso.Estatus;
                        worksheet.Cell(row, 16).Value = traspaso.UsuarioRegistra;
                        row++;
                    }

                    var fileName = "Traspasos_Salida";
                    using (var memoryStream = new MemoryStream())
                    {
                        workbook.SaveAs(memoryStream);
                        memoryStream.Seek(0, SeekOrigin.Begin);
                        return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error al exportar los traspasos de salida: {ex.Message}");
                return StatusCode(500, "Error interno del servidor.");
            }
        }
    }
}