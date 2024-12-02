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

        // 1. Reporte de Recetas por Rango de Fechas
        [HttpGet("recetas-rango-fechas")]
        public IActionResult GetRecetasPorRangoDeFechasExcel(DateTime fechaInicio, DateTime fechaFin)
        {
            try
            {
                var recetas = _excelService.GetRecetasPorRangoDeFechas(fechaInicio, fechaFin);

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Recetas");

                    // Encabezado
                    worksheet.Cells[1, 1].Value = "RecetaID";
                    worksheet.Cells[1, 2].Value = "NombreReceta";
                    worksheet.Cells[1, 3].Value = "FechaCreacion";
                    worksheet.Cells[1, 4].Value = "UsuarioRegistra";
                    worksheet.Row(1).Style.Font.Bold = true;

                    // Datos
                    int row = 2;
                    foreach (var receta in recetas)
                    {
                        worksheet.Cells[row, 1].Value = receta.RecetaID;
                        worksheet.Cells[row, 2].Value = receta.NombreReceta;
                        worksheet.Cells[row, 3].Value = receta.FechaCreacion;
                        worksheet.Cells[row, 4].Value = receta.UsuarioRegistra;
                        row++;
                    }

                    // Ajuste de columnas
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    // Guardar el archivo en memoria
                    var fileContents = package.GetAsByteArray();
                    return File(fileContents, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "RecetasPorRango.xlsx");
                }
            }
            catch (Exception ex)
            {
                return BadRequest("Error al generar el reporte: " + ex.Message);
            }
        }

        // 2. Reporte de Contenido de Receta por ID
        [HttpGet("contenido-receta/{idReceta}")]
        public IActionResult GetContenidoRecetaPorIdExcel(int idReceta)
        {
            try
            {
                var contenidoReceta = _excelService.GetContenidoRecetaPorId(idReceta);

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Contenido Receta");

                    // Encabezado
                    worksheet.Cells[1, 1].Value = "RecetaID";
                    worksheet.Cells[1, 2].Value = "NombreReceta";
                    worksheet.Cells[1, 3].Value = "Insumo";
                    worksheet.Cells[1, 4].Value = "Descripcion Insumo";
                    worksheet.Cells[1, 5].Value = "FechaCreacion";
                    worksheet.Cells[1, 6].Value = "UsuarioRegistra";
                    worksheet.Row(1).Style.Font.Bold = true;

                    // Datos
                    int row = 2;
                    foreach (var contenido in contenidoReceta)
                    {
                        worksheet.Cells[row, 1].Value = contenido.RecetaID;
                        worksheet.Cells[row, 2].Value = contenido.NombreReceta;
                        worksheet.Cells[row, 3].Value = contenido.Insumo;
                        worksheet.Cells[row, 4].Value = contenido.DescripcionInsumo;
                        worksheet.Cells[row, 5].Value = contenido.FechaCreacion;
                        worksheet.Cells[row, 6].Value = contenido.UsuarioRegistra;
                        row++;
                    }

                    // Ajuste de columnas
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    // Guardar el archivo en memoria
                    var fileContents = package.GetAsByteArray();
                    return File(fileContents, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ContenidoReceta.xlsx");
                }
            }
            catch (Exception ex)
            {
                return BadRequest("Error al generar el reporte: " + ex.Message);
            }
        }

        // 3. Reporte de Traspasos de Entrada
        [HttpGet("traspasos-entrada")]
        public IActionResult GetTraspasosEntradaExcel(DateTime fechaInicio, DateTime fechaFin, int almacen, int tipoMovimiento)
        {
            try
            {
                var traspasosEntrada = _excelService.GetTraspasosEntrada(fechaInicio, fechaFin, almacen, tipoMovimiento);

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Traspasos Entrada");

                    // Encabezado
                    worksheet.Cells[1, 1].Value = "IdTRSP";
                    worksheet.Cells[1, 2].Value = "AlmacenOrigen";
                    worksheet.Cells[1, 3].Value = "NombreAlmacenOrigen";
                    worksheet.Cells[1, 4].Value = "AlmacenDestino";
                    worksheet.Cells[1, 5].Value = "NombreAlmacenDestino";
                    worksheet.Cells[1, 6].Value = "IdInsumo";
                    worksheet.Cells[1, 7].Value = "DescripcionInsumo";
                    worksheet.Cells[1, 8].Value = "FechaEntrada";
                    worksheet.Cells[1, 9].Value = "FechaSalida";
                    worksheet.Cells[1, 10].Value = "Cantidad";
                    worksheet.Cells[1, 11].Value = "TipoMovimiento";
                    worksheet.Cells[1, 12].Value = "Descripcion";
                    worksheet.Cells[1, 13].Value = "NoFolio";
                    worksheet.Cells[1, 14].Value = "FechaRegistro";
                    worksheet.Cells[1, 15].Value = "UsuarioRegistra";
                    worksheet.Row(1).Style.Font.Bold = true;

                    // Datos
                    int row = 2;
                    foreach (var traspaso in traspasosEntrada)
                    {
                        worksheet.Cells[row, 1].Value = traspaso.IdTRSP;
                        worksheet.Cells[row, 2].Value = traspaso.AlmacenOrigen;
                        worksheet.Cells[row, 3].Value = traspaso.NombreAlmacenOrigen;
                        worksheet.Cells[row, 4].Value = traspaso.AlmacenDestino;
                        worksheet.Cells[row, 5].Value = traspaso.NombreAlmacenDestino;
                        worksheet.Cells[row, 6].Value = traspaso.IdInsumo;
                        worksheet.Cells[row, 7].Value = traspaso.DescripcionInsumo;
                        worksheet.Cells[row, 8].Value = traspaso.FechaEntrada;
                        worksheet.Cells[row, 9].Value = traspaso.FechaSalida;
                        worksheet.Cells[row, 10].Value = traspaso.Cantidad;
                        worksheet.Cells[row, 11].Value = traspaso.TipoMovimiento;
                        worksheet.Cells[row, 12].Value = traspaso.Descripcion;
                        worksheet.Cells[row, 13].Value = traspaso.NoFolio;
                        worksheet.Cells[row, 14].Value = traspaso.FechaRegistro;
                        worksheet.Cells[row, 15].Value = traspaso.UsuarioRegistra;
                        row++;
                    }

                    // Ajuste de columnas
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    // Guardar el archivo en memoria
                    var fileContents = package.GetAsByteArray();
                    return File(fileContents, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "TraspasosEntrada.xlsx");
                }
            }
            catch (Exception ex)
            {
                return BadRequest("Error al generar el reporte: " + ex.Message);
            }
        }

        // 4. Reporte de Traspasos de Salida
        [HttpGet("traspasos-salida")]
        public IActionResult GetTraspasosSalidaExcel(DateTime fechaInicio, DateTime fechaFin, int almacen, int tipoMovimiento)
        {
            try
            {
                var traspasosSalida = _excelService.GetTraspasosSalida(fechaInicio, fechaFin, almacen, tipoMovimiento);

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Traspasos Salida");

                    // Encabezado
                    worksheet.Cells[1, 1].Value = "IdTRSP";
                    worksheet.Cells[1, 2].Value = "AlmacenOrigen";
                    worksheet.Cells[1, 3].Value = "NombreAlmacenOrigen";
                    worksheet.Cells[1, 4].Value = "AlmacenDestino";
                    worksheet.Cells[1, 5].Value = "NombreAlmacenDestino";
                    worksheet.Cells[1, 6].Value = "IdInsumo";
                    worksheet.Cells[1, 7].Value = "DescripcionInsumo";
                    worksheet.Cells[1, 8].Value = "FechaEntrada";
                    worksheet.Cells[1, 9].Value = "FechaSalida";
                    worksheet.Cells[1, 10].Value = "Cantidad";
                    worksheet.Cells[1, 11].Value = "TipoMovimiento";
                    worksheet.Cells[1, 12].Value = "Descripcion";
                    worksheet.Cells[1, 13].Value = "NoFolio";
                    worksheet.Cells[1, 14].Value = "FechaRegistro";
                    worksheet.Cells[1, 15].Value = "UsuarioRegistra";
                    worksheet.Row(1).Style.Font.Bold = true;

                    // Datos
                    int row = 2;
                    foreach (var traspaso in traspasosSalida)
                    {
                        worksheet.Cells[row, 1].Value = traspaso.IdTRSP;
                        worksheet.Cells[row, 2].Value = traspaso.AlmacenOrigen;
                        worksheet.Cells[row, 3].Value = traspaso.NombreAlmacenOrigen;
                        worksheet.Cells[row, 4].Value = traspaso.AlmacenDestino;
                        worksheet.Cells[row, 5].Value = traspaso.NombreAlmacenDestino;
                        worksheet.Cells[row, 6].Value = traspaso.IdInsumo;
                        worksheet.Cells[row, 7].Value = traspaso.DescripcionInsumo;
                        worksheet.Cells[row, 8].Value = traspaso.FechaEntrada;
                        worksheet.Cells[row, 9].Value = traspaso.FechaSalida;
                        worksheet.Cells[row, 10].Value = traspaso.Cantidad;
                        worksheet.Cells[row, 11].Value = traspaso.TipoMovimiento;
                        worksheet.Cells[row, 12].Value = traspaso.Descripcion;
                        worksheet.Cells[row, 13].Value = traspaso.NoFolio;
                        worksheet.Cells[row, 14].Value = traspaso.FechaRegistro;
                        worksheet.Cells[row, 15].Value = traspaso.UsuarioRegistra;
                        row++;
                    }

                    // Ajuste de columnas
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    // Guardar el archivo en memoria
                    var fileContents = package.GetAsByteArray();
                    return File(fileContents, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "TraspasosSalida.xlsx");
                }
            }
            catch (Exception ex)
            {
                return BadRequest("Error al generar el reporte: " + ex.Message);
            }
        }
    }
}
