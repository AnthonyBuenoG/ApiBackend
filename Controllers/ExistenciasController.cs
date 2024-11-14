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
    public class ExistenciasController : ControllerBase
    {
   
        private readonly ExistenciasService _ExistenciasService;
        private readonly ILogger<ExistenciasController> _logger;
  
        private readonly IJwtAuthenticationService _authService;
        private readonly IWebHostEnvironment _hostingEnvironment;
        

        Encrypt enc = new Encrypt();

        public ExistenciasController(ExistenciasService ExistenciasService, ILogger<ExistenciasController> logger, IJwtAuthenticationService authService) {
            _ExistenciasService = ExistenciasService;
            _logger = logger;
       
            _authService = authService;
            // Configura la ruta base donde se almacenan los archivos.
            // Asegúrate de ajustar la ruta según tu estructura de directorios.

            
            
        }


        [HttpPost("InsertExistencias")]
        public IActionResult InsertExistencias([FromBody] InsertExistenciasModel req )
        {
            var objectResponse = Helper.GetStructResponse();
            try
            {
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = _ExistenciasService.InsertExistencias(req);

            }

            catch (System.Exception ex)
            {
                objectResponse.message = ex.Message;
            }

            return new JsonResult(objectResponse);
        }

         [HttpGet("GetExistencias")]
        public IActionResult GetExistencias()
        {
            var objectResponse = Helper.GetStructResponse();

            try
            {
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = "Data cargados exitosamente";
                var resultado = _ExistenciasService.GetExistencias();
               
               

                // Llamando a la función y recibiendo los dos valores.
               
                 objectResponse.response = resultado;
            }

            catch (System.Exception ex)
            {
                objectResponse.StatusCode = (int)HttpStatusCode.InternalServerError;
                objectResponse.success = false;
                objectResponse.message = ex.Message;
            }

            return new JsonResult(objectResponse);
        }

        // Método para obtener existencias con filtros opcionales
    // [HttpGet("GetExistencias")]
    // public IActionResult GetExistencias(DateTime? fechaInicio = null, DateTime? fechaFin = null, int? idAlmacen = null)
    // {
    //     var objectResponse = Helper.GetStructResponse();

    //     try
    //     {
    //         // Llamar al servicio con los filtros opcionales
    //         var resultado = _ExistenciasService.GetExistencias(fechaInicio, fechaFin, idAlmacen);

    //         objectResponse.StatusCode = (int)HttpStatusCode.OK;
    //         objectResponse.success = true;
    //         objectResponse.message = "Datos cargados exitosamente";
    //         objectResponse.response = resultado;
    //     }
    //     catch (System.Exception ex)
    //     {
    //         objectResponse.StatusCode = (int)HttpStatusCode.InternalServerError;
    //         objectResponse.success = false;
    //         objectResponse.message = ex.Message;
    //     }

    //     return new JsonResult(objectResponse);
    // }

    // Método para exportar existencias a Excel con filtros opcionales
        // [HttpGet("ExportarExistenciasExcel")]
        // public IActionResult ExportarExistenciasExcel(
        //     DateTime? fechaInicio = null,
        //     DateTime? fechaFin = null,
        //     int? idAlmacen = null,
        //     bool filtrarPorFecha = false,
        //     bool filtrarPorIdAlmacen = false)
        // {
        //     try
        //     {
        //         if (filtrarPorFecha && (!fechaInicio.HasValue || !fechaFin.HasValue))
        //         {
        //             return BadRequest(new { success = false, message = "Debe proporcionar tanto fechaInicio como fechaFin para el filtrado por fechas." });
        //         }

        //         var datos = _ExistenciasService.GetExistencias(
        //             filtrarPorFecha ? fechaInicio : null,
        //             filtrarPorFecha ? fechaFin : null,
        //             filtrarPorIdAlmacen ? idAlmacen : null);

        //         using (var workbook = new XLWorkbook())
        //         {
        //             var worksheet = workbook.Worksheets.Add("Existencias");

        //             string titulo = "Reporte de existencias";
        //             if (filtrarPorFecha && fechaInicio.HasValue && fechaFin.HasValue)
        //             {
        //                 titulo += $" del {fechaInicio.Value:yyyy-MM-dd} al {fechaFin.Value:yyyy-MM-dd}";
        //             }
        //             worksheet.Cell(1, 1).Value = titulo;
        //             worksheet.Range("A1:I1").Merge().Style.Font.SetBold().Font.FontSize = 14;
        //             worksheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        //             worksheet.Cell(2, 1).Value = "ID";
        //             worksheet.Cell(2, 2).Value = "Fecha";
        //             worksheet.Cell(2, 3).Value = "Insumo";
        //             worksheet.Cell(2, 4).Value = "Descripción Insumo";
        //             worksheet.Cell(2, 5).Value = "Cantidad";
        //             worksheet.Cell(2, 6).Value = "ID Almacén";
        //             worksheet.Cell(2, 7).Value = "Estatus";
        //             worksheet.Cell(2, 8).Value = "Fecha Registro";
        //             worksheet.Cell(2, 9).Value = "Usuario Registra";

        //             for (int col = 1; col <= 9; col++)
        //             {
        //                 worksheet.Cell(2, col).Style.Font.SetBold();
        //                 worksheet.Cell(2, col).Style.Fill.BackgroundColor = XLColor.BabyBlue;
        //                 worksheet.Cell(2, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        //             }

        //             for (int i = 0; i < datos.Count; i++)
        //             {
        //                 worksheet.Cell(i + 3, 1).Value = datos[i].Id;
        //                 worksheet.Cell(i + 3, 2).Value = datos[i].Fecha;
        //                 worksheet.Cell(i + 3, 3).Value = datos[i].Insumo;
        //                 worksheet.Cell(i + 3, 4).Value = datos[i].DescripcionInsumo;
        //                 worksheet.Cell(i + 3, 5).Value = datos[i].Cantidad;
        //                 worksheet.Cell(i + 3, 6).Value = datos[i].IdAlmacen;
        //                 worksheet.Cell(i + 3, 7).Value = datos[i].Estatus;
        //                 worksheet.Cell(i + 3, 8).Value = datos[i].FechaRegistro;
        //                 worksheet.Cell(i + 3, 9).Value = datos[i].UsuarioRegistra;
        //             }

        //             worksheet.Columns().AdjustToContents();

        //             using (var stream = new MemoryStream())
        //             {
        //                 workbook.SaveAs(stream);
        //                 var content = stream.ToArray();

        //                 return File(
        //                     content,
        //                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        //                     "ReporteExistencias.xlsx");
        //             }
        //         }
        //     }
        //     catch (Exception ex)
        //     {
        //         return StatusCode((int)HttpStatusCode.InternalServerError, new { success = false, message = ex.Message });
        //     }
        // }



        [HttpPut("UpdateExistencias")]
        public IActionResult UpdateExistencias([FromBody] UpdateExistenciasModel req )
        {
            var objectResponse = Helper.GetStructResponse();
            try
            {
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = _ExistenciasService.UpdateExistencias(req);

                ;

            }

            catch (System.Exception ex)
            {
                objectResponse.message = ex.Message;
            }

            return new JsonResult(objectResponse);
        }

        [HttpDelete("DeleteExistencias/{id}")]
        public IActionResult DeleteExistencias([FromRoute] int id )
        {
            var objectResponse = Helper.GetStructResponse();
            try
            {
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = "data cargado con exito";

                _ExistenciasService.DeleteExistencias(id);

            }

            catch (System.Exception ex)
            {
                objectResponse.message = ex.Message;
            }

            return new JsonResult(objectResponse);
        }
    }
}