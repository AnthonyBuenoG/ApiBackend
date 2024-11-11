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
    public class RenglonMovimientosController : ControllerBase
    {
   
        private readonly RenglonMovimientoService _RenglonMovimientos;
        private readonly ILogger<RenglonMovimientosController> _logger;
  
        private readonly IJwtAuthenticationService _authService;
        private readonly IWebHostEnvironment _hostingEnvironment;
        

        Encrypt enc = new Encrypt();

        public RenglonMovimientosController(RenglonMovimientoService RenglonMovimientoService, ILogger<RenglonMovimientosController> logger, IJwtAuthenticationService authService) {
            _RenglonMovimientos = RenglonMovimientoService;
            _logger = logger;
       
            _authService = authService;
            // Configura la ruta base donde se almacenan los archivos.
            // Asegúrate de ajustar la ruta según tu estructura de directorios.

            
            
        }


        [HttpPost("InsertRenglonMovimientos")]
        public IActionResult InsertRenglonMovimientos([FromBody] InsertRenglonMovimientosModel req )
        {
            var objectResponse = Helper.GetStructResponse();
            try
            {
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = _RenglonMovimientos.InsertRenglonMovimientos(req);

            }

            catch (System.Exception ex)
            {
                objectResponse.message = ex.Message;
            }

            return new JsonResult(objectResponse);
        }

        // [HttpGet("GetRenglonMovimientos")]
        // public IActionResult GetRenglonMovimientos([FromQuery] int IdMovimiento)
        // {
        //     var objectResponse = Helper.GetStructResponse();

        //     try
        //     {
        //         objectResponse.StatusCode = (int)HttpStatusCode.OK;
        //         objectResponse.success = true;
        //         objectResponse.message = "DetalleEntradas cargados con exito";
        //         var resultado = _RenglonMovimientos.GetRenglonMovimimentos(IdMovimiento);
               
               

        //         // Llamando a la función y recibiendo los dos valores.
               
        //          objectResponse.response = resultado;
        //     }

        //     catch (System.Exception ex)
        //     {
        //         objectResponse.StatusCode = (int)HttpStatusCode.InternalServerError;
        //         objectResponse.success = false;
        //         objectResponse.message = ex.Message;
        //     }

        //     return new JsonResult(objectResponse);
        // }

        [HttpGet("ExcelRenglonesMovimientos")]
        public IActionResult ExportarRenglonMovimientoExcel(
            DateTime? fechaInicio = null,
            DateTime? fechaFin = null,
            bool filtrarPorFecha = false,
            bool filtrarPorIdAlmacen = false)
        {
            try
            {
                if (filtrarPorFecha && (!fechaInicio.HasValue || !fechaFin.HasValue))
                {
                    return BadRequest(new { success = false, message = "Debe proporcionar tanto fechaInicio como fechaFin para el filtrado por fechas." });
                }

                var datos = _RenglonMovimientos.GetRenglonMovimimentos(
                    filtrarPorFecha ? fechaInicio : null,
                    filtrarPorFecha ? fechaFin : null);

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("RenglonMovimiento");

                    string titulo = "Reporte de Renglon Movimiento";
                    if (filtrarPorFecha && fechaInicio.HasValue && fechaFin.HasValue)
                    {
                        titulo += $" del {fechaInicio.Value:yyyy-MM-dd} al {fechaFin.Value:yyyy-MM-dd}";
                    }
                    worksheet.Cell(1, 1).Value = titulo;
                    worksheet.Range("A1:I1").Merge().Style.Font.SetBold().Font.FontSize = 14;
                    worksheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    worksheet.Cell(2, 1).Value = "ID";
                    worksheet.Cell(2, 2).Value = "IdMovimiento";
                    worksheet.Cell(2, 3).Value = "Nombre";
                    worksheet.Cell(2, 4).Value = "Insumo";
                    worksheet.Cell(2, 5).Value = "Descripción Insumo";
                    worksheet.Cell(2, 6).Value = "Cantidad";
                    worksheet.Cell(2, 7).Value = "Costo";
                    worksheet.Cell(2, 8).Value = "CostoTotal";
                    worksheet.Cell(2, 9).Value = "Estatus";
                    worksheet.Cell(2, 10).Value = "Fecha Registro";
                    worksheet.Cell(2, 11).Value = "Usuario Registra";

                    for (int col = 1; col <= 9; col++)
                    {
                        worksheet.Cell(2, col).Style.Font.SetBold();
                        worksheet.Cell(2, col).Style.Fill.BackgroundColor = XLColor.BabyBlue;
                        worksheet.Cell(2, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    }

                    for (int i = 0; i < datos.Count; i++)
                    {
                        worksheet.Cell(i + 3, 1).Value = datos[i].Id;
                        worksheet.Cell(i + 3, 2).Value = datos[i].IdMovimiento;
                        worksheet.Cell(i + 3, 2).Value = datos[i].Nombre;
                        worksheet.Cell(i + 3, 4).Value = datos[i].Insumo;
                        worksheet.Cell(i + 3, 5).Value = datos[i].DescripcionInsumo;
                        worksheet.Cell(i + 3, 6).Value = datos[i].Cantidad;
                        worksheet.Cell(i + 3, 7).Value = datos[i].Costo;
                        worksheet.Cell(i + 3, 8).Value = datos[i].CostoTotal;
                        worksheet.Cell(i + 3, 9).Value = datos[i].Estatus;
                        worksheet.Cell(i + 3, 10).Value = datos[i].FechaRegistro;
                        worksheet.Cell(i + 3, 11).Value = datos[i].UsuarioRegistra;
                    }

                    worksheet.Columns().AdjustToContents();

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        var content = stream.ToArray();

                        return File(
                            content,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            "ReporteExistencias.xlsx");
                    }
                }
            }
            catch (Exception ex)
            {
                return StatusCode((int)HttpStatusCode.InternalServerError, new { success = false, message = ex.Message });
            }
        }


        [HttpPut("UpdateRenglonMovimientos")]
        public IActionResult UpdateRenglonMovimientos([FromBody] UpdateRenglonMovimientosModel req )
        {
            var objectResponse = Helper.GetStructResponse();
            try
            {
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = _RenglonMovimientos.UpdateRenglonMovimientos(req);

                ;

            }

            catch (System.Exception ex)
            {
                objectResponse.message = ex.Message;
            }

            return new JsonResult(objectResponse);
        }

        [HttpDelete("DeleteRenglonMovimientos/{id}")]
        public IActionResult DeleteRenglonMovimientos([FromRoute] int id )
        {
            var objectResponse = Helper.GetStructResponse();
            try
            {
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = "data cargado con exito";

                _RenglonMovimientos.DeleteRenglonMovimientos(id);

            }

            catch (System.Exception ex)
            {
                objectResponse.message = ex.Message;
            }

            return new JsonResult(objectResponse);
        }
    }
}