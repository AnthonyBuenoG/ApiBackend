using System;
using Microsoft.AspNetCore.Mvc;
using reportesApi.Services;
using reportesApi.Models;
using reportesApi.Helpers;
using System.Net;
using Newtonsoft.Json;
using reportesApi.Models.Compras;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using Microsoft.AspNetCore.Authorization;
using ClosedXML.Excel;
using System.IO;
using Microsoft.AspNetCore.Mvc.Routing;

namespace reportesApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class TRSPController : ControllerBase
    {
        private readonly TRSPService _TRSPService;
        private readonly ILogger<TRSPController> _logger;

        public TRSPController(TRSPService TRSPService, ILogger<TRSPController> logger)
        {
            _TRSPService = TRSPService;
            _logger = logger;
        }

            // Endpoint para obtener transferencias con filtros opcionales
        [HttpPost("InsertTRSP")]
        public IActionResult InsertTRSP([FromBody] InsertTRSPModel req)
        {
            var objectResponse = Helper.GetStructResponse();

            try
            {
                var folioGenerado = _TRSPService.InsertTRSPTransferencia(req);
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = "Transferencia registrada con éxito";
                objectResponse.response = new { FolioGenerado = folioGenerado };
            }
            catch (System.Exception ex)
            {
                objectResponse.message = ex.Message;
            }

            return new JsonResult(objectResponse);
        }

         [HttpPost("InsertOrgTRSP")]
        public IActionResult InsertAlmacenOrgTRSP([FromBody] InsertTRSPRenglonModel req)
        {
            var objectResponse = Helper.GetStructResponse();

            try
            {
                var folioGenerado = _TRSPService.InsertTRSPRenglonTransferencia(req);
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = "Transferencia registrada con éxito";
                objectResponse.response = new { FolioGenerado = folioGenerado };
            }
            catch (System.Exception ex)
            {
                objectResponse.message = ex.Message;
            }

            return new JsonResult(objectResponse);
        }



        [HttpGet("GetTRSP")]
        public IActionResult GetTRSPTransferencias(int? almacen = null, DateTime? fechaInicio = null, DateTime? fechaFin = null, int? tipoMovimiento = null)
        {
            var objectResponse = Helper.GetStructResponse();
            try
            {
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = "Transferencias obtenidas con éxito";
                objectResponse.response = _TRSPService.GetTRSPTransferencias( almacen, fechaInicio, fechaFin, tipoMovimiento);
            }
            catch (Exception ex)
            {
                objectResponse.message = ex.Message;
            }
            return new JsonResult(objectResponse);
        }

          [HttpGet("GetTRSPOrgFechas")]
        public IActionResult GetTRSPOrgTransferencias(int? almacen = null, DateTime? fechaInicio = null, DateTime? fechaFin = null, int? tipoMovimiento = null)
        {
            var objectResponse = Helper.GetStructResponse();
            try
            {
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = "Transferencias obtenidas con éxito";
                objectResponse.response = _TRSPService.GetTRSPOrgTransferencias( almacen, fechaInicio, fechaFin, tipoMovimiento);
            }
            catch (Exception ex)
            {
                objectResponse.message = ex.Message;
            }
            return new JsonResult(objectResponse);
        }


         [HttpGet("GetEntradaSalidaTRSP")]
        public IActionResult GetEntradaSalidaTRSP([FromQuery] int tipoMovimiento)
        {
            var objectResponse = Helper.GetStructResponse();

            try
            {
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = "Datos cargados exitosamente";
                var resultado = _TRSPService.GetEntradaSalidaTRSP(tipoMovimiento);
               
               

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

         [HttpGet("GetOrgTRSP")]
        public IActionResult GetRenglonTRSP([FromQuery] int almacenOrigen)
        {
            var objectResponse = Helper.GetStructResponse();

            try
            {
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = "Datos cargados exitosamente";
                var resultado = _TRSPService.GetTRSPRenglon(almacenOrigen);
               
               

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


        [HttpGet("GetReporteTRSP")]
        public IActionResult GetReporteTRSP(string? folio = null, bool download = false)
        {
            var objectResponse = Helper.GetStructResponse();

            try
            {
                var reportes = _TRSPService.GetReporteTRSPTransferencias(folio);

                if (reportes == null || reportes.Count == 0)
                {
                    objectResponse.StatusCode = (int)HttpStatusCode.NoContent;
                    objectResponse.message = "No hay datos disponibles para exportar.";
                    return new JsonResult(objectResponse);
                }

                if (download) 
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Reporte TRSP");

                        worksheet.Cell(1, 1).Value = $"Reporte de Transferencias (Folio: {folio ?? "N/A"})";
                        worksheet.Cell(1, 1).Style.Font.Bold = true;
                        worksheet.Cell(1, 1).Style.Font.FontSize = 14;
                        worksheet.Range(1, 1, 1, 18).Merge(); 
                        worksheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                        var headers = new string[]
                        {
                            "ID TRSP", "Almacén Origen", "Nombre Almacén Origen", "Almacén Destino", "Nombre Almacén Destino",
                            "ID Insumo", "Descripción Insumo", "Fecha Entrada", "Fecha Salida", "Cantidad",
                            "Tipo Movimiento", "Descripción", "No. Folio", "Cantidad Origen", "Cantidad Destino",
                            "Fecha Registro", "Estatus", "Usuario Registra"
                        };

                        for (int i = 0; i < headers.Length; i++)
                        {
                            worksheet.Cell(2, i + 1).Value = headers[i];
                            worksheet.Cell(2, i + 1).Style.Font.Bold = true;
                            worksheet.Cell(2, i + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                            worksheet.Cell(2, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(2, i + 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        }

                        int currentRow = 3;
                        foreach (var reporte in reportes)
                        {
                            worksheet.Cell(currentRow, 1).Value = reporte.IdTRSP;
                            worksheet.Cell(currentRow, 2).Value = reporte.AlmacenOrigen;
                            worksheet.Cell(currentRow, 3).Value = reporte.NombreAlmacenOrgien;
                            worksheet.Cell(currentRow, 4).Value = reporte.AlmacenDestino;
                            worksheet.Cell(currentRow, 5).Value = reporte.NombreAlmacenDestino;
                            worksheet.Cell(currentRow, 6).Value = reporte.IdInsumo;
                            worksheet.Cell(currentRow, 7).Value = reporte.DescripcionInsumo;
                            worksheet.Cell(currentRow, 8).Value = reporte.FechaEntrada;
                            worksheet.Cell(currentRow, 9).Value = reporte.FechaSalida;
                            worksheet.Cell(currentRow, 10).Value = reporte.Cantidad;
                            worksheet.Cell(currentRow, 11).Value = reporte.TipoMovimiento;
                            worksheet.Cell(currentRow, 12).Value = reporte.Descripcion;
                            worksheet.Cell(currentRow, 13).Value = reporte.NoFolio;
                            worksheet.Cell(currentRow, 14).Value = reporte.CantidadMovimientoOrigen;
                            worksheet.Cell(currentRow, 15).Value = reporte.CantidadMovimientoDestino;
                            worksheet.Cell(currentRow, 16).Value = reporte.FechaRegistro;
                            worksheet.Cell(currentRow, 17).Value = reporte.Estatus;
                            worksheet.Cell(currentRow, 18).Value = reporte.UsuarioRegistra;

                            for (int col = 1; col <= headers.Length; col++)
                            {
                                worksheet.Cell(currentRow, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            }
                            currentRow++;
                        }

                        worksheet.Columns().AdjustToContents();

                        using (var stream = new MemoryStream())
                        {
                            workbook.SaveAs(stream);
                            stream.Seek(0, SeekOrigin.Begin); 
                            var content = stream.ToArray();

                            return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"ReporteTraspasos.xlsx");
                        }
                    }
                }
                else
                {
                    objectResponse.StatusCode = (int)HttpStatusCode.OK;
                    objectResponse.success = true;
                    objectResponse.message = "Datos cargados con éxito.";
                    objectResponse.response = reportes;
                    return new JsonResult(objectResponse);
                }
            }
            catch (Exception ex)
            {
                objectResponse.StatusCode = (int)HttpStatusCode.InternalServerError;
                objectResponse.message = $"Error: {ex.Message}";
                return new JsonResult(objectResponse);
            }
        }

    //   [HttpPut("UpdateTRSP")]
    //     public IActionResult UpdateTRSPTransferncia([FromBody] UpdateTRSPModel req )
    //     {
    //         var objectResponse = Helper.GetStructResponse();
    //         try
    //         {
    //             objectResponse.StatusCode = (int)HttpStatusCode.OK;
    //             objectResponse.success = true;
    //             objectResponse.message = _TRSPService.UpdateTRSPTransferncia(req);

    //             ;

    //         }

    //         catch (System.Exception ex)
    //         {
    //             objectResponse.message = ex.Message;
    //         }

    //         return new JsonResult(objectResponse);
    //     }


        [HttpDelete("InactivarTRSP/{id}")]
        public IActionResult DeleteTRSPTransferencias([FromRoute] int id )
        {
            var objectResponse = Helper.GetStructResponse();
            try
            {
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = "data cargado con exito";

                _TRSPService.DeleteTRSPTransferencia(id);

            }

            catch (System.Exception ex)
            {
                objectResponse.message = ex.Message;
            }

            return new JsonResult(objectResponse);
        }


    } 
}

