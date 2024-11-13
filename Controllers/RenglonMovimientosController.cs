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

[HttpGet("GetRenglonMovimientos")]
public IActionResult GetRenglonMovimientos(DateTime? FechaInicio = null, DateTime? FechaFin = null, bool download = false)
{
    var objectResponse = Helper.GetStructResponse();

    try
    {
        // Llama al método que ejecuta el stored procedure con los parámetros de fechas
        var movimientos = _RenglonMovimientos.GetRenglonMovimimentos(FechaInicio, FechaFin);

        if (movimientos == null || movimientos.Count == 0)
        {
            objectResponse.StatusCode = (int)HttpStatusCode.NoContent;
            objectResponse.message = "No hay datos disponibles para exportar.";
            return new JsonResult(objectResponse);
        }

        if (download) // Si se solicita la descarga del archivo Excel
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Movimientos");

                // Encabezado personalizado con rango de fechas
                var fechaInicioText = FechaInicio.HasValue ? FechaInicio.Value.ToString("yyyy-MM-dd") : "Inicio";
                var fechaFinText = FechaFin.HasValue ? FechaFin.Value.ToString("yyyy-MM-dd") : "Fin";
                worksheet.Cell(1, 1).Value = $"Reporte de movimientos {fechaInicioText} al {fechaFinText}";
                worksheet.Range(1, 1, 1, 11).Merge().Style
                    .Font.SetBold(true)
                    .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                // Encabezados de las columnas
                var headers = new string[] { "ID", "ID Movimiento", "TipoMovimiento", "Insumo", "Descripción", "Cantidad", "Costo", "Total Renglón", "Estatus", "Fecha de Registro", "Usuario Registra" };
                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cell(2, i + 1).Value = headers[i];
                    worksheet.Cell(2, i + 1).Style.Font.Bold = true;
                    worksheet.Cell(2, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }

                // Rellenar el Excel con los datos de movimientos
                int currentRow = 3;
                foreach (var movimiento in movimientos)
                {
                    worksheet.Cell(currentRow, 1).Value = movimiento.Id;
                    worksheet.Cell(currentRow, 2).Value = movimiento.IdMovimiento;
                    worksheet.Cell(currentRow, 3).Value = movimiento.Nombre;
                    worksheet.Cell(currentRow, 4).Value = movimiento.Insumo;
                    worksheet.Cell(currentRow, 5).Value = movimiento.DescripcionInsumo;
                    worksheet.Cell(currentRow, 6).Value = movimiento.Cantidad;
                    worksheet.Cell(currentRow, 7).Value = movimiento.Costo;
                    worksheet.Cell(currentRow, 8).Value = movimiento.CostoTotal;
                    worksheet.Cell(currentRow, 9).Value = movimiento.Estatus;
                    worksheet.Cell(currentRow, 10).Value = movimiento.FechaRegistro;
                    worksheet.Cell(currentRow, 11).Value = movimiento.UsuarioRegistra;
                    currentRow++;
                }

                // Ajustar automáticamente el tamaño de las columnas
                worksheet.Columns().AdjustToContents();

                // Guardar el archivo en un MemoryStream
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    // Devolver el archivo Excel como un archivo descargable
                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Movimientos.xlsx");
                }
            }
        }
        else // Si no se solicita la descarga, devolver los datos en formato JSON
        {
            objectResponse.StatusCode = (int)HttpStatusCode.OK;
            objectResponse.success = true;
            objectResponse.message = "Datos cargados con éxito.";
            objectResponse.response = movimientos;
            return new JsonResult(objectResponse);
        }
    }
    catch (Exception ex)
    {
        // Manejo de errores
        objectResponse.StatusCode = (int)HttpStatusCode.InternalServerError;
        objectResponse.message = $"Error: {ex.Message}";
        return new JsonResult(objectResponse);
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