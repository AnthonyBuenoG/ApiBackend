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
    public class AlmacenController: ControllerBase
    {
   
        private readonly AlmacenService _AlmacenService;
        private readonly ILogger<AlmacenController> _logger;
  
        private readonly IJwtAuthenticationService _authService;
        private readonly IWebHostEnvironment _hostingEnvironment;
        

        Encrypt enc = new Encrypt();

        public AlmacenController(AlmacenService AlmacenService, ILogger<AlmacenController> logger, IJwtAuthenticationService authService) {
            _AlmacenService = AlmacenService;
            _logger = logger;
       
            _authService = authService;
            // Configura la ruta base donde se almacenan los archivos.
            // Asegúrate de ajustar la ruta según tu estructura de directorios.

            
            
        }


        [HttpPost("InsertAlmacen")]
        public IActionResult InsertAlmacen([FromBody] InsertAlmacenModel req )
        {
            var objectResponse = Helper.GetStructResponse();
            try
            {
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = _AlmacenService.InsertAlmacen(req);

            }

            catch (System.Exception ex)
            {
                objectResponse.message = ex.Message;
            }

            return new JsonResult(objectResponse);
        }

        [HttpGet("GetAlmacenExcel")]
        public IActionResult GetAlmacen(bool download = false)
        {
            var objectResponse = Helper.GetStructResponse();

            try
            {
                // Obtener los datos del servicio de almacenes
                var almacenes = _AlmacenService.GetAlmacen();

                // Validar si la lista es null o tiene un conteo de 0
                if (almacenes == null || almacenes.Count == 0)
                {
                    objectResponse.StatusCode = (int)HttpStatusCode.NoContent;
                    objectResponse.message = "No hay datos disponibles para exportar.";
                    return new JsonResult(objectResponse);
                }

                if (download) // Si se solicita la descarga del archivo Excel
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Almacenes");

                        // Encabezados para las columnas de datos
                        var headers = new string[] { "ID Almacén", "Nombre", "Dirección", "Estatus", "Fecha Registro", "Usuario Registra" };
                        for (int i = 0; i < headers.Length; i++)
                        {
                            worksheet.Cell(1, i + 1).Value = headers[i];
                            worksheet.Cell(1, i + 1).Style.Font.Bold = true;
                            worksheet.Cell(1, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        }

                        // Llenar los datos en el Excel
                        int currentRow = 2;
                        foreach (var almacen in almacenes)
                        {
                            worksheet.Cell(currentRow, 1).Value = almacen.IdAlmacen;
                            worksheet.Cell(currentRow, 2).Value = almacen.Nombre;
                            worksheet.Cell(currentRow, 3).Value = almacen.Direccion;
                            worksheet.Cell(currentRow, 4).Value = almacen.Estatus;
                            worksheet.Cell(currentRow, 5).Value = almacen.FechaRegistro;
                            worksheet.Cell(currentRow, 6).Value = almacen.UsuarioRegistra;
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
                            return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Almacenes.xlsx");
                        }
                    }
                }
                else // Si no se solicita la descarga, devolver los datos en formato JSON
                {
                    objectResponse.StatusCode = (int)HttpStatusCode.OK;
                    objectResponse.success = true;
                    objectResponse.message = "Datos cargados con éxito.";
                    objectResponse.response = almacenes;
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

        [HttpPut("UpdateAlmacen")]
        public IActionResult UpdateAlmacen([FromBody] UpdateAlmacenModel req )
        {
            var objectResponse = Helper.GetStructResponse();
            try
            {
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = _AlmacenService.UpdateAlmacen(req);

                ;

            }

            catch (System.Exception ex)
            {
                objectResponse.message = ex.Message;
            }

            return new JsonResult(objectResponse);
        }

        [HttpDelete("DeleteAlmacen/{id}")]
        public IActionResult DeleteAlmacen([FromRoute] int id )
        {
            var objectResponse = Helper.GetStructResponse();
            try
            {
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = "data cargado con exito";

                _AlmacenService.DeleteAlmacen(id);

            }

            catch (System.Exception ex)
            {
                objectResponse.message = ex.Message;
            }

            return new JsonResult(objectResponse);
        }
    }
}