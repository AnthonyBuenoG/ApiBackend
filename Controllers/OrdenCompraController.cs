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

namespace reportesApi.Controllers
{
   
    [Route("api")]
    public class OrdenCompraController: ControllerBase
    {
   
        private readonly OrdenCompraService _OrdenCompraService;
        private readonly ILogger<OrdenCompraController> _logger;
  
        private readonly IJwtAuthenticationService _authService;
        private readonly IWebHostEnvironment _hostingEnvironment;
        

        Encrypt enc = new Encrypt();

        public OrdenCompraController(OrdenCompraService OrdenCompraService, ILogger<OrdenCompraController> logger, IJwtAuthenticationService authService) {
            _OrdenCompraService = OrdenCompraService;
            _logger = logger;
       
            _authService = authService;
            // Configura la ruta base donde se almacenan los archivos.
            // Asegúrate de ajustar la ruta según tu estructura de directorios.

            
            
        }


        [HttpPost("Insert OrdenCompra")]
        public IActionResult InsertOrdenCompra([FromBody] InsertOrdenCompraModel req )
        {
            var objectResponse = Helper.GetStructResponse();
            try
            {
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = _OrdenCompraService.InsertOrdenCompra(req);

            }

            catch (System.Exception ex)
            {
                objectResponse.message = ex.Message;
            }

            return new JsonResult(objectResponse);
        }

        [HttpGet("Get OrdenCompra")]
        public IActionResult GetOrdenCompra()
        {
            var objectResponse = Helper.GetStructResponse();
            var resultado = _OrdenCompraService.GetOrdenCompra();

            try
            {
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = "data cargado con exito";
                 objectResponse.response = resultado;
            }

            catch (System.Exception ex)
            {
                objectResponse.message = ex.Message;
            }

            return new JsonResult(objectResponse);
        }

        [HttpPut("Update OrdenCompra")]
        public IActionResult UpdateOrdenCompra([FromBody] UpdateOrdenCompraModel req )
        {
            var objectResponse = Helper.GetStructResponse();
            try
            {
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = _OrdenCompraService.UpdateOrdenCompra(req);

                ;

            }

            catch (System.Exception ex)
            {
                objectResponse.message = ex.Message;
            }

            return new JsonResult(objectResponse);
        }

        [HttpDelete("Delete OrdenCompra/{id}")]
        public IActionResult DeleteOrdenCompra([FromRoute] int id )
        {
            var objectResponse = Helper.GetStructResponse();
            try
            {
                objectResponse.StatusCode = (int)HttpStatusCode.OK;
                objectResponse.success = true;
                objectResponse.message = "data cargado con exito";

                _OrdenCompraService.DeleteOrdenCompra(id);

            }

            catch (System.Exception ex)
            {
                objectResponse.message = ex.Message;
            }

            return new JsonResult(objectResponse);
        }
    }
}