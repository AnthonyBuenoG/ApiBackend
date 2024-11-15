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

    [HttpGet("GetTRSPTransferencias")]
    public IActionResult GetTRSPTransferencias(int? almacenOrigen = null, int? almacenDestino = null, DateTime? fechaInicio = null, DateTime? fechaFin = null, int? tipoMovimiento = null)
    {
        var objectResponse = Helper.GetStructResponse();
        try
        {
            objectResponse.StatusCode = (int)HttpStatusCode.OK;
            objectResponse.success = true;
            objectResponse.message = "Transferencias obtenidas con éxito";
            objectResponse.response = _TRSPService.GetTRSPTransferencias(almacenOrigen, almacenDestino, fechaInicio, fechaFin, tipoMovimiento);
        }
        catch (Exception ex)
        {
            objectResponse.message = ex.Message;
        }
        return new JsonResult(objectResponse);
    }
}

        // Endpoint para eliminar una transferencia
        // [HttpDelete("DeleteTransferencia/{id}")]
        // public IActionResult DeleteTransferencia([FromRoute] int id)
        // {
        //     var objectResponse = Helper.GetStructResponse();
        //     try
        //     {
        //         // Llamar al servicio para eliminar la transferencia
        //         _TRSPService.DeleteTRSP(id);

        //         objectResponse.StatusCode = (int)HttpStatusCode.OK;
        //         objectResponse.success = true;
        //         objectResponse.message = "Transferencia eliminada con éxito";

        //     }
        //     catch (System.Exception ex)
        //     {
        //         objectResponse.message = ex.Message;
        //     }

        //     return new JsonResult(objectResponse);
        // }
    }

