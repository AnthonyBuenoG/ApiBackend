using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using reportesApi.DataContext;
using reportesApi.Models;
using reportesApi.Models.Compras;
using OfficeOpenXml;
using Microsoft.AspNetCore.Hosting;
using System.IO;
using Microsoft.AspNetCore.Mvc;
using System.Linq;
using System.Collections.Generic;

namespace reportesApi.Services
{
    public class TRSPService
    {
        private string connection;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private ArrayList parametros = new ArrayList();

        public TRSPService(IMarcatelDatabaseSetting settings, IWebHostEnvironment webHostEnvironment)
        {
            connection = settings.ConnectionString;
            _webHostEnvironment = webHostEnvironment;
        }
    public List<GetTRSPModel> GetTRSPTransferencias(int? almacenOrigen = null, int? almacenDestino = null, DateTime? fechaInicio = null, DateTime? fechaFin = null, int? tipoMovimiento = null)
    {
        ConexionDataAccess dac = new ConexionDataAccess(connection);
        List<GetTRSPModel> lista = new List<GetTRSPModel>();
        parametros = new ArrayList
        {
            new SqlParameter { ParameterName = "@AlmacenOrigen", SqlDbType = SqlDbType.Int, Value = (object)almacenOrigen ?? DBNull.Value },
            new SqlParameter { ParameterName = "@AlmacenDestino", SqlDbType = SqlDbType.Int, Value = (object)almacenDestino ?? DBNull.Value },
            new SqlParameter { ParameterName = "@FechaInicio", SqlDbType = SqlDbType.DateTime, Value = (object)fechaInicio ?? DBNull.Value },
            new SqlParameter { ParameterName = "@FechaFin", SqlDbType = SqlDbType.DateTime, Value = (object)fechaFin ?? DBNull.Value },
            new SqlParameter { ParameterName = "@TipoMovimiento", SqlDbType = SqlDbType.Int, Value = (object)tipoMovimiento ?? DBNull.Value }
        };

        try
        {
            DataSet ds = dac.Fill("sp_get_TRSP_Transferencias", parametros);
            if (ds.Tables[0].Rows.Count > 0)
            {
                lista = ds.Tables[0].AsEnumerable().Select(dataRow => new GetTRSPModel
                {
                    IdTRSP = dataRow["IdTRSP"].ToString(),
                    AlmacenOrigen = dataRow["AlmacenOrigen"].ToString(),
                    NombreAlmacenOrgien = dataRow["NombreAlmacenOrgien"].ToString(),
                    AlmacenDestino = dataRow["AlmacenDestino"].ToString(),
                    NombreAlmacenDestino = dataRow["NombreAlmacenDestino"].ToString(),
                    IdInsumo = dataRow["IdInsumo"].ToString(),
                    DescripcionInsumo = dataRow["DescripcionInsumo"].ToString(),
                    FechaEntrada = dataRow["FechaEntrada"].ToString(),
                    FechaSalida = dataRow["FechaSalida"].ToString(),
                    Cantidad = dataRow["Cantidad"].ToString(),
                    TipoMovimiento = dataRow["TipoMovimiento"].ToString(),
                    Descripcion = dataRow["Descripcion"].ToString(),
                    NoFolio = dataRow["No_Folio"].ToString(),
                    CantidadMovimientoOrigen = dataRow["CantidadMovimientoOrigen"].ToString(),
                    CantidadMovimientoDestino = dataRow["CantidadMovimientoDestino"].ToString(),
                    FechaRegistro = dataRow["FechaRegistro"].ToString(),
                    Estatus = dataRow["Estatus"].ToString(),
                    UsuarioRegistra = dataRow["UsuarioRegistra"].ToString()
                }).ToList();
            }
        }
        catch (Exception ex)
        {
            throw ex;
        }

        return lista;
    }


    private string GenerarNoFolio()
    {
        var random = new Random();
        const string letras = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        string letrasAleatorias = new string(Enumerable.Repeat(letras, 3).Select(s => s[random.Next(s.Length)]).ToArray());
        string fechaHora = DateTime.Now.ToString("yyyyMMddHHmmss");
        
        return $"{letrasAleatorias}{fechaHora}";
    }

    public string InsertTRSPTransferencia(InsertTRSPModel trsp)
    {
        ConexionDataAccess dac = new ConexionDataAccess(connection);
        parametros = new ArrayList();

        string folioGenerado = GenerarNoFolio(); // Generar NoFolio automático

        parametros.Add(new SqlParameter { ParameterName = "@AlmacenOrigen", SqlDbType = SqlDbType.Int, Value = trsp.AlmacenOrigen });
        parametros.Add(new SqlParameter { ParameterName = "@AlmacenDestino", SqlDbType = SqlDbType.Int, Value = trsp.AlmacenDestino });
        parametros.Add(new SqlParameter { ParameterName = "@IdInsumo", SqlDbType = SqlDbType.Int, Value = trsp.IdInsumo });
        parametros.Add(new SqlParameter { ParameterName = "@FechaEntrada", SqlDbType = SqlDbType.DateTime, Value = trsp.FechaEntrada });
        parametros.Add(new SqlParameter { ParameterName = "@FechaSalida", SqlDbType = SqlDbType.DateTime, Value = trsp.FechaSalida });
        parametros.Add(new SqlParameter { ParameterName = "@Cantidad", SqlDbType = SqlDbType.Int, Value = trsp.Cantidad });
        parametros.Add(new SqlParameter { ParameterName = "@TipoMovimiento", SqlDbType = SqlDbType.Int, Value = trsp.TipoMovimiento });
        parametros.Add(new SqlParameter { ParameterName = "@Descripcion", SqlDbType = SqlDbType.VarChar, Value = trsp.Descripcion });
        parametros.Add(new SqlParameter { ParameterName = "@No_Folio", SqlDbType = SqlDbType.VarChar, Value = folioGenerado });
        parametros.Add(new SqlParameter { ParameterName = "@UsuarioRegistra", SqlDbType = SqlDbType.Int, Value = trsp.UsuarioRegistra });

        try
        {
            DataSet ds = dac.Fill("sp_insert_TRSP_Transferencia", parametros);
            folioGenerado = ds.Tables[0].Rows[0]["FolioGenerado"].ToString();

        }
        catch (Exception ex)
        {
            throw ex;
        }

        return folioGenerado;
    }

        // // Método para actualizar una transferencia existente
        // public string UpdateTRSP(UpdateTRSPModel trsp)
        // {
        //     ConexionDataAccess dac = new ConexionDataAccess(connection);
        //     parametros = new ArrayList();

        //     parametros.Add(new SqlParameter("@IdTRSP", trsp.IdTRSP));
        //     parametros.Add(new SqlParameter("@AlmacenOrigen", trsp.AlmacenOrigen));
        //     parametros.Add(new SqlParameter("@AlmacenDestino", trsp.AlmacenDestino));
        //     parametros.Add(new SqlParameter("@IdInsumo", trsp.IdInsumo));
        //     parametros.Add(new SqlParameter("@FechaEntrada", trsp.FechaEntrada));
        //     parametros.Add(new SqlParameter("@FechaSalida", trsp.FechaSalida));
        //     parametros.Add(new SqlParameter("@Cantidad", trsp.Cantidad));
        //     parametros.Add(new SqlParameter("@TipoMovimiento", trsp.TipoMovimiento));
        //     parametros.Add(new SqlParameter("@Descripcion", trsp.Descripcion));
        //     parametros.Add(new SqlParameter("@Estatus", trsp.Estatus));
        //     parametros.Add(new SqlParameter("@UsuarioRegistra", trsp.UsuarioRegistra));

        //     try
        //     {
        //         dac.ExecuteNonQuery("sp_update_TRSP_Transferencia", parametros);
        //         return "Transferencia actualizada correctamente.";
        //     }
        //     catch (Exception ex)
        //     {
        //         throw ex;
        //     }
        // }
    }
}
