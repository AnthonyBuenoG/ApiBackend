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

        private string GenerarNoFolio()
        {
            var random = new Random();
            const string letras = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string letrasAleatorias = new string(Enumerable.Repeat(letras, 4).Select(s => s[random.Next(s.Length)]).ToArray());
            string fechaHora = DateTime.Now.ToString("yyyyMMddHHmmss");
            
            return $"{letrasAleatorias}{fechaHora}";
        }

    public List<GetTRSPModel> GetTRSPTransferencias(int? almacen = null, DateTime? fechaInicio = null, DateTime? fechaFin = null, int? tipoMovimiento = null)
    {
        ConexionDataAccess dac = new ConexionDataAccess(connection);
        List<GetTRSPModel> lista = new List<GetTRSPModel>();
        parametros = new ArrayList
        {
            new SqlParameter { ParameterName = "@Almacen", SqlDbType = SqlDbType.Int, Value = (object)almacen ?? DBNull.Value },
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

     public List<GetTRSPOrgModel> GetTRSPOrgTransferencias(int? almacen = null, DateTime? fechaInicio = null, DateTime? fechaFin = null, int? tipoMovimiento = null)
    {
        ConexionDataAccess dac = new ConexionDataAccess(connection);
        List<GetTRSPOrgModel> lista = new List<GetTRSPOrgModel>();
        parametros = new ArrayList
        {
            new SqlParameter { ParameterName = "@Almacen", SqlDbType = SqlDbType.Int, Value = (object)almacen ?? DBNull.Value },
            new SqlParameter { ParameterName = "@FechaInicio", SqlDbType = SqlDbType.DateTime, Value = (object)fechaInicio ?? DBNull.Value },
            new SqlParameter { ParameterName = "@FechaFin", SqlDbType = SqlDbType.DateTime, Value = (object)fechaFin ?? DBNull.Value },
            new SqlParameter { ParameterName = "@TipoMovimiento", SqlDbType = SqlDbType.Int, Value = (object)tipoMovimiento ?? DBNull.Value }
        };

        try
        {
            DataSet ds = dac.Fill("sp_get_TRSP_OrgFechas", parametros);
            if (ds.Tables[0].Rows.Count > 0)
            {
                lista = ds.Tables[0].AsEnumerable().Select(dataRow => new GetTRSPOrgModel
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

        public List<GetReporteTRSPModel> GetReporteTRSPTransferencias(string? folio = null)
    {
        ConexionDataAccess dac = new ConexionDataAccess(connection);
        List<GetReporteTRSPModel> lista = new List<GetReporteTRSPModel>();
        parametros = new ArrayList
        {
            new SqlParameter { ParameterName = "@Folio", SqlDbType = SqlDbType.VarChar, Value = (object)folio ?? DBNull.Value },
        
        };

        try
        {
            DataSet ds = dac.Fill("sp_get_reporte_TRSP", parametros);
            if (ds.Tables[0].Rows.Count > 0)
            {
                lista = ds.Tables[0].AsEnumerable().Select(dataRow => new GetReporteTRSPModel
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

    //         public List<GetEntradaSalidaTRSPPModel> GetEntradaSalidaTRSPTransferencias(string tipoMovimiento)
    // {
    //     ConexionDataAccess dac = new ConexionDataAccess(connection);
    //     List<GetEntradaSalidaTRSPPModel> lista = new List<GetEntradaSalidaTRSPPModel>();
    //     parametros = new ArrayList
    //     {
    //         parametros.Add(new SqlParameter { ParameterName = "@TipoMovimiento", SqlDbType = SqlDbType.VarChar, Value = tipoMovimiento }),
        
    //     };

    //     try
    //     {
    //         DataSet ds = dac.Fill("sp_get_TRSP_entradasalida", parametros);
    //         if (ds.Tables[0].Rows.Count > 0)
    //         {
    //             lista = ds.Tables[0].AsEnumerable().Select(dataRow => new GetEntradaSalidaTRSPPModel
    //             {
    //                 IdTRSP = dataRow["IdTRSP"].ToString(),
    //                 AlmacenOrigen = dataRow["AlmacenOrigen"].ToString(),
    //                 NombreAlmacenOrgien = dataRow["NombreAlmacenOrgien"].ToString(),
    //                 AlmacenDestino = dataRow["AlmacenDestino"].ToString(),
    //                 NombreAlmacenDestino = dataRow["NombreAlmacenDestino"].ToString(),
    //                 IdInsumo = dataRow["IdInsumo"].ToString(),
    //                 DescripcionInsumo = dataRow["DescripcionInsumo"].ToString(),
    //                 FechaEntrada = dataRow["FechaEntrada"].ToString(),
    //                 FechaSalida = dataRow["FechaSalida"].ToString(),
    //                 Cantidad = dataRow["Cantidad"].ToString(),
    //                 TipoMovimiento = dataRow["TipoMovimiento"].ToString(),
    //                 Descripcion = dataRow["Descripcion"].ToString(),
    //                 NoFolio = dataRow["No_Folio"].ToString(),
    //                 FechaRegistro = dataRow["FechaRegistro"].ToString(),
    //                 Estatus = dataRow["Estatus"].ToString(),
    //                 UsuarioRegistra = dataRow["UsuarioRegistra"].ToString()
    //             }).ToList();
    //         }
    //     }
    //     catch (Exception ex)
    //     {
    //         throw ex;
    //     }

    //     return lista;
    // }



         public List<GetEntradaSalidaTRSPPModel> GetEntradaSalidaTRSP(int tipoMovimiento)
        {

            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            parametros.Add(new SqlParameter { ParameterName = "@TipoMovimiento", SqlDbType = SqlDbType.Int, Value = tipoMovimiento });

            List<GetEntradaSalidaTRSPPModel> lista = new List<GetEntradaSalidaTRSPPModel>();
            try
            {
                DataSet ds = dac.Fill("sp_get_TRSP_entradasalida", parametros);
                if (ds.Tables[0].Rows.Count > 0)
                {

                  lista = ds.Tables[0].AsEnumerable()
                    .Select(dataRow => new GetEntradaSalidaTRSPPModel {
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
                    FechaRegistro = dataRow["FechaRegistro"].ToString(),
                    Estatus = dataRow["Estatus"].ToString(),
                    UsuarioRegistra = dataRow["UsuarioRegistra"].ToString()
                    }).ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw ex;
            }
            return lista;
        }

        
         public List<GetTRSPRenglonModel> GetTRSPRenglon(int almacenOrigen)
        {

            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            parametros.Add(new SqlParameter { ParameterName = "@Almacen", SqlDbType = SqlDbType.Int, Value = almacenOrigen });

            List<GetTRSPRenglonModel> lista = new List<GetTRSPRenglonModel>();
            try
            {
                DataSet ds = dac.Fill("sp_get_TRSP_RenglonTransferencias", parametros);
                if (ds.Tables[0].Rows.Count > 0)
                {

                  lista = ds.Tables[0].AsEnumerable()
                    .Select(dataRow => new GetTRSPRenglonModel {
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
                    FechaRegistro = dataRow["FechaRegistro"].ToString(),
                    Estatus = dataRow["Estatus"].ToString(),
                    UsuarioRegistra = dataRow["UsuarioRegistra"].ToString()
                    }).ToList();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw ex;
            }
            return lista;
        }

        public string InsertTRSPTransferencia(InsertTRSPModel trsp)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();

            string folioGenerado = GenerarNoFolio();

            parametros.Add(new SqlParameter { ParameterName = "@AlmacenOrigen", SqlDbType = SqlDbType.Int, Value = trsp.AlmacenOrigen });
            parametros.Add(new SqlParameter { ParameterName = "@AlmacenDestino", SqlDbType = SqlDbType.Int, Value = trsp.AlmacenDestino });
            parametros.Add(new SqlParameter { ParameterName = "@IdInsumo", SqlDbType = SqlDbType.Int, Value = trsp.IdInsumo });
            parametros.Add(new SqlParameter { ParameterName = "@Cantidad", SqlDbType = SqlDbType.Int, Value = trsp.Cantidad });
            parametros.Add(new SqlParameter { ParameterName = "@TipoMovimiento", SqlDbType = SqlDbType.Int, Value = trsp.TipoMovimiento });
            parametros.Add(new SqlParameter { ParameterName = "@Descripcion", SqlDbType = SqlDbType.VarChar, Value = trsp.Descripcion });
            parametros.Add(new SqlParameter { ParameterName = "@No_Folio", SqlDbType = SqlDbType.VarChar, Value = folioGenerado });
            parametros.Add(new SqlParameter { ParameterName = "@UsuarioRegistra", SqlDbType = SqlDbType.Int, Value = trsp.UsuarioRegistra });

            try
            {
                DataSet ds = dac.Fill("sp_insert_TRSP_RenglonTransferencia", parametros);
                folioGenerado = ds.Tables[0].Rows[0]["FolioGenerado"].ToString();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return folioGenerado;
        }

          public string InsertTRSPRenglonTransferencia(InsertTRSPRenglonModel trsp)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();

            string folioGenerado = GenerarNoFolio();

            parametros.Add(new SqlParameter { ParameterName = "@AlmacenOrigen", SqlDbType = SqlDbType.Int, Value = trsp.AlmacenOrigen });
            parametros.Add(new SqlParameter { ParameterName = "@AlmacenDestino", SqlDbType = SqlDbType.Int, Value = trsp.AlmacenDestino });
            parametros.Add(new SqlParameter { ParameterName = "@IdInsumo", SqlDbType = SqlDbType.Int, Value = trsp.IdInsumo });
            parametros.Add(new SqlParameter { ParameterName = "@Cantidad", SqlDbType = SqlDbType.Int, Value = trsp.Cantidad });
            parametros.Add(new SqlParameter { ParameterName = "@TipoMovimiento", SqlDbType = SqlDbType.Int, Value = trsp.TipoMovimiento });
            parametros.Add(new SqlParameter { ParameterName = "@Descripcion", SqlDbType = SqlDbType.VarChar, Value = trsp.Descripcion });
            parametros.Add(new SqlParameter { ParameterName = "@No_Folio", SqlDbType = SqlDbType.VarChar, Value = folioGenerado });
            parametros.Add(new SqlParameter { ParameterName = "@UsuarioRegistra", SqlDbType = SqlDbType.Int, Value = trsp.UsuarioRegistra });

            try
            {
                DataSet ds = dac.Fill("sp_insert_TRSP_RenglonTransferencia", parametros);
                folioGenerado = ds.Tables[0].Rows[0]["FolioGenerado"].ToString();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return folioGenerado;
        }

      public string UpdateTRSPTransferncia(UpdateTRSPModel trsp)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            string mensaje;
            parametros.Add(new SqlParameter { ParameterName = "@IdTRSP", SqlDbType = SqlDbType.Int, Value = trsp.IdTRSP });
            parametros.Add(new SqlParameter { ParameterName = "@AlmacenOrigen", SqlDbType = SqlDbType.Int, Value = trsp.AlmacenOrigen });
            parametros.Add(new SqlParameter { ParameterName = "@AlmacenDestino", SqlDbType = SqlDbType.Int, Value = trsp.AlmacenDestino });
            parametros.Add(new SqlParameter { ParameterName = "@IdInsumo", SqlDbType = SqlDbType.Int, Value = trsp.IdInsumo });
            parametros.Add(new SqlParameter { ParameterName = "@Cantidad", SqlDbType = SqlDbType.Int, Value = trsp.Cantidad });
            parametros.Add(new SqlParameter { ParameterName = "@TipoMovimiento", SqlDbType = SqlDbType.Int, Value = trsp.TipoMovimiento });
            parametros.Add(new SqlParameter { ParameterName = "@Descripcion", SqlDbType = SqlDbType.VarChar, Value = trsp.Descripcion });
            parametros.Add(new SqlParameter { ParameterName = "@Estatus", SqlDbType = SqlDbType.Int, Value = trsp.Estatus });
            parametros.Add(new SqlParameter { ParameterName = "@UsuarioRegistra", SqlDbType = SqlDbType.Int, Value = trsp.UsuarioRegistra });
            try
            {
                DataSet ds = dac.Fill("sp_update_TRSP_Transferencia", parametros);
                mensaje = ds.Tables[0].AsEnumerable().Select(dataRow => dataRow["mensaje"].ToString()).ToList()[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return mensaje;
        }

      public void DeleteTRSPTransferencia(int id)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            parametros.Add(new SqlParameter { ParameterName = "@IdTRSP", SqlDbType = SqlDbType.Int, Value = id });


            try
            {
                dac.ExecuteNonQuery("sp_delete_TRSP_Transferencia", parametros);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
