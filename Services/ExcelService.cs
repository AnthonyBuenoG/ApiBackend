using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using reportesApi.DataContext;
using reportesApi.Models;
using System.Collections.Generic;
using reportesApi.Models.Compras;
using OfficeOpenXml;
using Microsoft.AspNetCore.Hosting;
using System.IO;
using Microsoft.AspNetCore.Mvc;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
namespace reportesApi.Services
{
    public class ExcelService
    {
        private  string connection;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private ArrayList parametros = new ArrayList();


        public ExcelService(IMarcatelDatabaseSetting settings, IWebHostEnvironment webHostEnvironment)
        {
             connection = settings.ConnectionString;

             _webHostEnvironment = webHostEnvironment;
             
        }

         public List<RecetaRangoFechasModel> GetRecetasPorRangoDeFechas(DateTime fechaInicio, DateTime fechaFin)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            List<RecetaRangoFechasModel> lista = new List<RecetaRangoFechasModel>();

            try
            {
                parametros = new ArrayList
                {
                    new SqlParameter("@FechaInicio", fechaInicio),
                    new SqlParameter("@FechaFin", fechaFin)
                };

                DataSet ds = dac.Fill("sp_ExportarRecetasRangoFechas", parametros);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    lista = ds.Tables[0].AsEnumerable()
                        .Select(dataRow => new RecetaRangoFechasModel
                        {
                            RecetaID = dataRow["RecetaID"].ToString(),
                            NombreReceta = dataRow["NombreReceta"].ToString(),
                            FechaCreacion = dataRow["FechaCreacion"].ToString(),
                            UsuarioRegistra = dataRow["UsuarioRegistra"].ToString(),
                        }).ToList();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return lista;
        }


        public List<ContenidoRecetaModel> GetContenidoRecetaPorId(int idReceta)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            List<ContenidoRecetaModel> lista = new List<ContenidoRecetaModel>();

            try
            {
                parametros = new ArrayList
                {
                    new SqlParameter("@IdReceta", idReceta)
                };

                DataSet ds = dac.Fill("sp_ExportarContenidoReceta_Id", parametros);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    lista = ds.Tables[0].AsEnumerable()
                        .Select(dataRow => new ContenidoRecetaModel
                        {
                            RecetaID = dataRow["RecetaID"].ToString(),
                            NombreReceta = dataRow["NombreReceta"].ToString(),
                            Insumo = dataRow["Insumo"].ToString(), 
                            DescripcionInsumo = dataRow["DescripcionInsumo"].ToString(),
                            FechaCreacion = dataRow["FechaCreacion"].ToString(),
                            UsuarioRegistra = dataRow["UsuarioRegistra"].ToString(),
                            
                         
                        }).ToList();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return lista;
        }

        public List<TraspasoEntradaModel> GetTraspasosEntrada(DateTime fechaInicio, DateTime fechaFin, int almacen, int tipoMovimiento)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            List<TraspasoEntradaModel> lista = new List<TraspasoEntradaModel>();

            try
            {
                parametros = new ArrayList
                {
                    new SqlParameter("@FechaInicio", fechaInicio),
                    new SqlParameter("@FechaFin", fechaFin),
                    new SqlParameter("@Almacen", almacen),
                    new SqlParameter("@TipoMovimiento", tipoMovimiento)

                };

                DataSet ds = dac.Fill("sp_ExportarTraspasosEntrada", parametros);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    lista = ds.Tables[0].AsEnumerable()
                        .Select(dataRow => new TraspasoEntradaModel
                        {
                            IdTRSP = dataRow["IdTRSP"].ToString(),
                            AlmacenOrigen = dataRow["AlmacenOrigen"].ToString(),
                            NombreAlmacenOrigen = dataRow["NombreAlmacenOrgien"].ToString(),
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
                throw ex;
            }

            return lista;
        }

        public List<TraspasoSalidaModel> GetTraspasosSalida(DateTime fechaInicio, DateTime fechaFin, int almacen, int tipoMovimiento)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            List<TraspasoSalidaModel> lista = new List<TraspasoSalidaModel>();

            try
            {
                parametros = new ArrayList
                {
                    new SqlParameter("@FechaInicio", fechaInicio),
                    new SqlParameter("@FechaFin", fechaFin),
                    new SqlParameter("@Almacen", almacen),
                    new SqlParameter("@TipoMovimiento", tipoMovimiento)
                };

                DataSet ds = dac.Fill("sp_ExportarTraspasosEntrada", parametros);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    lista = ds.Tables[0].AsEnumerable()
                        .Select(dataRow => new TraspasoSalidaModel
                        {
                            IdTRSP = dataRow["IdTRSP"].ToString(),
                            AlmacenOrigen = dataRow["AlmacenOrigen"].ToString(),
                            NombreAlmacenOrigen = dataRow["NombreAlmacenOrgien"].ToString(),
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
                throw ex;
            }

            return lista;
        }




    }
}