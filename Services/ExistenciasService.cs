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
    public class ExistenciasService
    {
        private  string connection;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private ArrayList parametros = new ArrayList();


        public ExistenciasService(IMarcatelDatabaseSetting settings, IWebHostEnvironment webHostEnvironment)
        {
             connection = settings.ConnectionString;

             _webHostEnvironment = webHostEnvironment;
             
        }

        public List<GetExistenciasModel> GetExistencias()
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            GetExistenciasModel existencias = new GetExistenciasModel();

            List<GetExistenciasModel> lista = new List<GetExistenciasModel>();
            try
            {
                parametros = new ArrayList();
                DataSet ds = dac.Fill("sp_get_existencia", parametros);
                if (ds.Tables[0].Rows.Count > 0)
                {

                  lista = ds.Tables[0].AsEnumerable()
                    .Select(dataRow => new GetExistenciasModel {
                        Id = int.Parse(dataRow["Id"].ToString()),
                        Fecha = dataRow["Fecha"].ToString(),
                        Insumo = dataRow["Insumo"].ToString(),
                        DescripcionInsumo = dataRow["DescripcionInsumo"].ToString(),
                        Cantidad = dataRow["Cantidad"].ToString(),
                        IdAlmacen = int.Parse(dataRow["IdAlmacen"].ToString()),
                        Estatus = dataRow["Estatus"].ToString(),
                        FechaRegistro = dataRow["FechaRegistro"].ToString(),
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

        // public List<GetExistenciasModel> GetExistencias(DateTime? fechaInicio = null, DateTime? fechaFin = null, int? idAlmacen = null)
        // {
        //     ConexionDataAccess dac = new ConexionDataAccess(connection);
        //     List<GetExistenciasModel> lista = new List<GetExistenciasModel>();
            
        //     try
        //     {
        //         ArrayList parametros = new ArrayList();

        //         if (fechaInicio.HasValue)
        //             parametros.Add(new SqlParameter("@fechaInicio", fechaInicio.Value));
                
        //         if (fechaFin.HasValue)
        //             parametros.Add(new SqlParameter("@fechaFin", fechaFin.Value));
                
        //         if (idAlmacen.HasValue)
        //             parametros.Add(new SqlParameter("@idAlmacen", idAlmacen.Value));

        //         DataSet ds = dac.Fill("sp_get_existencia", parametros);

        //         if (ds.Tables[0].Rows.Count > 0)
        //         {
        //             lista = ds.Tables[0].AsEnumerable()
        //                 .Select(dataRow => new GetExistenciasModel
        //                 {
        //                     Id = int.Parse(dataRow["Id"].ToString()),
        //                     Fecha = dataRow["Fecha"].ToString(),
        //                     Insumo = dataRow["Insumo"].ToString(),
        //                     DescripcionInsumo = dataRow["DescripcionInsumo"].ToString(),
        //                     Cantidad = dataRow["Cantidad"].ToString(),
        //                     IdAlmacen = int.Parse(dataRow["IdAlmacen"].ToString()),
        //                     Estatus = dataRow["Estatus"].ToString(),
        //                     FechaRegistro = dataRow["FechaRegistro"].ToString(),
        //                     UsuarioRegistra = dataRow["UsuarioRegistra"].ToString()
        //                 })
        //                 .ToList();
        //         }
        //     }
        //     catch (Exception ex)
        //     {
        //         throw ex;
        //     }

        //     return lista;
        // }


        public string InsertExistencias(InsertExistenciasModel Existencias)
        {

            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            string mensaje;

            parametros.Add(new SqlParameter { ParameterName = "@Insumo", SqlDbType = System.Data.SqlDbType.VarChar, Value = Existencias.Insumo});
            parametros.Add(new SqlParameter { ParameterName = "@Cantidad", SqlDbType = System.Data.SqlDbType.Decimal, Value = Existencias.Cantidad});
            parametros.Add(new SqlParameter { ParameterName = "@IdAlmacen", SqlDbType = System.Data.SqlDbType.Int, Value = Existencias.IdAlmacen });
            parametros.Add(new SqlParameter { ParameterName = "@UsuarioRegistra", SqlDbType = System.Data.SqlDbType. Int, Value = Existencias.UsuarioRegistra });

            try 
            {
                DataSet ds = dac.Fill("sp_insert_existencias", parametros);
                mensaje = ds.Tables[0].AsEnumerable().Select(dataRow => dataRow["mensaje"].ToString()).ToList()[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return mensaje;
        }

        public string UpdateExistencias(UpdateExistenciasModel Existencias)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            string mensaje;


            parametros.Add(new SqlParameter { ParameterName = "@Id", SqlDbType = System.Data.SqlDbType.Int, Value = Existencias.Id });
            parametros.Add(new SqlParameter { ParameterName = "@Insumo", SqlDbType = System.Data.SqlDbType.VarChar, Value = Existencias.Insumo});
            parametros.Add(new SqlParameter { ParameterName = "@Cantidad", SqlDbType = System.Data.SqlDbType.Decimal, Value = Existencias.Cantidad});
            parametros.Add(new SqlParameter { ParameterName = "@IdAlmacen", SqlDbType = System.Data.SqlDbType.Int, Value = Existencias.IdAlmacen });
            parametros.Add(new SqlParameter { ParameterName = "@UsuarioRegistra", SqlDbType = System.Data.SqlDbType.Int, Value = Existencias.UsuarioRegistra });

            try
            {
                DataSet ds = dac.Fill("sp_update_existencias", parametros);
                mensaje = ds.Tables[0].AsEnumerable().Select(dataRow => dataRow["mensaje"].ToString()).ToList()[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return mensaje;
        }

      public void DeleteExistencias(int id)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            parametros.Add(new SqlParameter { ParameterName = "@IdExistencia", SqlDbType = SqlDbType.Int, Value = id });


            try
            {
                dac.ExecuteNonQuery("sp_delete_existencias", parametros);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}