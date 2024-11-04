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
    public class MovimientosService
    {
        private  string connection;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private ArrayList parametros = new ArrayList();


        public MovimientosService(IMarcatelDatabaseSetting settings, IWebHostEnvironment webHostEnvironment)
        {
             connection = settings.ConnectionString;

             _webHostEnvironment = webHostEnvironment;
             
        }

        // public List<GetMovimientosModel> GetMovimientos()
        // {
        //     ConexionDataAccess dac = new ConexionDataAccess(connection);
        //     GetMovimientosModel movimientos = new GetMovimientosModel();

        //     List<GetMovimientosModel> lista = new List<GetMovimientosModel>();
        //     try
        //     {
        //         parametros = new ArrayList();
        //         DataSet ds = dac.Fill("sp_get_movimientos", parametros);
        //         if (ds.Tables[0].Rows.Count > 0)
        //         {

        //           lista = ds.Tables[0].AsEnumerable()
        //             .Select(dataRow => new GetMovimientosModel {
        //                 Id = int.Parse(dataRow["Id"].ToString()),
        //                 IdTipoMovimiento = int.Parse(dataRow["IdTipoMovimiento"].ToString()),
        //                 IdAlmacen = int.Parse(dataRow["IdAlmacen"].ToString()),
        //                 Fecha  = dataRow["Fecha"].ToString(),
        //                 Estatus = dataRow["Estatus"].ToString(),
        //                 FechaRegistro = dataRow["FechaRegistro"].ToString(),
        //                 IdUsuario = dataRow["IdUsuario"].ToString(),

        //             }).ToList();
        //         }
        //     }
        //     catch (Exception ex)
        //     {
        //         throw ex;
        //     }
        //     return lista;
        // }
                         public List<GetMovimientosModel> GetMovimientos(int IdTipoMovimiento)
        {

            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            parametros.Add(new SqlParameter { ParameterName = "@IdTipoMovimiento", SqlDbType = SqlDbType.Int, Value = IdTipoMovimiento });

            List<GetMovimientosModel> lista = new List<GetMovimientosModel>();
            try
            {
                DataSet ds = dac.Fill("sp_get_movimientos", parametros);
                if (ds.Tables[0].Rows.Count > 0)
                {

                  lista = ds.Tables[0].AsEnumerable()
                    .Select(dataRow => new GetMovimientosModel {
                      Id = int.Parse(dataRow["Id"].ToString()),
                        IdTipoMovimiento = int.Parse(dataRow["IdTipoMovimiento"].ToString()),
                        IdAlmacen = int.Parse(dataRow["IdAlmacen"].ToString()),
                        Fecha  = dataRow["Fecha"].ToString(),
                        Estatus = dataRow["Estatus"].ToString(),
                        FechaRegistro = dataRow["FechaRegistro"].ToString(),
                        IdUsuario = dataRow["IdUsuario"].ToString(),
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

        public string InsertMovimientos(InsertMovimientosModel Movimientos)
        {
            int IdMovimiento;
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            string mensaje;

            parametros.Add(new SqlParameter { ParameterName = "@IdTipoMovimiento", SqlDbType = System.Data.SqlDbType.VarChar, Value = Movimientos.IdTipoMovimiento});
            parametros.Add(new SqlParameter { ParameterName = "@IdAlmacen", SqlDbType = System.Data.SqlDbType.Int, Value = Movimientos.IdAlmacen});
            parametros.Add(new SqlParameter { ParameterName = "@IdUsuario", SqlDbType = System.Data.SqlDbType.Int, Value = Movimientos.IdUsuario });

            try 
            {
                DataSet ds = dac.Fill("sp_insert_movimientos", parametros);
                IdMovimiento = ds.Tables[0].AsEnumerable().Select(dataRow=>int.Parse(dataRow["IdMovimiento"].ToString())).ToList()[0];
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw ex;
            }
            return IdMovimiento.ToString();
        }

        public string UpdateMovimientos(UpdateMovimientosModel Movimientos)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            string mensaje;

            parametros.Add(new SqlParameter { ParameterName = "@Id", SqlDbType = System.Data.SqlDbType.Int, Value = Movimientos.Id});
            parametros.Add(new SqlParameter { ParameterName = "@IdTipoMovimiento", SqlDbType = System.Data.SqlDbType.VarChar, Value = Movimientos.IdTipoMovimiento});
            parametros.Add(new SqlParameter { ParameterName = "@IdAlmacen", SqlDbType = System.Data.SqlDbType.Int, Value = Movimientos.IdAlmacen});
            parametros.Add(new SqlParameter { ParameterName = "@IdUsuario", SqlDbType = System.Data.SqlDbType.Int, Value = Movimientos.IdUsuario });

            try
            {
                DataSet ds = dac.Fill("sp_update_movimientos", parametros);
                mensaje = ds.Tables[0].AsEnumerable().Select(dataRow => dataRow["mensaje"].ToString()).ToList()[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return mensaje;
        }

      public void DeleteMovimientos(int id)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            parametros.Add(new SqlParameter { ParameterName = "@Id", SqlDbType = SqlDbType.Int, Value = id });


            try
            {
                dac.ExecuteNonQuery("sp_delete_movimientos", parametros);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}