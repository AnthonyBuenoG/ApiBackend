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
    public class RenglonMovimientoService
    {
        private  string connection;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private ArrayList parametros = new ArrayList();


        public RenglonMovimientoService (IMarcatelDatabaseSetting settings, IWebHostEnvironment webHostEnvironment)
        {
             connection = settings.ConnectionString;

             _webHostEnvironment = webHostEnvironment;

        }

        // public List<GetRenglonMovimientosModel> GetRenglonMovimientos(int IdMovimiento)
        // {
        //     ConexionDataAccess dac = new ConexionDataAccess(connection);
        //     GetRenglonMovimientosModel renglonmovimientos = new GetRenglonMovimientosModel();
        //     parametros.Add(new SqlParameter { ParameterName = "@IdMovimiento", SqlDbType = SqlDbType.Int, Value = IdMovimiento });

        //     List<GetRenglonMovimientosModel> lista = new List<GetRenglonMovimientosModel>();
        //     try
        //     {
        //         parametros = new ArrayList();
        //         DataSet ds = dac.Fill("sp_get_renglonesmovimientos", parametros);
        //         if (ds.Tables[0].Rows.Count > 0)
        //         {

        //           lista = ds.Tables[0].AsEnumerable()
        //             .Select(dataRow => new GetRenglonMovimientosModel {
        //                 Id = int.Parse(dataRow["Id"].ToString()),
        //                 IdMovimiento = int.Parse(dataRow["IdMovimiento"].ToString()),
        //                 Insumo = dataRow["Insumo"].ToString(),
        //                 DescripcionInsumo = dataRow["DescripcionInsumo"].ToString(),
        //                 Cantidad = decimal.Parse(dataRow["Cantidad"].ToString()),
        //                 Costo = decimal.Parse(dataRow["Costo"].ToString()),
        //                 Estatus = dataRow["Estatus"].ToString(),
        //                 FechaRegistro = dataRow["FechaRegistro"].ToString(),
        //                 UsuarioRegistra = dataRow["UsuarioRegistra"].ToString(),


        //             }).ToList();
        //         }
        //     }
        //     catch (Exception ex)
        //     {
        //         throw ex;
        //     }
        //     return lista;
        // }
        //          public List<GetRenglonMovimientosModel> GetRenglonMovimimentos(int IdMovimiento)
        // {

        //     ConexionDataAccess dac = new ConexionDataAccess(connection);
        //     parametros = new ArrayList();
        //     parametros.Add(new SqlParameter { ParameterName = "@IdMovimiento", SqlDbType = SqlDbType.Int, Value = IdMovimiento });

        //     List<GetRenglonMovimientosModel> lista = new List<GetRenglonMovimientosModel>();
        //     try
        //     {
        //         DataSet ds = dac.Fill("sp_get_renglonesmovimientos", parametros);
        //         if (ds.Tables[0].Rows.Count > 0)
        //         {

        //           lista = ds.Tables[0].AsEnumerable()
        //             .Select(dataRow => new GetRenglonMovimientosModel {
        //                 Id = int.Parse(dataRow["Id"].ToString()),
        //                 IdMovimiento = int.Parse(dataRow["IdMovimiento"].ToString()),
        //                 Insumo = dataRow["Insumo"].ToString(),
        //                 DescripcionInsumo = dataRow["DescripcionInsumo"].ToString(),
        //                 Cantidad = decimal.Parse(dataRow["Cantidad"].ToString()),
        //                 Costo = decimal.Parse(dataRow["Costo"].ToString()),
        //                 Estatus = dataRow["Estatus"].ToString(),
        //                 FechaRegistro = dataRow["FechaRegistro"].ToString(),
        //                 UsuarioRegistra = dataRow["UsuarioRegistra"].ToString(),

        //             }).ToList();
        //         }
        //     }
        //     catch (Exception ex)
        //     {
        //         Console.WriteLine(ex.Message);
        //         throw ex;
        //     }
        //     return lista;
        // }

        public List<GetRenglonMovimientosModel> GetRenglonMovimimentos(DateTime? fechaInicio = null, DateTime? fechaFin = null)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            List<GetRenglonMovimientosModel> lista = new List<GetRenglonMovimientosModel>();
            
            try
            {
                ArrayList parametros = new ArrayList();

                if (fechaInicio.HasValue)
                    parametros.Add(new SqlParameter("@FechaInicio", fechaInicio.Value));
                
                if (fechaFin.HasValue)
                    parametros.Add(new SqlParameter("@FechaFin", fechaFin.Value));
                

                DataSet ds = dac.Fill("sp_get_renglonesmovimientos", parametros);

                if (ds.Tables[0].Rows.Count > 0)
                {
                    lista = ds.Tables[0].AsEnumerable()
                        .Select(dataRow => new GetRenglonMovimientosModel
                        {
                            Id = int.Parse(dataRow["Id"].ToString()),
                            IdMovimiento = int.Parse(dataRow["IdMovimiento"].ToString()),
                            Nombre = dataRow["TipoMovimientoNombre"].ToString(),
                            Insumo = dataRow["Insumo"].ToString(),
                            DescripcionInsumo = dataRow["DescripcionInsumo"].ToString(),
                            Cantidad = decimal.Parse(dataRow["Cantidad"].ToString()),
                            Costo = decimal.Parse(dataRow["Costo"].ToString()),
                            CostoTotal = dataRow["CostoTotal"].ToString(),
                            Estatus = dataRow["Estatus"].ToString(),
                            FechaRegistro = dataRow["FechaRegistro"].ToString(),
                            UsuarioRegistra = dataRow["UsuarioRegistra"].ToString(),

                        })
                        .ToList();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return lista;
        }


        public string InsertRenglonMovimientos(InsertRenglonMovimientosModel RenglonMovimientos)
        {

            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            string mensaje;

            parametros.Add(new SqlParameter { ParameterName = "@IdMovimiento", SqlDbType = System.Data.SqlDbType.Int, Value = RenglonMovimientos.IdMovimiento});
            parametros.Add(new SqlParameter { ParameterName = "@Insumo", SqlDbType = System.Data.SqlDbType.VarChar, Value = RenglonMovimientos.Insumo});
            parametros.Add(new SqlParameter { ParameterName = "@Cantidad", SqlDbType = System.Data.SqlDbType.Decimal, Value = RenglonMovimientos.Cantidad });
            parametros.Add(new SqlParameter { ParameterName = "@Costo", SqlDbType = System.Data.SqlDbType.Decimal, Value = RenglonMovimientos.Costo});
            parametros.Add(new SqlParameter { ParameterName = "@UsuarioRegistra", SqlDbType = System.Data.SqlDbType.Int, Value = RenglonMovimientos.UsuarioRegistra });

            try
            {
                DataSet ds = dac.Fill("sp_insert_renglonesmovimientos", parametros);
                mensaje = ds.Tables[0].AsEnumerable().Select(dataRow => dataRow["mensaje"].ToString()).ToList()[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return mensaje;
        }

        public string UpdateRenglonMovimientos(UpdateRenglonMovimientosModel RenglonMovimientos)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            string mensaje;

            parametros.Add(new SqlParameter { ParameterName = "@Id", SqlDbType = System.Data.SqlDbType.Int, Value = RenglonMovimientos.Id});
            parametros.Add(new SqlParameter { ParameterName = "@IdMovimiento", SqlDbType = System.Data.SqlDbType.Int, Value = RenglonMovimientos.IdMovimiento});
            parametros.Add(new SqlParameter { ParameterName = "@Insumo", SqlDbType = System.Data.SqlDbType.VarChar, Value = RenglonMovimientos.Insumo});
            parametros.Add(new SqlParameter { ParameterName = "@Cantidad", SqlDbType = System.Data.SqlDbType.Decimal, Value = RenglonMovimientos.Cantidad });
            parametros.Add(new SqlParameter { ParameterName = "@Costo", SqlDbType = System.Data.SqlDbType.Decimal, Value = RenglonMovimientos.Costo});
            parametros.Add(new SqlParameter { ParameterName = "@UsuarioRegistro", SqlDbType = System.Data.SqlDbType.Int, Value = RenglonMovimientos.UsuarioRegistra });
            try
            {
                DataSet ds = dac.Fill("sp_update_renglonmovimientos", parametros);
                mensaje = ds.Tables[0].AsEnumerable().Select(dataRow => dataRow["mensaje"].ToString()).ToList()[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return mensaje;
        }

      public void DeleteRenglonMovimientos(int id)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            parametros.Add(new SqlParameter { ParameterName = "@Id", SqlDbType = SqlDbType.Int, Value = id });


            try
            {
                dac.ExecuteNonQuery("sp_delete_renglonesmovimientos", parametros);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}