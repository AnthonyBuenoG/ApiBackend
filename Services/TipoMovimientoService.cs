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
    public class TipoMovimientoService
    {
        private  string connection;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private ArrayList parametros = new ArrayList();


        public TipoMovimientoService(IMarcatelDatabaseSetting settings, IWebHostEnvironment webHostEnvironment)
        {
             connection = settings.ConnectionString;

             _webHostEnvironment = webHostEnvironment;
             
        }

        public List<GetTipoMovimientoModel> GetTipoMovimiento()
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            GetTipoMovimientoModel tipomovimiento = new GetTipoMovimientoModel();

            List<GetTipoMovimientoModel> lista = new List<GetTipoMovimientoModel>();
            try
            {
                parametros = new ArrayList();
                DataSet ds = dac.Fill("sp_get_tipomovimiento", parametros);
                if (ds.Tables[0].Rows.Count > 0)
                {

                  lista = ds.Tables[0].AsEnumerable()
                    .Select(dataRow => new GetTipoMovimientoModel {
                        Id = int.Parse(dataRow["Id"].ToString()),
                        Nombre = dataRow["Nombre"].ToString(),
                        EntradaSalida = dataRow["EntradaSalida"].ToString(),
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

        public string InsertTipoMovimiento(InsertTipoMovimientoModel TipoMovimiento)
        {

            int IdTipoMovimiento;
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            string mensaje;

            parametros.Add(new SqlParameter { ParameterName = "@Nombre", SqlDbType = System.Data.SqlDbType.VarChar, Value = TipoMovimiento.Nombre});
            parametros.Add(new SqlParameter { ParameterName = "@EntradaSalida", SqlDbType = System.Data.SqlDbType.Int, Value = TipoMovimiento.EntradaSalida});
            parametros.Add(new SqlParameter { ParameterName = "@UsuarioRegistra", SqlDbType = System.Data.SqlDbType.Int, Value = TipoMovimiento.UsuarioRegistra });

            try 
            {
                DataSet ds = dac.Fill("sp_insert_tipomovimiento", parametros);
               IdTipoMovimiento = ds.Tables[0].AsEnumerable().Select(dataRow=>int.Parse(dataRow["IdTipoMovimiento"].ToString())).ToList()[0];
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw ex;
            }
            return IdTipoMovimiento.ToString();
        }

        public string UpdateTipoMovimiento(UpdateTipoMovimientoModel TipoMovimiento)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            string mensaje;


            parametros.Add(new SqlParameter { ParameterName = "@Id", SqlDbType = System.Data.SqlDbType.VarChar, Value = TipoMovimiento.Id });
            parametros.Add(new SqlParameter { ParameterName = "@Nombre", SqlDbType = System.Data.SqlDbType.VarChar, Value = TipoMovimiento.Nombre});
            parametros.Add(new SqlParameter { ParameterName = "@EntradaSalida", SqlDbType = System.Data.SqlDbType.Int, Value = TipoMovimiento.EntradaSalida});
            parametros.Add(new SqlParameter { ParameterName = "@UsuarioRegistra", SqlDbType = System.Data.SqlDbType.Int, Value = TipoMovimiento.UsuarioRegistra });

            try
            {
                DataSet ds = dac.Fill("sp_update_tipomovimiento", parametros);
                mensaje = ds.Tables[0].AsEnumerable().Select(dataRow => dataRow["mensaje"].ToString()).ToList()[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return mensaje;
        }

      public void DeleteTipoMovimiento(int id)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            parametros.Add(new SqlParameter { ParameterName = "@Id", SqlDbType = SqlDbType.Int, Value = id });


            try
            {
                dac.ExecuteNonQuery("sp_delete_tipomovimiento", parametros);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}