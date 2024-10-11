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
    public class DetalleRecetaService
    {
        private  string connection;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private ArrayList parametros = new ArrayList();


        public DetalleRecetaService(IMarcatelDatabaseSetting settings, IWebHostEnvironment webHostEnvironment)
        {
             connection = settings.ConnectionString;

             _webHostEnvironment = webHostEnvironment;
             
        }

         public List<GetDetalleRecetaModel> GetDetalleReceta(int IdReceta)
        {

            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            parametros.Add(new SqlParameter { ParameterName = "@IdReceta", SqlDbType = SqlDbType.Int, Value = IdReceta });

            List<GetDetalleRecetaModel> lista = new List<GetDetalleRecetaModel>();
            try
            {
                DataSet ds = dac.Fill("sp_get_detallereceta", parametros);
                if (ds.Tables[0].Rows.Count > 0)
                {

                  lista = ds.Tables[0].AsEnumerable()
                    .Select(dataRow => new GetDetalleRecetaModel {
                        Id = int.Parse(dataRow["Id"].ToString()),
                        IdReceta = int.Parse(dataRow["IdReceta"].ToString()),
                        Insumo = dataRow["Insumo"].ToString(),
                        Cantidad = int.Parse(dataRow["Cantidad"].ToString()),
                        UsuarioRegistra = dataRow["UsuarioRegistra"].ToString(),
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
        public string InsertDetalleReceta(InsertDetalleRecetaModel detallereceta)
        {
            int IdDetalleReceta;
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            string mensaje;

            parametros.Add(new SqlParameter { ParameterName = "@IdReceta", SqlDbType = System.Data.SqlDbType.Int, Value = detallereceta.IdReceta });
            parametros.Add(new SqlParameter { ParameterName = "@Insumo", SqlDbType = System.Data.SqlDbType.VarChar, Value = detallereceta.Insumo});
            parametros.Add(new SqlParameter { ParameterName = "@Cantidad", SqlDbType = System.Data.SqlDbType.Decimal, Value = detallereceta.Cantidad});
            parametros.Add(new SqlParameter { ParameterName = "@UsuarioRegistra", SqlDbType = System.Data.SqlDbType.Int, Value = detallereceta.UsuarioRegistra });


             try 
            {
                DataSet ds = dac.Fill("sp_insert_detallereceta", parametros);
                IdDetalleReceta = ds.Tables[0].AsEnumerable().Select(dataRow=>int.Parse(dataRow["IdDetalleReceta"].ToString())).ToList()[0];
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw ex;
            }
            return IdDetalleReceta.ToString();
        }

        public string UpdateDetalleReceta(UpdateDetalleRecetaModel detallereceta)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            string mensaje;


            parametros.Add(new SqlParameter { ParameterName = "@Id", SqlDbType = System.Data.SqlDbType.Int, Value = detallereceta.Id });
            parametros.Add(new SqlParameter { ParameterName = "@IdReceta", SqlDbType = System.Data.SqlDbType.Int, Value = detallereceta.IdReceta });
            parametros.Add(new SqlParameter { ParameterName = "@Insumo", SqlDbType = System.Data.SqlDbType.VarChar, Value = detallereceta.Insumo});
            parametros.Add(new SqlParameter { ParameterName = "@Cantidad", SqlDbType = System.Data.SqlDbType.Decimal, Value = detallereceta.Cantidad});
            parametros.Add(new SqlParameter { ParameterName = "@UsuarioRegistra", SqlDbType = System.Data.SqlDbType.Int, Value = detallereceta.UsuarioRegistra });

            try
            {
                DataSet ds = dac.Fill("sp_update_detallereceta", parametros);
                mensaje = ds.Tables[0].AsEnumerable().Select(dataRow => dataRow["mensaje"].ToString()).ToList()[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return mensaje;
        }

      public void DeleteDetalleReceta(int id)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            parametros.Add(new SqlParameter { ParameterName = "@Id", SqlDbType = SqlDbType.Int, Value = id });


            try
            {
                dac.ExecuteNonQuery("sp_delete_detallereceta", parametros);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}