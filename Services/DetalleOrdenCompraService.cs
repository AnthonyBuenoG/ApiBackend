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
    public class DetalleOrdenCompraService
    {
        private  string connection;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private ArrayList parametros = new ArrayList();


        public DetalleOrdenCompraService(IMarcatelDatabaseSetting settings, IWebHostEnvironment webHostEnvironment)
        {
             connection = settings.ConnectionString;

             _webHostEnvironment = webHostEnvironment;
             
        }

        public List<GetDetalleOrdenCompraModel> GetDetalleOrdenCompra(int IdOredenCompra)
        {

            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            parametros.Add(new SqlParameter { ParameterName = "@IdOrdenCompra", SqlDbType = SqlDbType.Int, Value = IdOredenCompra });

            List<GetDetalleOrdenCompraModel> lista = new List<GetDetalleOrdenCompraModel>();
            try
            {
                DataSet ds = dac.Fill("sp_get_detalleordencompra", parametros);
                if (ds.Tables[0].Rows.Count > 0)
                {

                  lista = ds.Tables[0].AsEnumerable()
                    .Select(dataRow => new GetDetalleOrdenCompraModel {
                        Id = int.Parse(dataRow["Id"].ToString()),
                        IdOredenCompra = int.Parse(dataRow["IdOrdenCompra"].ToString()),
                        Insumo = dataRow["Insumo"].ToString(),
                        Cantidad = decimal.Parse(dataRow["Cantidad"].ToString()),
                        CantidadRecibida = decimal.Parse(dataRow["CantidadRecibida"].ToString()),
                        Costo = decimal.Parse(dataRow["Costo"].ToString()),
                        CostoRenglon = decimal.Parse(dataRow["CostoRenglon"].ToString()),
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

        public string InsertDetalleOrdenCompra(InsertDetalleOrdenCompraModel detalleordencompra)
        {
            int IdDetalleOrdenCmpra;
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            string mensaje;

            parametros.Add(new SqlParameter { ParameterName = "@IdOrdenCompra", SqlDbType = System.Data.SqlDbType.VarChar, Value = detalleordencompra.IdOredenCompra });
            parametros.Add(new SqlParameter { ParameterName = "@Insumo", SqlDbType = System.Data.SqlDbType.VarChar, Value = detalleordencompra.Insumo});
            parametros.Add(new SqlParameter { ParameterName = "@Cantidad", SqlDbType = System.Data.SqlDbType.Decimal, Value = detalleordencompra.Cantidad});
            parametros.Add(new SqlParameter { ParameterName = "@CantidadRecibida", SqlDbType = System.Data.SqlDbType.Decimal, Value = detalleordencompra.CantidadRecibida});
            parametros.Add(new SqlParameter { ParameterName = "@Costo", SqlDbType = System.Data.SqlDbType.Decimal, Value = detalleordencompra.Costo});
            parametros.Add(new SqlParameter { ParameterName = "@CostoRenglon", SqlDbType = System.Data.SqlDbType.Decimal, Value = detalleordencompra.CostoRenglon});
            parametros.Add(new SqlParameter { ParameterName = "@UsuarioRegistra", SqlDbType = System.Data.SqlDbType.Int, Value = detalleordencompra.UsuarioRegistra });

            try 
            {
                DataSet ds = dac.Fill("sp_insert_detalleordencompra", parametros);
                IdDetalleOrdenCmpra = ds.Tables[0].AsEnumerable().Select(dataRow=>int.Parse(dataRow["IdDetalleOrdenCompra"].ToString())).ToList()[0];
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw ex;
            }
            return IdDetalleOrdenCmpra.ToString();
        }

        public string UpdateDetalleOrdenCompra(UpdateDetalleOrdenCompraModel detalleordencompra)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            string mensaje;


            parametros.Add(new SqlParameter { ParameterName = "@Id", SqlDbType = System.Data.SqlDbType.Int, Value = detalleordencompra.Id });
            parametros.Add(new SqlParameter { ParameterName = "@IdOrdenCompra", SqlDbType = System.Data.SqlDbType.Int, Value = detalleordencompra.IdOredenCompra });
            parametros.Add(new SqlParameter { ParameterName = "@Insumo", SqlDbType = System.Data.SqlDbType.VarChar, Value = detalleordencompra.Insumo});
            parametros.Add(new SqlParameter { ParameterName = "@Cantidad", SqlDbType = System.Data.SqlDbType.Decimal, Value = detalleordencompra.Cantidad});
            parametros.Add(new SqlParameter { ParameterName = "@CantidadRecibida", SqlDbType = System.Data.SqlDbType.Decimal, Value = detalleordencompra.CantidadRecibida});
            parametros.Add(new SqlParameter { ParameterName = "@Costo", SqlDbType = System.Data.SqlDbType.Decimal, Value = detalleordencompra.Costo});
            parametros.Add(new SqlParameter { ParameterName = "@CostoRenglon", SqlDbType = System.Data.SqlDbType.Decimal, Value = detalleordencompra.CostoRenglon});
            parametros.Add(new SqlParameter { ParameterName = "@UsuarioRegistra", SqlDbType = System.Data.SqlDbType.Int, Value = detalleordencompra.UsuarioRegistra });


            try
            {
                DataSet ds = dac.Fill("sp_update_detalleordencompra", parametros);
                mensaje = ds.Tables[0].AsEnumerable().Select(dataRow => dataRow["mensaje"].ToString()).ToList()[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return mensaje;
        }

      public void DeleteDetalleOrdenCompra(int id)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            parametros.Add(new SqlParameter { ParameterName = "@Id", SqlDbType = SqlDbType.Int, Value = id });


            try
            {
                dac.ExecuteNonQuery("sp_delete_detalleordencompra", parametros);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}