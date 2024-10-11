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
    public class OrdenCompraService
    {
        private  string connection;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private ArrayList parametros = new ArrayList();


        public OrdenCompraService(IMarcatelDatabaseSetting settings, IWebHostEnvironment webHostEnvironment)
        {
             connection = settings.ConnectionString;

             _webHostEnvironment = webHostEnvironment;
             
        }

        public List<GetOrdenCompraModel> GetOrdenCompra()
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            GetOrdenCompraModel ordencompra = new GetOrdenCompraModel();

            List<GetOrdenCompraModel> lista = new List<GetOrdenCompraModel>();
            try
            {
                parametros = new ArrayList();
                DataSet ds = dac.Fill("sp_get_ordencompra", parametros);
                if (ds.Tables[0].Rows.Count > 0)
                {

                  lista = ds.Tables[0].AsEnumerable()
                    .Select(dataRow => new GetOrdenCompraModel {
                        Id = int.Parse(dataRow["Id"].ToString()),
                        IdProveedor = int.Parse(dataRow["IdProveedor"].ToString()),
                        FechaLlegada = dataRow["FechaLlegada"].ToString(),
                        IdSucursal = int.Parse(dataRow["IdSucursal"].ToString()),
                        IdComprador = int.Parse(dataRow["IdComprador"].ToString()),
                        Fecha = dataRow["Fecha"].ToString(),
                        Total = decimal.Parse(dataRow["Total"].ToString()),
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

        public string InsertOrdenCompra(InsertOrdenCompraModel ordencompra)
        {
            int IdOrdenCompra;
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            string mensaje;

            parametros.Add(new SqlParameter { ParameterName = "@IdProveedor", SqlDbType = System.Data.SqlDbType.Int, Value = ordencompra.IdProveedor });
            parametros.Add(new SqlParameter { ParameterName = "@IdSucursal", SqlDbType = System.Data.SqlDbType.Int, Value = ordencompra.IdSucursal});
            parametros.Add(new SqlParameter { ParameterName = "@IdComprador", SqlDbType = System.Data.SqlDbType.Int, Value = ordencompra.IdComprador});
            parametros.Add(new SqlParameter { ParameterName = "@Total", SqlDbType = System.Data.SqlDbType.Decimal, Value = ordencompra.Total});
            parametros.Add(new SqlParameter { ParameterName = "@UsuarioRegistra", SqlDbType = System.Data.SqlDbType.Int, Value = ordencompra.UsuarioRegistra });

             try 
            {
                DataSet ds = dac.Fill("sp_insert_ordencompra", parametros);
                IdOrdenCompra = ds.Tables[0].AsEnumerable().Select(dataRow=>int.Parse(dataRow["IdOrdenCompra"].ToString())).ToList()[0];
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw ex;
            }
            return IdOrdenCompra.ToString();
        }

        public string UpdateOrdenCompra(UpdateOrdenCompraModel ordencompra)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            string mensaje;


            parametros.Add(new SqlParameter { ParameterName = "@Id", SqlDbType = System.Data.SqlDbType.Int, Value = ordencompra.Id });
            parametros.Add(new SqlParameter { ParameterName = "@IdProveedor", SqlDbType = System.Data.SqlDbType.Int, Value = ordencompra.IdProveedor });
            parametros.Add(new SqlParameter { ParameterName = "@IdSucursal", SqlDbType = System.Data.SqlDbType.VarChar, Value = ordencompra.IdSucursal});
            parametros.Add(new SqlParameter { ParameterName = "@IdComprador", SqlDbType = System.Data.SqlDbType.Decimal, Value = ordencompra.IdComprador});
            parametros.Add(new SqlParameter { ParameterName = "@Total", SqlDbType = System.Data.SqlDbType.Decimal, Value = ordencompra.Total});
            parametros.Add(new SqlParameter { ParameterName = "@UsuarioRegistra", SqlDbType = System.Data.SqlDbType.Int, Value = ordencompra.UsuarioRegistra });


            try
            {
                DataSet ds = dac.Fill("sp_update_ordencompra", parametros);
                mensaje = ds.Tables[0].AsEnumerable().Select(dataRow => dataRow["mensaje"].ToString()).ToList()[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return mensaje;
        }

      public void DeleteOrdenCompra(int id)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            parametros.Add(new SqlParameter { ParameterName = "@Id", SqlDbType = SqlDbType.Int, Value = id });


            try
            {
                dac.ExecuteNonQuery("sp_delete_ordencompra", parametros);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}