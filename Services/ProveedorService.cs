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
    public class ProveedorService
    {
        private  string connection;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private ArrayList parametros = new ArrayList();


        public ProveedorService(IMarcatelDatabaseSetting settings, IWebHostEnvironment webHostEnvironment)
        {
             connection = settings.ConnectionString;

             _webHostEnvironment = webHostEnvironment;
             
        }

        public List<GetProveedorModel> GetProveedor()
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            GetProveedorModel proveedor = new GetProveedorModel();

            List<GetProveedorModel> lista = new List<GetProveedorModel>();
            try
            {
                parametros = new ArrayList();
                DataSet ds = dac.Fill("sp_get_proveedores", parametros);
                if (ds.Tables[0].Rows.Count > 0)
                {

                  lista = ds.Tables[0].AsEnumerable()
                    .Select(dataRow => new GetProveedorModel {
                        IdProveedor = int.Parse(dataRow["IdProveedor"].ToString()),
                        Nombre = dataRow["Nombre"].ToString(),
                        Direccion = dataRow["Direccion"].ToString(),
                        Email = dataRow["Email"].ToString(),
                        RFC = dataRow["RFC"].ToString(),
                        PlazoPago = int.Parse(dataRow["PlazoPago"].ToString()),
                        PorcentajeRetencion = decimal.Parse(dataRow["PorcentajeRetencion"].ToString()),
                        Estatus = dataRow["Estatus"].ToString(),
                        FechaRegistro= dataRow["FechaRegistro"].ToString(),
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

        public string InsertProveedor(InsertProveedorModel Proveedor)
        {
            int IdProveedor;
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            string mensaje;

            parametros.Add(new SqlParameter { ParameterName = "@Nombre", SqlDbType = System.Data.SqlDbType.VarChar, Value = Proveedor.Nombre});
            parametros.Add(new SqlParameter { ParameterName = "@Direccion", SqlDbType = System.Data.SqlDbType.VarChar, Value = Proveedor.Direccion});
            parametros.Add(new SqlParameter { ParameterName = "@Email", SqlDbType = System.Data.SqlDbType.VarChar, Value = Proveedor.Email});
            parametros.Add(new SqlParameter { ParameterName = "@RFC", SqlDbType = System.Data.SqlDbType.VarChar, Value = Proveedor.RFC});
            parametros.Add(new SqlParameter { ParameterName = "@PlazoPago", SqlDbType = System.Data.SqlDbType.Int, Value = Proveedor.PlazoPago});
            parametros.Add(new SqlParameter { ParameterName = "@PorcentajeRetencion", SqlDbType = System.Data.SqlDbType.Decimal, Value = Proveedor.PorcentajeRetencion});
            parametros.Add(new SqlParameter { ParameterName = "@UsuarioRegistra", SqlDbType = System.Data.SqlDbType.Int, Value = Proveedor.UsuarioRegistra });


            try 
            {
                DataSet ds = dac.Fill("sp_insert_proveedores", parametros);
                IdProveedor = ds.Tables[0].AsEnumerable().Select(dataRow=>int.Parse(dataRow["IdProveedor"].ToString())).ToList()[0];
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw ex;
            }
            return IdProveedor.ToString();
        }

        public string UpdateProveedor(UpdateProveedorModel Proveedor)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            string mensaje;


            parametros.Add(new SqlParameter { ParameterName = "@IdProveedor", SqlDbType = System.Data.SqlDbType.VarChar, Value = Proveedor.IdProveedor });
            parametros.Add(new SqlParameter { ParameterName = "@Nombre", SqlDbType = System.Data.SqlDbType.VarChar, Value = Proveedor.Nombre});
            parametros.Add(new SqlParameter { ParameterName = "@Direccion", SqlDbType = System.Data.SqlDbType.VarChar, Value = Proveedor.Direccion});
            parametros.Add(new SqlParameter { ParameterName = "@Email", SqlDbType = System.Data.SqlDbType.VarChar, Value = Proveedor.Email});
            parametros.Add(new SqlParameter { ParameterName = "@RFC", SqlDbType = System.Data.SqlDbType.VarChar, Value = Proveedor.RFC});
            parametros.Add(new SqlParameter { ParameterName = "@PlazoPago", SqlDbType = System.Data.SqlDbType.Int, Value = Proveedor.PlazoPago });
            parametros.Add(new SqlParameter { ParameterName = "@PorcentajeRetencion", SqlDbType = System.Data.SqlDbType.Decimal, Value = Proveedor.PorcentajeRetencion });
            parametros.Add(new SqlParameter { ParameterName = "@UsuarioRegistra", SqlDbType = System.Data.SqlDbType.Int, Value = Proveedor.UsuarioRegistra });


            try
            {
                DataSet ds = dac.Fill("sp_update_proveedores", parametros);
                mensaje = ds.Tables[0].AsEnumerable().Select(dataRow => dataRow["mensaje"].ToString()).ToList()[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return mensaje;
        }

        public void DeleteProveedor(int id)
        {
            ConexionDataAccess dac = new ConexionDataAccess(connection);
            parametros = new ArrayList();
            parametros.Add(new SqlParameter { ParameterName = "@IdProveedor", SqlDbType = SqlDbType.Int, Value = id });


            try
            {
                dac.ExecuteNonQuery("sp_delete_proveedores", parametros);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}