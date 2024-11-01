using System;
namespace reportesApi.Models
{
    public class GetExistenciasModel
    {
        public int Id { get; set; }
        public string Fecha { get; set; }
        public string Insumo{ get; set; }
        public string DescripcionInsumo {get; set;}
        public string Cantidad { get; set; }
        public int IdAlmacen { get; set; }
        public string Estatus { get; set; }
        public string FechaRegistro { get; set; }
        public string UsuarioRegistra { get; set; }

    }

    public class InsertExistenciasModel 
    {
  
        public string Insumo{ get; set; }
        public decimal Cantidad { get; set; }
        public int IdAlmacen { get; set; }
        public int UsuarioRegistra { get; set; }
       
    }

    public class UpdateExistenciasModel
    {
        public int Id { get; set; }
        public string Insumo{ get; set; }
        public decimal Cantidad { get; set; }
        public int IdAlmacen { get; set; }
        public int UsuarioRegistra { get; set; }

    }

}