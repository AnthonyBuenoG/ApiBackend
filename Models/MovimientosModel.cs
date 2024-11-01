using System;
namespace reportesApi.Models
{
    public class GetMovimientosModel
    {
        public int Id {get; set;}
        public int IdTipoMovimiento { get; set; }
        public int IdAlmacen { get; set; }
        public string Fecha { get; set; }
        public string Estatus { get; set; }
        public string IdUsuario { get; set; }
        public string FechaRegistro { get; set; }

    }

    public class InsertMovimientosModel 
    {
        public int IdTipoMovimiento { get; set; }
        public int IdAlmacen { get; set; }
        public string IdUsuario { get; set; }
       
    }

    public class UpdateMovimientosModel
    {
        public int Id {get; set;}
        public int IdTipoMovimiento { get; set; }
        public int IdAlmacen { get; set; }
        public string IdUsuario { get; set; }

    }

}