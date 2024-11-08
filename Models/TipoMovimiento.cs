using System;
namespace reportesApi.Models
{
    public class GetTipoMovimientoModel
    {
        public int Id {get; set;}
        public string Nombre { get; set; }
        public string EntradaSalida { get; set; }
        public string Estatus { get; set; }
        public string FechaRegistro { get; set; }
        public string UsuarioRegistra { get; set; }

    }

    public class InsertTipoMovimientoModel 
    {
        public string Nombre { get; set; }
        public int EntradaSalida { get; set; }
        public int UsuarioRegistra { get; set; }
       
    }

    public class UpdateTipoMovimientoModel
    {
        public int Id { get; set; }
        public string Nombre { get; set; }
        public int EntradaSalida { get; set; }
        public int UsuarioRegistra { get; set; }

    }

}