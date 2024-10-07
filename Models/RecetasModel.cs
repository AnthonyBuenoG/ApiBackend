using System;
namespace reportesApi.Models
{
    public class GetRecetasModel
    {
        public int Id { get; set; }
        public string Nombre { get; set; }
        public int Estatus{ get; set; }
        public string FechaCreacion {get; set;}
        public int UsuarioRegista { get; set; }
    }

    public class InsertRecetasModel 
    {
        public string Nombre { get; set; }
        public int Estatus{ get; set; }
        public int UsuarioRegista {get; set;}
       
    }

    public class UpdateRecetasModel
    {
        public int Id { get; set; }
        public string Nombre { get; set; }
        public int Estatus{ get; set; }
        public int UsuarioRegista {get; set;}
       
    }

}