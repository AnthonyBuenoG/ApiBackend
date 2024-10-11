using System;
using System.Text;
namespace reportesApi.Models
{
    public class GetDetalleRecetaModel
    {
        public int Id { get; set; }
        public int IdReceta { get; set; }

        public string Insumo{ get; set; }
        public decimal Cantidad {get; set;}
        public string UsuarioRegistra { get; set; }


       
    }

    public class InsertDetalleRecetaModel 
    {
        public int IdReceta { get; set; }
        public string Insumo{ get; set; }
        public decimal Cantidad {get; set;}
        public int UsuarioRegistra {get; set;}
       
    }

    public class UpdateDetalleRecetaModel
    {
        public int Id { get; set; }
        public int IdReceta { get; set; }
        public string Insumo { get; set; }
        public decimal Cantidad {get; set;}
        public int UsuarioRegistra {get; set;}
       
    }

}