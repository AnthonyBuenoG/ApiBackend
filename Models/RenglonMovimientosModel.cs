using System;
namespace reportesApi.Models
{
    public class GetRenglonMovimientosModel
    {
        public int Id {get; set;}
        public int IdMovimiento { get; set; }
        public string Nombre { get; set; }
        public string Insumo { get; set; }
        public string DescripcionInsumo { get; set; }
        public decimal Cantidad { get; set; }
        public decimal Costo { get; set; }
        public string CostoTotal { get; set; }

        public string Estatus { get; set; }
        public string FechaRegistro { get; set; }
        public string UsuarioRegistra { get; set; }

    }

    public class InsertRenglonMovimientosModel 
    {
        public int IdMovimiento { get; set; }
        public string Insumo { get; set; }
        public decimal Cantidad { get; set; }
        public decimal Costo { get; set; }
        public int UsuarioRegistra { get; set; }
       
    }

    public class UpdateRenglonMovimientosModel
    {
        public int Id { get; set;}
        public int IdMovimiento { get; set; }
        public string Insumo { get; set; }
        public decimal Cantidad { get; set; }
        public decimal Costo { get; set; }
        public int UsuarioRegistra { get; set; }

    }

}