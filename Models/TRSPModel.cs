using System;
namespace reportesApi.Models
{
    public class GetTRSPModel
    {
        public string IdTRSP { get; set; }
        public string AlmacenOrigen { get; set; }
        public string NombreAlmacenOrgien { get; set; }
        public string AlmacenDestino { get; set; }
        public string NombreAlmacenDestino { get; set; }
        public string IdInsumo { get; set; }
        public string DescripcionInsumo { get; set; }
        public string FechaEntrada { get; set; }
        public string FechaSalida { get; set; }
        public string Cantidad { get; set; }
        public string TipoMovimiento { get; set; }
        public string Descripcion { get; set; }
        public string NoFolio { get; set; }
        public string CantidadMovimientoOrigen { get; set; }
        public string CantidadMovimientoDestino { get; set; }
        public string FechaRegistro { get; set; }
        public string Estatus { get; set; }
        public string UsuarioRegistra { get; set; }
    }

        public class GetEntradaSalidaTRSPPModel
    {       
         public string IdTRSP { get; set; }
        public string AlmacenOrigen { get; set; }
        public string NombreAlmacenOrgien { get; set; }
        public string AlmacenDestino { get; set; }
        public string NombreAlmacenDestino { get; set; }
        public string IdInsumo { get; set; }
        public string DescripcionInsumo { get; set; }
        public string FechaEntrada { get; set; }
        public string FechaSalida { get; set; }
        public string Cantidad { get; set; }
        public string TipoMovimiento { get; set; }
        public string Descripcion { get; set; }
        public string NoFolio { get; set; }
        public string FechaRegistro { get; set; }
        public string Estatus { get; set; }
        public string UsuarioRegistra { get; set; }
    }


        public class GetReporteTRSPModel
    {
        public string IdTRSP { get; set; }
        public string AlmacenOrigen { get; set; }
        public string NombreAlmacenOrgien { get; set; }
        public string AlmacenDestino { get; set; }
        public string NombreAlmacenDestino { get; set; }
        public string IdInsumo { get; set; }
        public string DescripcionInsumo { get; set; }
        public string FechaEntrada { get; set; }
        public string FechaSalida { get; set; }
        public string Cantidad { get; set; }
        public string TipoMovimiento { get; set; }
        public string Descripcion { get; set; }
        public string NoFolio { get; set; }
        public string CantidadMovimientoOrigen { get; set; }
        public string CantidadMovimientoDestino { get; set; }
        public string FechaRegistro { get; set; }
        public string Estatus { get; set; }
        public string UsuarioRegistra { get; set; }
    }
    
        public class GetTRSPRenglonModel
    {       
         public string IdTRSP { get; set; }
        public string AlmacenOrigen { get; set; }
        public string NombreAlmacenOrgien { get; set; }
        public string AlmacenDestino { get; set; }
        public string NombreAlmacenDestino { get; set; }
        public string IdInsumo { get; set; }
        public string DescripcionInsumo { get; set; }
        public string FechaEntrada { get; set; }
        public string FechaSalida { get; set; }
        public string Cantidad { get; set; }
        public string TipoMovimiento { get; set; }
        public string Descripcion { get; set; }
        public string NoFolio { get; set; }
        public string FechaRegistro { get; set; }
        public string Estatus { get; set; }
        public string UsuarioRegistra { get; set; }
    }

       public class GetTRSPOrgModel
    {       
         public string IdTRSP { get; set; }
        public string AlmacenOrigen { get; set; }
        public string NombreAlmacenOrgien { get; set; }
        public string AlmacenDestino { get; set; }
        public string NombreAlmacenDestino { get; set; }
        public string IdInsumo { get; set; }
        public string DescripcionInsumo { get; set; }
        public string FechaEntrada { get; set; }
        public string FechaSalida { get; set; }
        public string Cantidad { get; set; }
        public string TipoMovimiento { get; set; }
        public string Descripcion { get; set; }
        public string NoFolio { get; set; }
        public string CantidadMovimientoOrigen { get; set; }
        public string CantidadMovimientoDestino { get; set; }
        public string FechaRegistro { get; set; }
        public string Estatus { get; set; }
        public string UsuarioRegistra { get; set; }
    }





    public class InsertTRSPModel 
    {
        public int AlmacenDestino { get; set; }

        public int AlmacenOrigen { get; set; }
        public int IdInsumo {get; set; }
        public int Cantidad { get; set; }
        public int TipoMovimiento {get; set; }
        public string Descripcion { get; set; }
        public int UsuarioRegistra {get; set; }

    }

        public class InsertTRSPRenglonModel 
    {
        public int AlmacenOrigen { get; set; }
        public int AlmacenDestino { get; set; }
        public int IdInsumo {get; set; }
        public int Cantidad { get; set; }
        public int TipoMovimiento {get; set; }
        public string Descripcion { get; set; }
        public int UsuarioRegistra {get; set; }

    }

    public class UpdateTRSPModel
    {
        public int IdTRSP { get; set; }
        public int AlmacenOrigen { get; set; }
        public int AlmacenDestino { get; set; }
        public int IdInsumo {get; set;}
        public int Cantidad { get; set;}
        public int TipoMovimiento {get; set;}
        public string Descripcion { get; set;}
        public int Estatus { get; set; }
        public int UsuarioRegistra {get; set;}

    }

}