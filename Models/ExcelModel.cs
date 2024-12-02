using System;
namespace reportesApi.Models
{
public class TraspasoEntradaModel
{
    public string IdTRSP { get; set; }
    public string AlmacenOrigen { get; set; }
    public string NombreAlmacenOrigen { get; set; }
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
public class TraspasoSalidaModel
{
    public string IdTRSP { get; set; }
    public string AlmacenOrigen { get; set; }
    public string NombreAlmacenOrigen { get; set; }
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
public class ContenidoRecetaModel
{
    public string RecetaID { get; set; }
    public string NombreReceta { get; set; }
    public string Insumo { get; set; }
    public string DescripcionInsumo { get; set; }

    public string FechaCreacion { get; set; }
    public string UsuarioRegistra { get; set; }
  
}

public class RecetaRangoFechasModel
{
    public string RecetaID { get; set; }
    public string NombreReceta { get; set; }
    public string FechaCreacion { get; set; }
    public string UsuarioRegistra { get; set; }
  
  
}


}