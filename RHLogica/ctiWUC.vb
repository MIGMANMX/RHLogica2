Imports System.Data.SqlClient

Public Class ctiWUC
    Public Function wucBancos() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idbanco, banco FROM Bancos ORDER BY banco", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("banco").ToString, rdr("idbanco").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucBancosOtros() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idefectivo, efectivo FROM Efectivo WHERE tipo = 2 OR tipo = 3 ORDER BY tipo, efectivo", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("efectivo").ToString, rdr("idefectivo").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucClasesCC() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idclase, clave + '-' + clase as cla FROM ClasesCH ORDER BY clase", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("cla").ToString, rdr("idclase").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucClasesG() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idclase, clave + '-' + clase as cla FROM ClasesG ORDER BY clase", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("cla").ToString, rdr("idclase").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucClasesInsumos() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idclase, clase FROM Clases ORDER BY clase", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("clase").ToString, rdr("idclase").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucConceptosReferencia() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idconcep, concep FROM Conc_Ref ORDER BY concep", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("concep").ToString, rdr("idconcep").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucConceptosVales() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idvale, concepto FROM ConceptosVales ORDER BY concepto", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("concepto").ToString, rdr("idvale").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucDescuentos() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT iddescuento, descuento FROM Descuentos ORDER BY descuento", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("descuento").ToString, rdr("iddescuento").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucDiscrepancias() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT iddiscrepancia, discrepancia FROM Discrepancias ORDER BY discrepancia", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("discrepancia").ToString, rdr("iddiscrepancia").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucEmpleados(ByVal activos As Boolean, ByVal idSucursal As Integer) As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("", dbC)
        Dim wre As String = ""
        If activos Then wre = " WHERE activo = 1 "
        If idSucursal > 0 Then
            If wre = "" Then wre = " WHERE idsucursal = @idS " Else wre += " AND idsucursal = @idS "
            cmd.Parameters.AddWithValue("idS", idSucursal)
        End If
        cmd.CommandText = "SELECT idempleado, empleado FROM Empleados " & wre & " ORDER BY empleado"
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("empleado").ToString, rdr("idempleado").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucEmpleados2(ByVal idSucursal As Integer) As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("", dbC)
        Dim wre As String = ""

        If idSucursal > 0 Then
            If wre = "" Then wre = " WHERE idsucursal = @idS " Else wre += " AND idsucursal = @idS "
            cmd.Parameters.AddWithValue("idS", idSucursal)
        End If

        cmd.CommandText = "SELECT idempleado, empleado FROM Empleados " & wre & " AND activo = 1 ORDER BY empleado"
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("empleado").ToString, rdr("idempleado").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucEmpresas() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idempresa, empresa FROM Empresas ORDER BY empresa", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("empresa").ToString, rdr("idempresa").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucFamiliasTiposProducto() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idfam, familia FROM FamiliasVenta ORDER BY familia", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("familia").ToString, rdr("idfam").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucFormasPagoSub() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idfopasub, formapago2 FROM FormaPagoSub ORDER BY idfopasub", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("formapago2").ToString, rdr("idfopasub").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucGruposGasto() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idgrupo, grupo FROM GruposG ORDER BY grupo", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("grupo").ToString, rdr("idgrupo").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucIncidencias() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idincidencia, incidencia FROM Incidencia ORDER BY incidencia", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("incidencia").ToString, rdr("idincidencia").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucInsumos() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idinsumo, insumo AS clvInsumo FROM Insumos WHERE medible = 1 ORDER BY insumo", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("clvInsumo").ToString, rdr("idinsumo").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucOtros() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idotros, concepto FROM Otros ORDER BY concepto", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("concepto").ToString, rdr("idotros").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucProductos() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idproducto, producto FROM Productos ORDER BY producto", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("producto").ToString, rdr("idproducto").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucProductosActivos() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idproducto, producto FROM Productos WHERE activo = 1 ORDER BY producto", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("producto").ToString, rdr("idproducto").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucProductosGrupo() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idgrupo, grupovta FROM GruposVta ORDER BY grupovta", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("grupovta").ToString, rdr("idgrupo").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucProductosFam() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idfam, familia FROM FamiliasVenta ORDER BY familia", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("familia").ToString, rdr("idfam").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucProductosClase() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idclasei, clase_item FROM ClasesItem ORDER BY clase_item", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("clase_item").ToString, rdr("idclasei").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucProveedores() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idproveedor, proveedor FROM Proveedores ORDER BY proveedor", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("proveedor").ToString, rdr("idproveedor").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucProveedoresG() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idproveedor, proveedor FROM ProveedoresG ORDER BY proveedor", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("proveedor").ToString, rdr("idproveedor").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucPuestos() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idpuesto, puesto FROM Puestos ORDER BY puesto", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("puesto").ToString, rdr("idpuesto").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucJornadas() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idjornada, jornada FROM Jornada ORDER BY jornada", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("jornada").ToString, rdr("idjornada").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucSemiprocesados() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idSemiprocesado, semiprocesado FROM SemiProcesados ORDER BY semiprocesado", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("semiprocesado").ToString, rdr("idSemiprocesado").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucSucursales() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idsucursal, sucursal FROM Sucursales ORDER BY sucursal", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("sucursal").ToString, rdr("idsucursal").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucSuc() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idsucursal, sucursal FROM Sucursales ORDER BY sucursal", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("sucursal").ToString, rdr("idsucursal").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucTiposProducto() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idclasei, clase_item FROM TipoProductos ORDER BY clase_item", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("clase_item").ToString, rdr("idclasei").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function wucUnidadMedida() As SortedList
        Dim lista As New SortedList
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idUnidadMedida, unidadMedida FROM UnidadMedida ORDER BY unidadMedida", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            lista.Add(rdr("unidadMedida").ToString, rdr("idUnidadMedida").ToString)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
    Public Function ctrlIncidencias() As String
        Dim lista As String = "<option value='0' selected='selected'>Seleccionar...</option>"
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idincidencia, incidencia FROM Incidencias ORDER BY incidencia", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim ik As Integer = 1
        While rdr.Read
            lista += "<option value='" & rdr("idincidencia").ToString & "' onclick=""incidenciaSel();"">" & rdr("incidencia").ToString & "</option>"
            ik += 1
        End While
        lista = "<select id='lstIncidencias' size='" & ik & "'>" + lista + "</select>"
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return lista
    End Function
End Class
