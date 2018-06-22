Imports System.Data.SqlClient

Public Class ctiVentas
    '''''Ventas
    Public Function datosVentas(ByVal idpuesto As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idpuesto, puesto FROM Puestos WHERE idpuesto = @idP", dbC)
        cmd.Parameters.AddWithValue("idP", idpuesto)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(1)
            dsP(0) = rdr("puesto").ToString
        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra este puesto de empleado."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function actualizarVentas(ByVal idpuesto As Integer,
                                     ByVal puesto As String) As String
        Dim err As String
        If puesto = "" Then
            err = "Error: no se actualizó, es necesario capturar el puesto de empleado."
        Else
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idpuesto FROM Puestos WHERE puesto = @puesto AND idpuesto <> @idP", dbC)
            cmd.Parameters.AddWithValue("puesto", puesto)
            cmd.Parameters.AddWithValue("idP", idpuesto)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                err = "Error: no se actualizó, ya existe."
            Else
                rdr.Close()
                cmd.CommandText = "UPDATE Puestos SET puesto = @puesto WHERE idpuesto = @idP"
                cmd.ExecuteNonQuery()
                err = "Datos del puesto de empleados actualizados."
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        End If
        Return err
    End Function
    Public Function agregarVentas(ByVal idsucursal As Integer,
                                  ByVal Fecha As String,
                                  ByVal VentaN As String,
                                  ByVal IVA As String) As String()
        Dim ans() As String
        If Fecha <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idventas2 FROM Ventas2 WHERE Fecha = @Fecha AND idsucursal = @idsucursal ", dbC)
            cmd.Parameters.AddWithValue("Fecha", Fecha)
            cmd.Parameters.AddWithValue("idsucursal", idsucursal)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe."
                rdr.Close()
            Else
                rdr.Close()
                cmd.CommandText = "INSERT INTO Ventas2 (idsucursal,Fecha,VentaN,IVA,cerrado) 
                    values('" & idsucursal & "','" & Fecha & "','" & VentaN & "','" & IVA & "','0')"
                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT idventas2 FROM Ventas2  WHERE Fecha = @Fecha AND idsucursal = @idsucursal"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Agregado."
                ans(1) = rdr("idventas2").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar."
        End If
        Return ans
    End Function
    Public Function eliminarVentas(ByVal idventas2 As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("DELETE FROM Ventas2 WHERE idventas2 = @idventas2", dbC)
        cmd.Parameters.AddWithValue("idventas2", idventas2)
        Dim err As String
        cmd.ExecuteNonQuery()
        err = "Eliminado."
        cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return err
    End Function
    Public Function gvVentas(ByVal idsucursal As Integer) As DataTable
        Dim dt As New DataTable
        ' Dim a As New int
        dt.Columns.Add(New DataColumn("idventas2", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("Fecha", System.Type.GetType("System.DateTime")))
        dt.Columns.Add(New DataColumn("sucursal", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("VentaN", System.Type.GetType("System.String")))

        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idventas2, Fecha, sucursal, VentaN  FROM vm_Ventas where idsucursal = '" & idsucursal & "' ORDER BY  Fecha DESC ", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow
            r(0) = rdr("idventas2").ToString : r(1) = rdr("Fecha").ToString
            r(2) = rdr("sucursal").ToString : r(3) = rdr("VentaN").ToString
            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dt
    End Function
End Class
