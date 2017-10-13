Imports System.Data.SqlClient

Public Class ctiAdmin
    Public Function ingresar(ByVal usr As String, ByVal clv As String) As String
        Dim ci As String = "0,,,,,"
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim fecha As Date = DateTime.Now
        Dim cmd As New SqlCommand("SELECT idusuario, usuario, nivel, Usuarios.idsucursal, sucursal FROM Usuarios INNER JOIN Sucursales ON Usuarios.idsucursal = Sucursales.idsucursal WHERE usuario=@usuario AND clave=@clave", dbC)
        cmd.Parameters.AddWithValue("usuario", usr)
        cmd.Parameters.AddWithValue("clave", clv)
        Dim rdr As SqlDataReader = cmd.ExecuteReader()
        If rdr.Read Then
            ci = rdr("idusuario").ToString & "," & rdr("usuario").ToString & "," & rdr("nivel").ToString & "," & rdr("idsucursal").ToString & "," & rdr("sucursal").ToString
            rdr.Close()
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        'Dim fecha As Date = DateTime.Now
        'cmd2 = New SqlCommand("insert into Bitacora(usuario,fecha) values ('" & ingreso.Split(",")(1) & "','" & fecha & "')", dbC2)
        'cmd2.ExecuteNonQuery()
        'cmd2.Dispose()
        Return ci
    End Function
End Class
