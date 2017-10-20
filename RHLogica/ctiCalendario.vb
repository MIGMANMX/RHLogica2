Imports System.Data.SqlClient

Public Class ctiCalendario
    '
    Public Function datosCalendario() As String()
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT * FROM MostrarCalendario", dbC)
        ' cmd.Parameters.AddWithValue("idE", idEmpleado)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(6)
            dsP(0) = rdr("idempleado").ToString
            dsP(1) = rdr("jornada").ToString
            dsP(2) = rdr("inicio").ToString
            dsP(3) = rdr("fin").ToString
            dsP(4) = rdr("fecha").ToString
            dsP(5) = rdr("color").ToString
        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra este Evento."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function

    '''''Jornada
    Public Function datosJornada(ByVal idjornada As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT * FROM Jornada WHERE idjornada = @idjornada", dbC)
        cmd.Parameters.AddWithValue("idjornada", idjornada)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(6)
            dsP(0) = rdr("idjornada").ToString
            dsP(1) = rdr("jornada").ToString
            dsP(2) = rdr("inicio").ToString
            dsP(3) = rdr("fin").ToString
            dsP(4) = rdr("color").ToString
            dsP(5) = rdr("id_att").ToString
            dsP(6) = rdr("cierre").ToString
        Else
            ReDim dsP(1) : dsP(1) = "Error: no se encuentra esta jornada."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function datosJornada2(ByVal idjornada As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idjornada FROM Jornada WHERE idjornada = @idjornada", dbC)
        cmd.Parameters.AddWithValue("idjornada", idjornada)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(0)
            dsP(0) = rdr("idjornada").ToString
        Else
            ReDim dsP(1) : dsP(1) = "Error: no se encuentra esta jornada."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarJornada(ByVal jornada As String, ByVal inicio As String, ByVal fin As String, ByVal color As String, ByVal id_att As Integer, ByVal cierre As Boolean) As String()
        Dim ans() As String
        If jornada <> "" Then
            Dim dbC As New SqlConnection(StarTconnStrRH)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idjornada FROM Jornada WHERE jornada = @jornada", dbC)
            cmd.Parameters.AddWithValue("jornada", jornada)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe el dato."
                rdr.Close()
            Else
                rdr.Close()
                cmd.CommandText = "INSERT INTO Jornada SELECT @jornada,@inicio,@fin,@color,@id_att,@cierre"

                cmd.Parameters.AddWithValue("cierre", cierre)
                cmd.Parameters.AddWithValue("inicio", inicio)
                cmd.Parameters.AddWithValue("fin", fin)
                cmd.Parameters.AddWithValue("color", color)
                cmd.Parameters.AddWithValue("id_att", id_att)
                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT idjornada FROM Jornada WHERE jornada = @jornada"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Agregado."
                ans(1) = rdr("idjornada").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar."
        End If
        Return ans
    End Function
    Public Function actualizarJornada(ByVal idjornada As Integer, ByVal jornada As String, ByVal inicio As String, ByVal fin As String, ByVal color As String, ByVal id_att As Integer, ByVal cierre As Boolean) As String
        Dim err As String
        If jornada = "" Then
            err = "Error: no se actualizó, es necesario capturar el puesto de empleado."
        Else
            Dim dbC As New SqlConnection(StarTconnStrRH)
            dbC.Open()

            Dim cmd As New SqlCommand("SELECT idjornada FROM Jornada WHERE jornada = @jornada AND idjornada <> @idjornada", dbC)
            cmd.Parameters.AddWithValue("idjornada", idjornada)
            cmd.Parameters.AddWithValue("jornada", jornada)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                err = "Error: no se actualizó, ya existe este nombre."
            Else
                rdr.Close()
                cmd.CommandText = "UPDATE Jornada SET jornada = @jornada, inicio = @inicio, fin = @fin, color = @color, id_att = @id_att, cierre = @cierre WHERE idjornada = @idjornada"
                cmd.Parameters.AddWithValue("cierre", cierre)
                cmd.Parameters.AddWithValue("inicio", inicio)
                cmd.Parameters.AddWithValue("fin", fin)
                cmd.Parameters.AddWithValue("color", color)
                cmd.Parameters.AddWithValue("id_att", id_att)
                cmd.ExecuteNonQuery()
                err = "Datos actualizados."
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        End If
        Return err
    End Function
    Public Function eliminarJornada(ByVal idjornada As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idjornada FROM Jornada WHERE idjornada = @idjornada", dbC)
        cmd.Parameters.AddWithValue("idjornada", idjornada)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim err As String
        If rdr.HasRows Then
            err = "Error: no se puede eliminar."
            rdr.Close()
        Else
            rdr.Close()
            cmd.CommandText = "DELETE FROM Jornada WHERE idjornada = @idjornada"
            cmd.ExecuteNonQuery()
            err = "Jornada eliminada."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return err
    End Function
    Public Function gvJornada() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idjornada", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("jornada", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("inicio", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("fin", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("color", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("cierre", System.Type.GetType("System.Boolean")))


        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idjornada,jornada,inicio,fin,color,cierre FROM Jornada  ORDER BY jornada", dbC)
        'cmd.Parameters.AddWithValue("idjornada", idjornada)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow
            r(0) = rdr("idjornada").ToString
            r(1) = rdr("jornada").ToString
            r(2) = rdr("inicio").ToString
            r(3) = rdr("fin").ToString
            r(4) = rdr("color").ToString
            r(5) = rdr("cierre").ToString

            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dt
    End Function

    '''''Particulares
    Public Function gvParticulares(ByVal idempleado As Integer) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idparticulares", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("tipo", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("fecha", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("observaciones", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("cantidad", System.Type.GetType("System.String")))


        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idparticulares,tipo,fecha,observaciones,cantidad FROM Particulares where idempleado=@idempleado ORDER BY fecha desc", dbC)
        cmd.Parameters.AddWithValue("idempleado", idempleado)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow
            r(0) = rdr("idparticulares").ToString
            r(1) = rdr("tipo").ToString
            r(2) = rdr("fecha").ToString
            r(3) = rdr("observaciones").ToString
            r(4) = rdr("cantidad").ToString
            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dt
    End Function
    Public Function datosParticulares(ByVal idparticulares As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT * FROM Particulares WHERE idparticulares = @idparticulares", dbC)
        cmd.Parameters.AddWithValue("idparticulares", idparticulares)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(6)
            dsP(0) = rdr("idparticulares").ToString
            dsP(1) = rdr("idempleado").ToString
            dsP(2) = rdr("tipo").ToString
            dsP(3) = rdr("fecha").ToString
            dsP(4) = rdr("observaciones").ToString
            dsP(5) = rdr("cantidad").ToString
            dsP(6) = rdr("fecha_ued").ToString

        Else
            ReDim dsP(1) : dsP(1) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarParticulares(ByVal idempleado As Integer, ByVal tipo As String, ByVal fecha As String, ByVal observaciones As String, ByRef cantidad As Integer) As String()
        Dim ans() As String
        If tipo <> "" Then
            Dim dbC As New SqlConnection(StarTconnStrRH)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idparticulares FROM Particulares WHERE tipo = @tipo AND fecha = @fecha", dbC)
            cmd.Parameters.AddWithValue("tipo", tipo)
            cmd.Parameters.AddWithValue("fecha", fecha)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe el dato."
                rdr.Close()
            Else
                rdr.Close()
                ' cmd.CommandText = "INSERT INTO Particulares SELECT @idempleado,@tipo,@fecha,@observaciones,@cantidad"
                cmd.CommandText = "INSERT INTO Particulares (idempleado,tipo,fecha,observaciones,cantidad) 
                    values('" & idempleado & "','" & tipo & "','" & fecha & "','" & observaciones & "','" & cantidad & "')"
                cmd.Parameters.AddWithValue("idempleado", idempleado)
                cmd.Parameters.AddWithValue("observaciones", observaciones)
                cmd.Parameters.AddWithValue("cantidad", cantidad)

                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT idparticulares FROM Particulares WHERE tipo = @tipo AND fecha = @fecha AND idempleado = @idempleado"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Agregado."
                ans(1) = rdr("idparticulares").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar."
        End If
        Return ans
    End Function
    Public Function actualizarParticulares(ByVal idparticulares As Integer, ByVal idempleado As Integer, ByVal tipo As String, ByVal fecha As String, ByVal observaciones As String, ByVal cantidad As Integer, ByVal fecha_ued As String) As String
        Dim err As String
        If tipo = "" Then
            err = "Error: no se actualizó, es necesario capturar"
        Else
            Dim dbC As New SqlConnection(StarTconnStrRH)
            dbC.Open()

            Dim cmd As New SqlCommand("UPDATE Particulares SET tipo = @tipo, fecha = @fecha, observaciones = @observaciones ,cantidad = @cantidad,fecha_ued = @fecha_ued  WHERE idparticulares = @idparticulares", dbC)
            cmd.Parameters.AddWithValue("tipo", tipo)
            cmd.Parameters.AddWithValue("fecha", fecha)
            cmd.Parameters.AddWithValue("idempleado", idempleado)
            cmd.Parameters.AddWithValue("observaciones", observaciones)
            cmd.Parameters.AddWithValue("cantidad", cantidad)
            cmd.Parameters.AddWithValue("fecha_ued", fecha_ued)
            cmd.Parameters.AddWithValue("idparticulares", idparticulares)
            cmd.ExecuteNonQuery()
            err = "Datos actualizados."

            cmd.Dispose() : dbC.Close() : dbC.Dispose()
        End If
        Return err
    End Function
    Public Function eliminarParticulares(ByVal idparticulares As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idparticulares FROM Particulares WHERE idparticulares = @idparticulares", dbC)
        cmd.Parameters.AddWithValue("idparticulares", idparticulares)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim err As String
        If rdr.HasRows Then
            err = "Error: no se puede eliminar."
            rdr.Close()
        Else
            rdr.Close()
            cmd.CommandText = "DELETE FROM Particulares WHERE idparticulares = @idparticulares"
            cmd.ExecuteNonQuery()
            err = "Eliminada."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return err
    End Function
End Class
