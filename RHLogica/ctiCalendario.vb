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
    Public Function agregarJornada(ByVal jornada As String, ByVal inicio As String, ByVal fin As String, ByVal color As String, ByRef id_att As Integer) As String()
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
                cmd.CommandText = "INSERT INTO Jornada SELECT @jornada,@inicio,@fin,@color,@id_att"

                'cmd.Parameters.AddWithValue("jornada", jornada)
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
    Public Function actualizarJornada(ByVal idjornada As Integer, ByVal jornada As String, ByVal inicio As String, ByVal fin As String, ByVal color As String, ByVal id_att As Integer) As String
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
                cmd.CommandText = "UPDATE Jornada SET jornada = @jornada, inicio = @inicio, fin = @fin, color = @color, id_att = @id_att WHERE idjornada = @idjornada"
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
        'dt.Columns.Add(New DataColumn("id_att", System.Type.GetType("System.Int32")))


        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idjornada,jornada,inicio,fin,color,id_att FROM Jornada  ORDER BY jornada", dbC)
        'cmd.Parameters.AddWithValue("idjornada", idjornada)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow
            r(0) = rdr("idjornada").ToString
            r(1) = rdr("jornada").ToString
            r(2) = rdr("inicio").ToString
            r(3) = rdr("fin").ToString
            r(4) = rdr("color").ToString
            'r(5) = rdr("id_att").ToString

            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dt
    End Function


End Class
