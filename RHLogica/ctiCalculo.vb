Imports System.Data.SqlClient
Public Class ctiCalculo
    'Calculo de Horas
    '''Gridview Chequeo
    'Public Function gvChequeo(ByVal idempleado As Integer, ByVal Fech1 As String, ByVal Fech2 As String, ByVal idincidencia As Integer) As DataTable
    '    Dim dt As New DataTable
    '    dt.Columns.Add(New DataColumn("idchequeo", System.Type.GetType("System.String")))
    '    dt.Columns.Add(New DataColumn("chec", System.Type.GetType("System.String")))
    '    dt.Columns.Add(New DataColumn("tipo", System.Type.GetType("System.String")))
    '    dt.Columns.Add(New DataColumn("incidencia", System.Type.GetType("System.String")))
    '    dt.Columns.Add(New DataColumn("observaciones", System.Type.GetType("System.String")))

    '    Dim r As DataRow
    '    Dim dbC As New SqlConnection(StarTconnStrRH)
    '    dbC.Open()
    '    Dim cmd As New SqlCommand("delete from Chequeo where  idchequeo in (select a1.idchequeo from Chequeo a1 inner join Chequeo a2 on a1.chec = a2.chec and a1.idchequeo > a2.idchequeo and a1.idempleado = a2.idempleado)", dbC)

    '    ' dbC.Open()
    '    cmd.ExecuteNonQuery()
    '    cmd.CommandText = "SELECT DISTINCT idchequeo,chec, tipo,incidencia,observaciones FROM vm_ChequeoIncidencia Where chec between '" & Fech1 & "' and '" & Fech2 & "' AND idempleado=@idempleado  ORDER BY chec"
    '    cmd.Parameters.AddWithValue("idempleado", idempleado)
    '    Dim rdr As SqlDataReader = cmd.ExecuteReader
    '    While rdr.Read
    '        r = dt.NewRow
    '        r(0) = rdr("idchequeo").ToString
    '        r(1) = rdr("chec").ToString : r(2) = rdr("tipo").ToString
    '        r(3) = rdr("incidencia").ToString : r(4) = rdr("observaciones").ToString
    '        dt.Rows.Add(r)
    '    End While
    '    rdr.Close() : rdr = Nothing : cmd.Dispose()
    '    dbC.Close() : dbC.Dispose()
    '    Return dt
    'End Function
    Public Function gvChequeo() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("fecha", System.Type.GetType("System.DateTime")))
        dt.Columns.Add(New DataColumn("clockin", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("clockout", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("hrstrab", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("detalle", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("horario", System.Type.GetType("System.String")))
        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT  fecha,clockin, clockout,hrstrab,detalle,horario FROM Temp_Calculo order by fecha asc", dbC)

        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow
            r(0) = rdr("fecha").ToString
            r(1) = rdr("clockin").ToString
            r(2) = rdr("clockout").ToString
            r(3) = rdr("hrstrab").ToString
            r(4) = rdr("detalle").ToString
            r(5) = rdr("horario").ToString
            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose()
        dbC.Close() : dbC.Dispose()
        Return dt
    End Function
    Public Function gvCalculoSucursal() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("empleado", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("fecha", System.Type.GetType("System.DateTime")))
        dt.Columns.Add(New DataColumn("clockin", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("clockout", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("hrstrab", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("detalle", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("horario", System.Type.GetType("System.String")))
        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT DISTINCT empleado,fecha,clockin, clockout,hrstrab,detalle,horario FROM Temp_CalculoSucursal where detalle != '' order by empleado asc,fecha", dbC)
        'cmd.Parameters.AddWithValue("FIn", FIn)
        'cmd.Parameters.AddWithValue("FFn", FFn)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow
            r(0) = rdr("empleado").ToString
            r(1) = rdr("fecha").ToString
            r(2) = rdr("clockin").ToString
            r(3) = rdr("clockout").ToString
            r(4) = rdr("hrstrab").ToString
            r(5) = rdr("detalle").ToString
            r(6) = rdr("horario").ToString
            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose()
        dbC.Close() : dbC.Dispose()
        Return dt
    End Function
    Public Function datosHora(ByVal idchequeo As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("Select Convert(varchar(10),chec, 8)as chec  from Chequeo  where idchequeo=@idchequeo ", dbC)
        cmd.Parameters.AddWithValue("idchequeo", idchequeo)

        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(1)
            dsP(0) = rdr("chec").ToString
        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarHora(ByVal fech As Date, ByVal idempleado As Integer, ByVal horain As Date, ByVal horafn As Date, ByVal horas As TimeSpan) As String()
        Dim ans As String()

        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("INSERT INTO Horas SELECT @idempleado,@horain,@horafn,@horas,@fecha", dbC)
        cmd.Parameters.AddWithValue("fecha", fech)
        cmd.Parameters.AddWithValue("idempleado", idempleado)
        cmd.Parameters.AddWithValue("horain", horain)
        cmd.Parameters.AddWithValue("horafn", horafn)
        cmd.Parameters.AddWithValue("horas", horas)
        cmd.ExecuteNonQuery()
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        cmd.CommandText = "SELECT idHoras FROM Horas WHERE fecha = @fecha"
        rdr = cmd.ExecuteReader
        rdr.Read()
        ReDim ans(1)
        ans(0) = "Agregado."
        ans(1) = rdr("idpuesto").ToString
        rdr.Close()

        'If rdr.HasRows Then
        '    ReDim ans(0)
        '    ans(0) = "Error: no se puede agregar, ya existe."
        '    rdr.Close()
        'Else
        '    rdr.Close()
        '    cmd.CommandText = "INSERT INTO Horas SELECT @idempleado,@horain,@horafn,@horas,@fecha"
        '    cmd.Parameters.AddWithValue("idempleado", idempleado)
        '    cmd.Parameters.AddWithValue("horain", horain)
        '    cmd.Parameters.AddWithValue("horafn", horafn)
        '    cmd.Parameters.AddWithValue("horas", horas)

        '    cmd.ExecuteNonQuery()
        '    cmd.CommandText = "SELECT idHoras FROM Horas WHERE fecha = @fecha"
        '    rdr = cmd.ExecuteReader
        '    rdr.Read()
        '    ReDim ans(1)
        '    ans(0) = "Agregado."
        '    ans(1) = rdr("idpuesto").ToString
        '    rdr.Close()
        'End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()

        Return ans
    End Function
    '''''''Asignar incidencia
    Public Function gvAsigIncidencias(ByVal idincidencia As Integer, ByVal idempleado As Integer, ByVal fecha As String, ByVal observaciones As String) As String()
        Dim ans() As String
        If fecha <> "" And observaciones <> "" And idincidencia <> 0 And idempleado <> 0 Then
            Dim dbC As New SqlConnection(StarTconnStrRH)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idincidencia FROM Detalle_incidencias WHERE fecha = @fecha AND idempleado = @idempleado", dbC)
            cmd.Parameters.AddWithValue("idincidencia", idincidencia)
            cmd.Parameters.AddWithValue("idempleado", idempleado)
            cmd.Parameters.AddWithValue("fecha", Convert.ToDateTime(fecha))
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe registro a este nombre."
                rdr.Close()
                cmd.CommandText = "UPDATE Detalle_incidencias SET idincidencia = @idincidencia, idempleado = @idempleado, fecha = @fecha, observaciones = @observaciones  WHERE iddetalle_incidencia = @iddetalle_incidencia"
                cmd.ExecuteNonQuery()
                ans(0) = "Datos de incidencia actualizados."
            Else
                rdr.Close()
                cmd.CommandText = "INSERT INTO Detalle_incidencias SELECT @idincidencia, @idempleado,@fecha, @observaciones"
                cmd.Parameters.AddWithValue("observaciones", observaciones)

                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT iddetalle_incidencia FROM Detalle_incidencias WHERE idempleado = @idempleado and fecha = @fecha"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Incidencia agregada."
                ans(1) = rdr("iddetalle_incidencia").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar."
        End If
        Return ans
    End Function
    '''''Actualizar Incidencias en Chequeo
    Public Function actualizarIncidencias(ByVal idchequeo As Integer, ByVal idincidencia As Integer, ByVal observaciones As String) As String
        Dim err As String
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idchequeo FROM Chequeo", dbC)
        cmd.Parameters.AddWithValue("idchequeo", idchequeo)
        cmd.Parameters.AddWithValue("idincidencia", idincidencia)
        cmd.Parameters.AddWithValue("observaciones", observaciones)

        Dim rdr As SqlDataReader = cmd.ExecuteReader

        rdr.Close()
        cmd.CommandText = "UPDATE Chequeo SET idincidencia = @idincidencia, observaciones = @observaciones  WHERE idchequeo = @idchequeo"
        cmd.ExecuteNonQuery()
        err = "Datos actualizados."

        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        'End If
        Return err
    End Function
    '''''Datos Incidencias en Chequeo
    Public Function datosCheqIncidencias(ByVal idchequeo As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idincidencia, observaciones, idchequeo FROM Chequeo WHERE idchequeo = @idchequeo", dbC)
        cmd.Parameters.AddWithValue("idchequeo", idchequeo)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(5)
            dsP(0) = rdr("idincidencia").ToString
            dsP(1) = rdr("observaciones").ToString
            dsP(2) = rdr("idchequeo").ToString
        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra ."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    '''''Consuta si asistio con Chequeo
    Public Function ConsultaAsistencia(ByVal idempleado As Integer, ByVal F As String, ByVal FF As String) As String()
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("Select * From Chequeo where chec  BETWEEN '" & F & "' AND '" & FF & "' AND idempleado=@idempleado AND tipo='Entrada' Order BY chec ", dbC)
        cmd.Parameters.AddWithValue("idempleado", idempleado)

        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(4)
            dsP(0) = rdr("idchequeo").ToString
            dsP(1) = rdr("idempleado").ToString
            dsP(2) = rdr("chec").ToString
            dsP(3) = rdr("tipo").ToString
        Else
            ReDim dsP(0) : dsP(0) = 0
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function

    '''''''Salarios


End Class
