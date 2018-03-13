Imports System.Data.SqlClient
Public Class ctiConfiguracion
    ''''''''''''''''Horario
    Public Function actualizarHorarios(ByVal hora As String, ByVal lunes As Boolean, ByVal martes As Boolean, ByVal miercoles As Boolean, ByVal jueves As Boolean, ByVal viernes As Boolean, ByVal sabado As Boolean, ByVal domingo As Boolean) As String
        Dim err As String

        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()

        Dim cmd As New SqlCommand("UPDATE Configuracion SET  hora = '" & hora & "', lunes = '" & lunes & "', martes = '" & martes & "', miercoles = '" & miercoles & "', jueves = '" & jueves & "', viernes = '" & viernes & "' , sabado ='" & sabado & "' , domingo = '" & domingo & "'", dbC)

        cmd.Parameters.AddWithValue("hora", hora)
        cmd.Parameters.AddWithValue("lunes", lunes)
        cmd.Parameters.AddWithValue("martes", martes)
        cmd.Parameters.AddWithValue("miercoles", miercoles)
        cmd.Parameters.AddWithValue("jueves", jueves)
        cmd.Parameters.AddWithValue("viernes", viernes)
        cmd.Parameters.AddWithValue("sabado", sabado)
        cmd.Parameters.AddWithValue("domingo", domingo)


        cmd.ExecuteNonQuery()
        err = "Datos actualizados."

        cmd.Dispose() : dbC.Close() : dbC.Dispose()

        Return err
    End Function
    Public Function datosHorario() As String()
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT top(1) * FROM Config where idconfig = 1", dbC)

        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(8)
            dsP(0) = rdr("hora").ToString
            dsP(1) = rdr("lunes").ToString
            dsP(2) = rdr("martes").ToString
            dsP(3) = rdr("miercoles").ToString
            dsP(4) = rdr("jueves").ToString
            dsP(5) = rdr("viernes").ToString
            dsP(6) = rdr("sabado").ToString
            dsP(7) = rdr("domingo").ToString

        Else
            ReDim dsP(1) : dsP(1) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function



    Public Function actualizarDiaCaptura(ByVal hora As String, ByVal lunes As Boolean, ByVal martes As Boolean, ByVal miercoles As Boolean, ByVal jueves As Boolean, ByVal viernes As Boolean, ByVal sabado As Boolean, ByVal domingo As Boolean) As String
        Dim err As String

        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()

        Dim cmd As New SqlCommand("UPDATE Config SET hora = @hora, lunes = @lunes, martes = @martes, miercoles = @miercoles, jueves = @jueves ,viernes = @viernes ,sabado = @sabado, domingo = @domingo where idconfig = 1", dbC)
        cmd.Parameters.AddWithValue("hora", hora)
        cmd.Parameters.AddWithValue("lunes", lunes)
        cmd.Parameters.AddWithValue("martes", martes)
        cmd.Parameters.AddWithValue("miercoles", miercoles)
        cmd.Parameters.AddWithValue("jueves", jueves)
        cmd.Parameters.AddWithValue("viernes", viernes)
        cmd.Parameters.AddWithValue("sabado", sabado)
        cmd.Parameters.AddWithValue("domingo", domingo)

        cmd.ExecuteNonQuery()
        err = "Datos actualizados."

        cmd.Dispose() : dbC.Close() : dbC.Dispose()

        Return err
    End Function
End Class
