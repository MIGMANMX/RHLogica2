Imports System.Data.SqlClient
Public Class ctiConfiguracion
    ''''''''''''''''Horario
    Public Function actualizarHorarios(ByVal dialimitecaptura As String, ByVal hora As String) As String
        Dim err As String
        'If dialimitecaptura = 0 Then
        '    err = "Error: no se actualizó, es necesario capturar"
        'Else
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()

        Dim cmd As New SqlCommand("UPDATE Configuracion SET dialimitecaptura = @dialimitecaptura, hora = @hora", dbC)
        cmd.Parameters.AddWithValue("dialimitecaptura", dialimitecaptura)
        cmd.Parameters.AddWithValue("hora", hora)

        cmd.ExecuteNonQuery()
        err = "Datos actualizados."

        cmd.Dispose() : dbC.Close() : dbC.Dispose()
        'End If
        Return err
    End Function
    Public Function datosHorario() As String()
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT top(1) * FROM Configuracion", dbC)

        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(2)
            dsP(0) = rdr("dialimitecaptura").ToString
            dsP(1) = rdr("hora").ToString
        Else
            ReDim dsP(1) : dsP(1) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
End Class
