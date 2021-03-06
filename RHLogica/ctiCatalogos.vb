﻿Imports System.Data.SqlClient

Public Class ctiCatalogos
    '''''Puestos
    Public Function datosPuesto(ByVal idpuesto As Integer) As String()
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
    Public Function agregarPuesto(ByVal puesto As String) As String()
        Dim ans() As String
        If puesto <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idpuesto FROM Puestos WHERE puesto = @puesto", dbC)
            cmd.Parameters.AddWithValue("puesto", puesto)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe un puesto de empleado con este nombre."
                rdr.Close()
            Else
                rdr.Close()
                cmd.CommandText = "INSERT INTO Puestos SELECT @puesto"
                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT idpuesto FROM Puestos WHERE puesto = @puesto"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Puesto de empleados agregado."
                ans(1) = rdr("idpuesto").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar el puesto de empleado."
        End If
        Return ans
    End Function
    Public Function actualizarPuesto(ByVal idpuesto As Integer,
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
    Public Function eliminarPuesto(ByVal idpuesto As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idpuesto FROM Empleados WHERE idpuesto = @idP", dbC)
        cmd.Parameters.AddWithValue("idP", idpuesto)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim err As String
        If rdr.HasRows Then
            err = "Error: este puesto de empleados no se puede eliminar, tiene empleados asociadas."
            rdr.Close()
        Else
            rdr.Close()
            cmd.CommandText = "DELETE FROM Puestos WHERE idpuesto = @idP"
            cmd.ExecuteNonQuery()
            err = "Puesto de empleados eliminado."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return err
    End Function
    Public Function gvPuesto() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idpuesto", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("puesto", System.Type.GetType("System.String")))
        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idpuesto, puesto FROM Puestos ORDER BY puesto", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow
            r(0) = rdr("idpuesto").ToString : r(1) = rdr("puesto").ToString
            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dt
    End Function
    ''''''''''Usuarios
    Public Function datosUsuario(ByVal idUsuario As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT nombre, usuario, clave, nivel, idsucursal FROM Usuarios WHERE idusuario = @idU", dbC)
        cmd.Parameters.AddWithValue("idU", idUsuario)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(4)
            dsP(0) = rdr("nombre").ToString
            dsP(1) = rdr("usuario").ToString
            dsP(2) = rdr("clave").ToString
            dsP(3) = rdr("nivel").ToString
            dsP(4) = rdr("idsucursal").ToString
        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra este usuario."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP

    End Function
    Public Function datosUsuarioV(ByVal idUsuario As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT  nivel, idsucursal FROM Usuarios WHERE idusuario = @idU", dbC)
        cmd.Parameters.AddWithValue("idU", idUsuario)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(1)
            dsP(0) = rdr("nivel").ToString
            dsP(1) = rdr("idsucursal").ToString
        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra este usuario."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP

    End Function
    Public Function agregarUsuario(ByVal nombre As String,
                                   ByVal usuario As String,
                                   ByVal clave As String,
                                   ByVal nivel As Integer,
                                   ByVal idSucursal As Integer) As String()
        Dim au() As String
        If nombre <> "" And usuario <> "" And clave <> "" Then
            If nivel > 0 And nivel < 8 Then
                Dim dbC As New SqlConnection(StarTconnStr)
                dbC.Open()
                Dim cmd As New SqlCommand("SELECT sucursal FROM Sucursales WHERE idsucursal = @ids", dbC)
                cmd.Parameters.AddWithValue("ids", idSucursal)
                Dim rdr As SqlDataReader = cmd.ExecuteReader
                If rdr.HasRows Then
                    rdr.Close()
                    cmd.CommandText = "SELECT idusuario FROM Usuarios WHERE usuario = @usuario"
                    cmd.Parameters.AddWithValue("usuario", usuario)
                    rdr = cmd.ExecuteReader
                    If rdr.HasRows Then
                        ReDim au(0)
                        au(0) = "Error: no se puede agregar, ya existe este usuario."
                        rdr.Close()
                    Else
                        rdr.Close()
                        'cmd.CommandText = "INSERT INTO Usuarios values @nombre, @usuario, @clave, @nivel, @ids"
                        cmd.CommandText = "INSERT INTO Usuarios (nombre,usuario,clave,nivel,idsucursal) values('" & nombre & "','" & usuario & "','" & clave & "','" & nivel & "','" & idSucursal & "')"

                        'cmd.Parameters.AddWithValue("nombre", nombre)
                        'cmd.Parameters.AddWithValue("clave", clave)
                        'cmd.Parameters.AddWithValue("nivel", nivel)
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "SELECT idusuario FROM Usuarios WHERE usuario = @usuario"
                        rdr = cmd.ExecuteReader
                        rdr.Read()
                        ReDim au(1)
                        au(0) = "Usuario agregado."
                        au(1) = rdr("idusuario").ToString
                        rdr.Close()
                    End If
                Else
                    ReDim au(0)
                    au(0) = "Error: no se puede agregar, es necesario seleccionar la sucursal."
                    rdr.Close()
                End If
                rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
            Else
                ReDim au(0)
                au(0) = "Error: no se puede agregar, el nivel debe ser entre 1 y 7."
            End If
        Else
            ReDim au(0)
            au(0) = "Error: no se puede agregar, es necesario capturar el nombre, usuario y clave."
        End If
        Return au
    End Function
    Public Function actualizarUsuario(ByVal idUsuario As Integer,
                                      ByVal nombre As String,
                                      ByVal usuario As String,
                                      ByVal clave As String,
                                      ByVal nivel As Integer,
                                      ByVal idSucursal As Integer) As String
        Dim aci As String
        If nombre <> "" And usuario <> "" And clave <> "" Then
            If nivel > 0 And nivel < 8 Then
                Dim dbC As New SqlConnection(StarTconnStr)
                dbC.Open()
                Dim cmd As New SqlCommand("SELECT sucursal FROM Sucursales WHERE idsucursal = @ids", dbC)
                cmd.Parameters.AddWithValue("ids", idSucursal)
                Dim rdr As SqlDataReader = cmd.ExecuteReader
                If rdr.HasRows Then
                    rdr.Close()
                    cmd.CommandText = "SELECT idusuario FROM Usuarios WHERE usuario = @usuario AND idusuario <> @idU"
                    cmd.Parameters.AddWithValue("usuario", usuario)
                    cmd.Parameters.AddWithValue("idU", idUsuario)
                    rdr = cmd.ExecuteReader
                    If rdr.HasRows Then
                        aci = "Error: no se actualizó, ya existe este usuario."
                        rdr.Close()
                    Else
                        rdr.Close()
                        cmd.CommandText = "UPDATE Usuarios SET nombre = @nombre, usuario = @usuario, clave = @clave, nivel = @nivel, idsucursal = @ids WHERE idusuario = @idU"
                        cmd.Parameters.AddWithValue("nombre", nombre)
                        cmd.Parameters.AddWithValue("clave", clave)
                        cmd.Parameters.AddWithValue("nivel", nivel)
                        cmd.ExecuteNonQuery()
                        aci = "Datos del usuario actualizados."
                    End If
                Else
                    aci = "Error: no se actualizó, es necesario seleccionar la sucursal."
                    rdr.Close()
                End If
                rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
            Else
                aci = "Error: no se actualizó, el nivel debe ser entre 1 y 7."
            End If
        Else
            aci = "Error: no se actualizó, es necesario capturar el nombre, usuario y clave."
        End If
        Return aci
    End Function
    Public Function eliminarUsuario(ByVal idUsuario As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("DELETE FROM Usuarios WHERE idusuario = @idU", dbC)
        cmd.Parameters.AddWithValue("idU", idUsuario)
        cmd.ExecuteNonQuery()
        cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return "Usuario eliminado."
    End Function
    Public Function gvUsuarios(ByVal idSucursal As Integer) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idusuario", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("nombre", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("usuario", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("nivel", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("sucursal", System.Type.GetType("System.String")))
        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idusuario,nombre,usuario,nivel,sucursal FROM Vista_Suc_ WHERE idsucursal=@idS ORDER BY sucursal", dbC)
        cmd.Parameters.AddWithValue("idS", idSucursal)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow
            r(0) = rdr("idusuario").ToString
            r(1) = rdr("nombre").ToString
            r(2) = rdr("usuario").ToString
            r(3) = rdr("nivel").ToString
            r(4) = rdr("sucursal").ToString
            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dt
    End Function
    'Empleados
    Public Function datosEmpleado(ByVal idEmpleado As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT empleado, idsucursal, idpuesto, activo, nss, fecha_ingreso, rfc, fecha_nacimiento, calle, numero, colonia, cp, telefono, correo, fecha_baja, idempleado, clave_att, idtipojornada, baja, curp, cnombre, ctelefono, bnota FROM Empleados WHERE idempleado = @idE", dbC)
        cmd.Parameters.AddWithValue("idE", idEmpleado)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(23)
            dsP(0) = rdr("empleado").ToString
            dsP(1) = rdr("idsucursal").ToString
            dsP(2) = rdr("idpuesto").ToString
            dsP(3) = rdr("activo").ToString
            dsP(4) = rdr("nss").ToString
            dsP(5) = rdr("fecha_ingreso").ToString
            dsP(6) = rdr("rfc").ToString
            dsP(7) = rdr("fecha_nacimiento").ToString
            dsP(8) = rdr("calle").ToString
            dsP(9) = rdr("numero").ToString

            dsP(10) = rdr("colonia").ToString
            dsP(11) = rdr("cp").ToString
            dsP(12) = rdr("telefono").ToString
            dsP(13) = rdr("correo").ToString
            dsP(14) = rdr("fecha_baja").ToString
            dsP(15) = rdr("idempleado").ToString
            dsP(16) = rdr("clave_att").ToString

            dsP(17) = rdr("idtipojornada").ToString
            dsP(18) = rdr("baja").ToString

            dsP(19) = rdr("curp").ToString
            dsP(20) = rdr("cnombre").ToString
            dsP(21) = rdr("ctelefono").ToString
            dsP(22) = rdr("bnota").ToString
        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra este empleado."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function clave_att() As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()

        Dim cmd As New SqlCommand("Select top 1 clave_att from Empleados order by idempleado desc", dbC)
        'cmd.Parameters.AddWithValue("idE", idEmpleado)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(1)

            dsP(0) = rdr("clave_att").ToString
        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarEmpleado(ByVal empleado As String,
                                    ByVal idsucursal As Integer,
                                    ByVal activo As Boolean,
                                    ByVal nss As String,
                                    ByVal fecha_ingreso As String,
                                    ByVal rfc As String,
                                    ByVal fecha_nacimiento As String,
                                    ByVal calle As String,
                                    ByVal numero As String,
                                    ByVal colonia As String,
                                    ByVal cp As String,
                                    ByVal telefono As String,
                                    ByVal correo As String, ByVal idpuesto As Integer,
                                    ByVal clave_att As String, ByVal idtipojornada As Integer,
                                    ByVal baja As Boolean,
                                    ByVal curp As String,
                                    ByVal cnombre As String,
                                    ByVal ctelefono As String,
                                    ByVal bnota As String, ByVal fecha_baja As String) As String()
        Dim ae() As String
        If empleado <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT sucursal FROM Sucursales WHERE idsucursal = @idsucursal", dbC)
            cmd.Parameters.AddWithValue("idsucursal", idsucursal)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                cmd.CommandText = "SELECT idempleado FROM Empleados WHERE empleado = @empleado"
                cmd.Parameters.AddWithValue("empleado", empleado)
                rdr = cmd.ExecuteReader
                If rdr.HasRows Then
                    ReDim ae(0)
                    ae(0) = "Error: no se puede agregar, ya existe este empleado."
                    rdr.Close()
                Else
                    rdr.Close()
                    'cmd.CommandText = "INSERT INTO Empleados SELECT @empleado, @idsucursal, @activo"
                    cmd.CommandText = "INSERT INTO Empleados (empleado,idsucursal,activo,nss,fecha_ingreso,rfc,fecha_nacimiento,calle,numero,colonia,cp,telefono,correo,idpuesto, clave_att, idtipojornada,baja,curp, cnombre, ctelefono, bnota, fecha_baja) 
                    values('" & empleado & "','" & idsucursal & "','" & activo & "','" & nss & "','" & fecha_ingreso & "','" & rfc & "','" & fecha_nacimiento & "','" & calle & "','" & numero & "','" & colonia & "','" & cp & "','" & telefono & "','" & correo & "','" & idpuesto & "','" & clave_att & "','" & idtipojornada & "','" & baja & "','" & curp & "','" & cnombre & "','" & ctelefono & "','" & bnota & "','" & fecha_baja & "')"

                    'cmd.Parameters.AddWithValue("idpuesto", idpuesto)
                    cmd.Parameters.AddWithValue("activo", activo)
                    cmd.ExecuteNonQuery()
                    cmd.CommandText = "SELECT idempleado FROM Empleados WHERE empleado = @empleado"
                    rdr = cmd.ExecuteReader
                    rdr.Read()
                    ReDim ae(1)
                    ae(0) = "Empleado agregado."
                    ae(1) = rdr("idempleado").ToString
                    rdr.Close()
                End If
            Else
                ReDim ae(0)
                ae(0) = "Error: no se puede agregar, es necesario seleccionar la sucursal."
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ae(0)
            ae(0) = "Error: no se puede agregar, es necesario capturar el nombre del empleado."
        End If
        Return ae
    End Function
    Public Function actualizarDirectorio(ByVal idEmpleado As Integer,
                                        ByVal nombre As String,
                                        ByVal idSucursal As Integer,
                                        ByVal calle As String,
                                        ByVal numero As String,
                                        ByVal colonia As String,
                                        ByVal cp As String,
                                        ByVal telefono As String,
                                        ByVal correo As String, ByVal clave_att As String,
                                        ByVal cnombre As String,
                                        ByVal ctelefono As String) As String
        Dim aci As String
        If nombre <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT sucursal FROM Sucursales WHERE idsucursal = @idS", dbC)
            cmd.Parameters.AddWithValue("idS", idSucursal)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                cmd.CommandText = "SELECT idempleado FROM Empleados WHERE empleado = @nombre AND idempleado <> @idE"
                cmd.Parameters.AddWithValue("nombre", nombre)
                cmd.Parameters.AddWithValue("idE", idEmpleado)
                rdr = cmd.ExecuteReader
                If rdr.HasRows Then
                    aci = "Error: no se actualizó, ya existe este empleado."
                    rdr.Close()
                Else
                    rdr.Close()
                    cmd.CommandText = "UPDATE Empleados SET calle = @calle, numero = @numero, colonia = @colonia, cp = @cp, telefono = @telefono, correo = @correo, clave_att = @clave_att, cnombre = @cnombre, ctelefono = @ctelefono  WHERE idempleado = @idE"
                    cmd.Parameters.AddWithValue("calle", calle)
                    cmd.Parameters.AddWithValue("numero", numero)
                    cmd.Parameters.AddWithValue("colonia", colonia)
                    cmd.Parameters.AddWithValue("cp", cp)
                    cmd.Parameters.AddWithValue("telefono", telefono)
                    cmd.Parameters.AddWithValue("correo", correo)
                    cmd.Parameters.AddWithValue("clave_att", clave_att)

                    cmd.Parameters.AddWithValue("cnombre", cnombre)
                    cmd.Parameters.AddWithValue("ctelefono", ctelefono)

                    cmd.ExecuteNonQuery()
                    aci = "Datos del empleado actualizados."
                End If
            Else
                aci = "Error: no se actualizó, es necesario seleccionar la sucursal."
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            aci = "Error: no se actualizó, es necesario capturar el nombre del empleado."
        End If
        Return aci
    End Function
    Public Function actualizarEmpleado(ByVal idEmpleado As Integer,
                                        ByVal nombre As String,
                                        ByVal idSucursal As Integer,
                                        ByVal idpuesto As String,
                                        ByVal activo As Boolean,
                                        ByVal nss As String,
                                        ByVal fecha_ingreso As String,
                                        ByVal rfc As String,
                                        ByVal fecha_nacimiento As String,
                                        ByVal calle As String,
                                        ByVal numero As String,
                                        ByVal colonia As String,
                                        ByVal cp As String,
                                        ByVal telefono As String,
                                        ByVal correo As String,
                                        ByVal fecha_baja As String, ByVal idtipojornada As Integer, ByVal baja As Boolean,
                                        ByVal curp As String,
                                        ByVal cnombre As String,
                                        ByVal ctelefono As String,
                                        ByVal bnota As String) As String
        Dim aci As String
        If nombre <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT sucursal FROM Sucursales WHERE idsucursal = @idS", dbC)
            cmd.Parameters.AddWithValue("idS", idSucursal)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                cmd.CommandText = "SELECT idempleado FROM Empleados WHERE empleado = @nombre AND idempleado <> @idE"
                cmd.Parameters.AddWithValue("nombre", nombre)
                cmd.Parameters.AddWithValue("idE", idEmpleado)
                rdr = cmd.ExecuteReader
                If rdr.HasRows Then
                    aci = "Error: no se actualizó, ya existe este empleado."
                    rdr.Close()
                Else
                    rdr.Close()
                    cmd.CommandText = "UPDATE Empleados SET empleado = @nombre, idsucursal = @idS, idpuesto = @idpuesto, activo = @activo ,nss = @nss, fecha_ingreso = @fecha_ingreso, rfc = @rfc, fecha_nacimiento =  @fecha_nacimiento, calle = @calle, numero = @numero, colonia = @colonia, cp = @cp, telefono = @telefono, correo = @correo, fecha_baja = @fecha_baja,idtipojornada = @idtipojornada, baja = @baja, curp = @curp, cnombre = @cnombre, ctelefono = @ctelefono, bnota = @bnota WHERE idempleado = @idE"
                    cmd.Parameters.AddWithValue("idpuesto", idpuesto)
                    cmd.Parameters.AddWithValue("activo", activo)
                    cmd.Parameters.AddWithValue("nss", nss)
                    cmd.Parameters.AddWithValue("fecha_ingreso", Convert.ToDateTime(fecha_ingreso))
                    cmd.Parameters.AddWithValue("rfc", rfc)
                    cmd.Parameters.AddWithValue("fecha_nacimiento", Convert.ToDateTime(fecha_nacimiento))
                    cmd.Parameters.AddWithValue("calle", calle)
                    cmd.Parameters.AddWithValue("numero", numero)
                    cmd.Parameters.AddWithValue("colonia", colonia)
                    cmd.Parameters.AddWithValue("cp", cp)
                    cmd.Parameters.AddWithValue("telefono", telefono)
                    cmd.Parameters.AddWithValue("correo", correo)
                    cmd.Parameters.AddWithValue("fecha_baja", Convert.ToDateTime(fecha_baja))
                    cmd.Parameters.AddWithValue("idtipojornada", idtipojornada)
                    cmd.Parameters.AddWithValue("baja", baja)

                    cmd.Parameters.AddWithValue("curp", curp)
                    cmd.Parameters.AddWithValue("cnombre", cnombre)
                    cmd.Parameters.AddWithValue("ctelefono", ctelefono)
                    cmd.Parameters.AddWithValue("bnota", bnota)
                    cmd.ExecuteNonQuery()
                    aci = "Datos del empleado actualizados."
                End If
            Else
                aci = "Error: no se actualizó, es necesario seleccionar la sucursal."
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            aci = "Error: no se actualizó, es necesario capturar el nombre del empleado."
        End If
        Return aci
    End Function
    Public Function eliminarEmpleado(ByVal idEmpleado As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idempleado FROM Vales WHERE idempleado = @idE", dbC)
        cmd.Parameters.AddWithValue("idE", idEmpleado)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim ee As String
        If rdr.HasRows Then
            rdr.Close()
            ee = "Error: este empleado no se puede eliminar, tiene vales registrados."
        Else
            rdr.Close()
            cmd.CommandText = "DELETE FROM Empleados WHERE idempleado = @idE"
            cmd.ExecuteNonQuery()
            ee = "Empleado eliminado."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return ee
    End Function
    Public Function gvEmpleados(ByVal idsucursal As Integer, ByVal activo As Boolean, ByVal baja As Boolean) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idempleado", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("empleado", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("puesto", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("activo", System.Type.GetType("System.Boolean")))
        dt.Columns.Add(New DataColumn("clave_att", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("baja", System.Type.GetType("System.Boolean")))
        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        'If activo = True And baja = False Then
        Dim cmd As New SqlCommand("SELECT idempleado, empleado, puesto, activo, clave_att,baja FROM Vista_Empleados WHERE idsucursal = @idsucursal and activo = '" & activo & "' and baja = '" & baja & "' ORDER BY empleado", dbC)
            cmd.Parameters.AddWithValue("idsucursal", idsucursal)
            cmd.Parameters.AddWithValue("activo", activo)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            While rdr.Read
                r = dt.NewRow
                r(0) = rdr("idempleado").ToString
                r(1) = rdr("empleado").ToString
                r(2) = rdr("puesto").ToString
                r(3) = rdr("activo").ToString
                r(4) = rdr("clave_att").ToString
                r(5) = rdr("baja").ToString
                dt.Rows.Add(r)
            End While
            rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()

        'ElseIf activo = False And baja = True Then
        '    Dim cmd As New SqlCommand("SELECT idempleado, empleado, puesto, activo, clave_att,baja FROM Vista_Empleados WHERE idsucursal = @idsucursal activo = '" & activo & "' and baja = '" & baja & "'  ORDER BY empleado", dbC)
        '    cmd.Parameters.AddWithValue("idsucursal", idsucursal)
        '    cmd.Parameters.AddWithValue("baja", baja)
        '    Dim rdr As SqlDataReader = cmd.ExecuteReader
        '    While rdr.Read
        '        r = dt.NewRow
        '        r(0) = rdr("idempleado").ToString
        '        r(1) = rdr("empleado").ToString
        '        r(2) = rdr("puesto").ToString
        '        r(3) = rdr("activo").ToString
        '        r(4) = rdr("clave_att").ToString
        '        r(5) = rdr("baja").ToString
        '        dt.Rows.Add(r)
        '    End While
        '    rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        'End If

        'If activo = True And baja = True Then
        '    Dim cmd As New SqlCommand("SELECT idempleado, empleado, puesto, activo, clave_att,baja FROM Vista_Empleados WHERE idsucursal = @idsucursal and baja = @baja activo = @activo ORDER BY empleado", dbC)
        '    cmd.Parameters.AddWithValue("idsucursal", idsucursal)
        '    cmd.Parameters.AddWithValue("activo", activo)
        '    cmd.Parameters.AddWithValue("baja", baja)
        '    Dim rdr As SqlDataReader = cmd.ExecuteReader
        '    While rdr.Read
        '        r = dt.NewRow
        '        r(0) = rdr("idempleado").ToString
        '        r(1) = rdr("empleado").ToString
        '        r(2) = rdr("puesto").ToString
        '        r(3) = rdr("activo").ToString
        '        r(4) = rdr("clave_att").ToString
        '        r(5) = rdr("baja").ToString
        '        dt.Rows.Add(r)
        '    End While
        '    rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        'ElseIf activo = False And baja = False Then
        '    Dim cmd As New SqlCommand("SELECT idempleado, empleado, puesto, activo, clave_att,baja FROM Vista_Empleados WHERE idsucursal = @idsucursal and baja = @baja activo = @activo ORDER BY empleado", dbC)
        '    cmd.Parameters.AddWithValue("idsucursal", idsucursal)
        '    cmd.Parameters.AddWithValue("activo", activo)
        '    cmd.Parameters.AddWithValue("baja", baja)
        '    Dim rdr As SqlDataReader = cmd.ExecuteReader
        '    While rdr.Read
        '        r = dt.NewRow
        '        r(0) = rdr("idempleado").ToString
        '        r(1) = rdr("empleado").ToString
        '        r(2) = rdr("puesto").ToString
        '        r(3) = rdr("activo").ToString
        '        r(4) = rdr("clave_att").ToString
        '        r(5) = rdr("baja").ToString
        '        dt.Rows.Add(r)
        '    End While
        '    rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        'End If

        Return dt
    End Function
    'Empleados/Sucursales
    Public Function datosEmpleSuc(ByVal sucursal As String) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idempleado, empleado FROM Vista_Empleados WHERE sucursal=@sucursal ORDER BY sucursal", dbC)
        cmd.Parameters.AddWithValue("sucursal", sucursal)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        While rdr.Read
            ReDim dsP(4)
            dsP(0) = rdr("idempleado").ToString
            dsP(1) = rdr("empleado").ToString

        End While
        ReDim dsP(0) : dsP(0) = "Error: no se encuentra este empleado."
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    'Proveedores de compras
    Public Function gvProveedorCompras() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idproveedor", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("proveedor", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("razon_social", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("cuenta", System.Type.GetType("System.String")))
        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idproveedor, proveedor, razon_social, cuenta FROM Proveedores ORDER BY proveedor", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow
            r(0) = rdr("idproveedor").ToString : r(1) = rdr("proveedor").ToString
            r(2) = rdr("razon_social").ToString : r(3) = rdr("cuenta").ToString
            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dt
    End Function
    Public Function datosProveedorCompras(ByVal idproveedor As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT * FROM Proveedores WHERE idproveedor = @idP", dbC)
        cmd.Parameters.AddWithValue("idP", idproveedor)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(10)
            dsP(0) = rdr("proveedor").ToString
            dsP(1) = rdr("cuenta").ToString
            dsP(2) = rdr("tel2").ToString
            dsP(3) = rdr("tel3").ToString
            dsP(4) = rdr("razon_social").ToString
            dsP(5) = rdr("contacto2").ToString
            dsP(6) = rdr("contacto3").ToString
            dsP(7) = rdr("diascredito").ToString
            dsP(8) = rdr("limitecredito").ToString
            dsP(9) = rdr("diaspago").ToString

        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarProveedorCompras(ByVal proveedor As String,
                                            ByVal cuenta As String,
                                            ByVal tel2 As String,
                                            ByVal tel3 As String,
                                            ByVal razon_social As String,
                                            ByVal contacto2 As String,
                                            ByVal contacto3 As String,
                                            ByVal diascredito As String,
                                            ByVal limitecredito As String,
                                            ByVal diaspago As String) As String()
        Dim ans() As String
        If proveedor <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idproveedor FROM Proveedores WHERE proveedor = @proveedor", dbC)
            cmd.Parameters.AddWithValue("proveedor", proveedor)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe."
                rdr.Close()
            Else
                rdr.Close()
                cmd.CommandText = "INSERT INTO Proveedores (proveedor,cuenta,tel2,tel3,razon_social,contacto2,contacto3,diascredito,limitecredito,diaspago) 
                    values('" & proveedor & "','" & cuenta & "','" & tel2 & "','" & tel3 & "','" & razon_social & "','" & contacto2 & "','" & contacto3 & "','" & diascredito & "','" & limitecredito & "','" & diaspago & "')"

                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT idproveedor FROM Proveedores WHERE proveedor = @proveedor"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Agregado."
                ans(1) = rdr("idproveedor").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar."
        End If
        Return ans
    End Function
    Public Function actualizarProveedorCompras(ByVal idproveedor As Integer,
                                               ByVal proveedor As String,
                                            ByVal cuenta As String,
                                            ByVal tel2 As String,
                                            ByVal tel3 As String,
                                            ByVal razon_social As String,
                                            ByVal contacto2 As String,
                                            ByVal contacto3 As String,
                                            ByVal diascredito As String,
                                            ByVal limitecredito As String,
                                            ByVal diaspago As String) As String
        Dim err As String
        If proveedor = "" Then
            err = "Error: no se actualizó, es necesario capturar."
        Else
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idproveedor FROM Proveedores WHERE proveedor = @proveedor AND idproveedor <> @idproveedor", dbC)
            cmd.Parameters.AddWithValue("proveedor", proveedor)
            cmd.Parameters.AddWithValue("idproveedor", idproveedor)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                err = "Error: no se actualizó, ya existe."
            Else
                rdr.Close()
                cmd.CommandText = "UPDATE Proveedores SET proveedor = @proveedor, cuenta = @cuenta, tel2 = @tel2, tel3 = @tel3, razon_social = @razon_social, contacto2 = @contacto2, contacto3 = @contacto3, diascredito = @diascredito, limitecredito = @limitecredito, diaspago = @diaspago  WHERE idproveedor = @idproveedor"

                cmd.Parameters.AddWithValue("cuenta", cuenta)
                cmd.Parameters.AddWithValue("tel2", tel2)
                cmd.Parameters.AddWithValue("tel3", tel3)
                cmd.Parameters.AddWithValue("razon_social", razon_social)
                cmd.Parameters.AddWithValue("contacto2", contacto2)
                cmd.Parameters.AddWithValue("contacto3", contacto3)
                cmd.Parameters.AddWithValue("diascredito", diascredito)
                cmd.Parameters.AddWithValue("limitecredito", limitecredito)
                cmd.Parameters.AddWithValue("diaspago", diaspago)

                cmd.ExecuteNonQuery()
                err = "Datos del puesto de empleados actualizados."
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        End If
        Return err
    End Function
    Public Function eliminarProveedorCompras(ByVal idproveedor As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idproveedor FROM Proveedores WHERE idproveedor = @idproveedor", dbC)
        cmd.Parameters.AddWithValue("idproveedor", idproveedor)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim err As String
        If rdr.HasRows Then

            rdr.Close()
            cmd.CommandText = "DELETE FROM Proveedores WHERE idproveedor = @idproveedor"
            cmd.ExecuteNonQuery()
            err = "Eliminado."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return err
    End Function

    'Proveedores de Gastos
    Public Function gvProveedorGastos() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idproveedor", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("proveedor", System.Type.GetType("System.String")))
        'dt.Columns.Add(New DataColumn("razon_social", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("cuenta", System.Type.GetType("System.String")))
        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idproveedor, proveedor, cuenta FROM ProveedoresG ORDER BY proveedor", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow
            r(0) = rdr("idproveedor").ToString : r(1) = rdr("proveedor").ToString
            'r(2) = rdr("razon_social").ToString
            r(2) = rdr("cuenta").ToString
            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dt
    End Function
    Public Function datosProveedorGastos(ByVal idproveedor As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT * FROM ProveedoresG WHERE idproveedor = @idP", dbC)
        cmd.Parameters.AddWithValue("idP", idproveedor)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(10)
            dsP(0) = rdr("proveedor").ToString
            dsP(1) = rdr("cuenta").ToString
            dsP(2) = rdr("tel1").ToString
            dsP(3) = rdr("tel2").ToString
            dsP(4) = rdr("tel3").ToString
            dsP(5) = rdr("contacto1").ToString
            dsP(6) = rdr("contacto2").ToString
            dsP(7) = rdr("contacto3").ToString
            dsP(8) = rdr("diascredito").ToString
            dsP(9) = rdr("limitecredito").ToString
            dsP(10) = rdr("diaspago").ToString

        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarProveedorGastos(ByVal proveedor As String,
                                            ByVal cuenta As String,
                                           ByVal tel1 As String,
                                            ByVal tel2 As String,
                                            ByVal tel3 As String,
                                            ByVal contacto1 As String,
                                            ByVal contacto2 As String,
                                            ByVal contacto3 As String,
                                            ByVal diascredito As String,
                                            ByVal limitecredito As String,
                                            ByVal diaspago As String) As String()
        Dim ans() As String
        If proveedor <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idproveedor FROM ProveedoresG WHERE proveedor = @proveedor", dbC)
            cmd.Parameters.AddWithValue("proveedor", proveedor)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe."
                rdr.Close()
            Else
                rdr.Close()
                cmd.CommandText = "INSERT INTO ProveedoresG (proveedor,cuenta,tel1,tel2,tel3,contacto1,contacto2,contacto3,diascredito,limitecredito,diaspago) 
                    values('" & proveedor & "','" & cuenta & "','" & tel1 & "','" & tel2 & "','" & tel3 & "','" & contacto1 & "','" & contacto2 & "','" & contacto3 & "','" & diascredito & "','" & limitecredito & "','" & diaspago & "')"

                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT idproveedor FROM ProveedoresG WHERE proveedor = @proveedor"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Agregado."
                ans(1) = rdr("idproveedor").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar."
        End If
        Return ans
    End Function
    Public Function actualizarProveedorGastos(ByVal idproveedor As Integer,
                                               ByVal proveedor As String,
                                            ByVal cuenta As String,
                                            ByVal tel1 As String,
                                            ByVal tel2 As String,
                                            ByVal tel3 As String,
                                            ByVal contacto1 As String,
                                            ByVal contacto2 As String,
                                            ByVal contacto3 As String,
                                            ByVal diascredito As String,
                                            ByVal limitecredito As String,
                                            ByVal diaspago As String) As String
        Dim err As String
        If proveedor = "" Then
            err = "Error: no se actualizó, es necesario capturar."
        Else
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idproveedor FROM ProveedoresG WHERE proveedor = @proveedor AND idproveedor <> @idproveedor", dbC)
            cmd.Parameters.AddWithValue("proveedor", proveedor)
            cmd.Parameters.AddWithValue("idproveedor", idproveedor)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                err = "Error: no se actualizó, ya existe."
            Else
                rdr.Close()
                cmd.CommandText = "UPDATE ProveedoresG SET proveedor = @proveedor, cuenta = @cuenta, tel1 = @tel1, tel2 = @tel2, tel3 = @tel3, contacto1 = @contacto1, contacto2 = @contacto2, contacto3 = @contacto3, diascredito = @diascredito, limitecredito = @limitecredito, diaspago = @diaspago  WHERE idproveedor = @idproveedor"

                cmd.Parameters.AddWithValue("cuenta", cuenta)
                cmd.Parameters.AddWithValue("tel1", tel1)
                cmd.Parameters.AddWithValue("tel2", tel2)
                cmd.Parameters.AddWithValue("tel3", tel3)
                cmd.Parameters.AddWithValue("contacto1", contacto1)
                cmd.Parameters.AddWithValue("contacto2", contacto2)
                cmd.Parameters.AddWithValue("contacto3", contacto3)
                cmd.Parameters.AddWithValue("diascredito", diascredito)
                cmd.Parameters.AddWithValue("limitecredito", limitecredito)
                cmd.Parameters.AddWithValue("diaspago", diaspago)

                cmd.ExecuteNonQuery()
                err = "Datos actualizados."
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        End If
        Return err
    End Function
    Public Function eliminarProveedorGastos(ByVal idproveedor As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idproveedor FROM ProveedoresG WHERE idproveedor = @idproveedor", dbC)
        cmd.Parameters.AddWithValue("idproveedor", idproveedor)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim err As String
        If rdr.HasRows Then

            rdr.Close()
            cmd.CommandText = "DELETE FROM ProveedoresG WHERE idproveedor = @idproveedor"
            cmd.ExecuteNonQuery()
            err = "Eliminado."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return err
    End Function
    'Insumos
    Public Function datosInsumos(ByVal idinsumo As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idinsumo,insumo, clave, idclase, mpc, ppm, ppc, critico, iva, medible, activo, precion  FROM Insumos WHERE idinsumo = @idinsumo", dbC)
        cmd.Parameters.AddWithValue("idinsumo", idinsumo)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(12)
            dsP(0) = rdr("idinsumo").ToString
            dsP(1) = rdr("insumo").ToString
            dsP(2) = rdr("clave").ToString
            dsP(3) = rdr("idclase").ToString

            dsP(4) = rdr("mpc").ToString
            dsP(5) = rdr("ppm").ToString
            dsP(6) = rdr("ppc").ToString

            dsP(7) = rdr("critico").ToString
            dsP(8) = rdr("iva").ToString
            dsP(9) = rdr("medible").ToString
            dsP(10) = rdr("activo").ToString

            dsP(11) = rdr("precion").ToString

        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarInsumos(ByVal insumo As String,
                                    ByVal clave As String,
                                    ByVal idclase As Integer,
                                    ByVal mpc As String,
                                    ByVal ppm As String,
                                    ByVal ppc As String,
                                    ByVal critico As Boolean,
                                    ByVal iva As Boolean,
                                    ByVal medible As Boolean,
                                    ByVal activo As Boolean,
                                    ByVal precion As String
                                    ) As String()
        Dim ans() As String
        If insumo <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT insumo FROM Insumos WHERE insumo = @insumo", dbC)
            cmd.Parameters.AddWithValue("insumo", insumo)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe."
                rdr.Close()
            Else
                rdr.Close()
                cmd.CommandText = "INSERT INTO Insumos (insumo,clave,idclase,mpc,ppm,ppc,critico,iva,medible,activo,precion,precioa,preciop) 
                    values('" & insumo & "','" & clave & "','" & idclase & "','" & mpc & "','" & ppm & "','" & ppc & "','" & critico & "','" & iva & "','" & medible & "','" & activo & "','" & precion & "','0.0','0.0')"

                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT idinsumo FROM Insumos WHERE insumo = @insumo"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Agregado."
                ans(1) = rdr("idinsumo").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar."
        End If
        Return ans
    End Function
    Public Function actualizarInsumos(ByVal idinsumo As Integer,
                                       ByVal insumo As String,
                                    ByVal clave As String,
                                    ByVal idclase As Integer,
                                    ByVal mpc As String,
                                    ByVal ppm As String,
                                    ByVal ppc As String,
                                    ByVal critico As Boolean,
                                    ByVal iva As Boolean,
                                    ByVal medible As Boolean,
                                    ByVal activo As Boolean,
                                    ByVal precion As String) As String

        Dim err As String
        If insumo = "" Then
            err = "Error: no se actualizó, es necesario capturar."
        Else
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idinsumo FROM Insumos WHERE insumo = @insumo AND idinsumo <> @idinsumo", dbC)
            cmd.Parameters.AddWithValue("insumo", insumo)
            cmd.Parameters.AddWithValue("idinsumo", idinsumo)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                err = "Error: no se actualizó, ya existe."
            Else
                rdr.Close()
                cmd.CommandText = "UPDATE Insumos SET insumo = @insumo, clave = @clave, idclase = @idclase, mpc = @mpc, ppm = @ppm, critico = @critico, iva = @iva, medible =  @medible, activo = @activo, precion = @precion WHERE idinsumo = @idinsumo"

                cmd.Parameters.AddWithValue("clave", clave)
                cmd.Parameters.AddWithValue("idclase", idclase)
                cmd.Parameters.AddWithValue("mpc", mpc)
                cmd.Parameters.AddWithValue("ppm", ppm)
                cmd.Parameters.AddWithValue("critico", critico)
                cmd.Parameters.AddWithValue("iva", iva)
                cmd.Parameters.AddWithValue("medible", medible)
                cmd.Parameters.AddWithValue("activo", activo)
                cmd.Parameters.AddWithValue("precion", precion)

                cmd.ExecuteNonQuery()
                err = "Datos actualizados."
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        End If
        Return err
    End Function
    Public Function eliminarInsumos(ByVal idinsumo As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idinsumo FROM Insumos WHERE idinsumo = @idinsumo", dbC)
        cmd.Parameters.AddWithValue("idinsumo", idinsumo)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim ee As String
        If rdr.HasRows Then

            rdr.Close()
            cmd.CommandText = "DELETE FROM Insumos WHERE idinsumo = @idinsumo"
            cmd.ExecuteNonQuery()
            ee = "Eliminado."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return ee
    End Function
    Public Function gvInsumos(ByVal idclase As Integer) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idinsumo", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("insumo", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("clase", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("clave", System.Type.GetType("System.String")))

        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        'If activo = True And baja = False Then
        Dim cmd As New SqlCommand("SELECT idinsumo,insumo, clase, clave FROM Vista_Insumos WHERE idclase = @idclase ORDER BY insumo", dbC)
        cmd.Parameters.AddWithValue("idclase", idclase)

        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow
            r(0) = rdr("idinsumo").ToString
            r(1) = rdr("insumo").ToString
            r(2) = rdr("clase").ToString
            r(3) = rdr("clave").ToString

            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()

        Return dt
    End Function

    'Productos
    Public Function datosProductos(ByVal idproducto As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idproducto, producto, clave, precio  FROM Productos WHERE idproducto = @idproducto", dbC)
        cmd.Parameters.AddWithValue("idproducto", idproducto)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(4)
            dsP(0) = rdr("idproducto").ToString
            dsP(1) = rdr("producto").ToString
            dsP(2) = rdr("clave").ToString
            dsP(3) = rdr("precio").ToString



        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarProductos(ByVal producto As String,
                                    ByVal clave As String,
                                    ByVal precio As String
                                    ) As String()
        Dim ans() As String
        If producto <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT producto FROM Productos WHERE producto = @producto", dbC)
            cmd.Parameters.AddWithValue("producto", producto)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe."
                rdr.Close()
            Else
                rdr.Close()
                cmd.CommandText = "insert into Productos(producto,clave,precio) values('" & producto & "','" & clave & "'," & precio & ")"

                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT idproducto FROM Productos WHERE producto = @producto"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Agregado."
                ans(1) = rdr("idproducto").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar."
        End If
        Return ans
    End Function
    Public Function actualizarProductos(ByVal idproducto As Integer,
                                       ByVal producto As String,
                                    ByVal clave As String,
                                    ByVal precio As String) As String

        Dim err As String
        If producto = "" Then
            err = "Error: no se actualizó, es necesario capturar."
        Else
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idproducto FROM Productos WHERE producto = @producto AND idproducto <> @idproducto", dbC)
            cmd.Parameters.AddWithValue("producto", producto)
            cmd.Parameters.AddWithValue("idproducto", idproducto)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                err = "Error: no se actualizó, ya existe."
            Else
                rdr.Close()
                cmd.CommandText = "UPDATE Productos SET producto = @producto, clave = @clave, precio = @precio WHERE idproducto = @idproducto"

                cmd.Parameters.AddWithValue("clave", clave)
                cmd.Parameters.AddWithValue("precio", precio)


                cmd.ExecuteNonQuery()
                err = "Datos actualizados."
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        End If
        Return err
    End Function
    Public Function eliminarProductos(ByVal idproducto As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idproducto FROM Productos WHERE idproducto = @idproducto", dbC)
        cmd.Parameters.AddWithValue("idproducto", idproducto)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim ee As String
        If rdr.HasRows Then

            rdr.Close()
            cmd.CommandText = "DELETE FROM Productos WHERE idproducto = @idproducto"
            cmd.ExecuteNonQuery()
            ee = "Eliminado."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return ee
    End Function
    Public Function gvProductos() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idproducto", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("producto", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("clave", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("precio", System.Type.GetType("System.String")))

        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        'If activo = True And baja = False Then
        Dim cmd As New SqlCommand("SELECT idproducto, producto, clave, precio FROM Productos  ORDER BY producto", dbC)


        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow
            r(0) = rdr("idproducto").ToString
            r(1) = rdr("producto").ToString
            r(2) = rdr("clave").ToString
            r(3) = rdr("precio").ToString

            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()

        Return dt
    End Function

    'Receta
    Public Function datosReceta(ByVal idpartida As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idinsumo, cantidad, idreceta  FROM Vista_Recetas WHERE idpartida = @idpartida", dbC)
        cmd.Parameters.AddWithValue("idpartida", idpartida)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(3)
            dsP(0) = rdr("idinsumo").ToString
            dsP(1) = rdr("cantidad").ToString
            dsP(2) = rdr("idreceta").ToString

        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function datosRecetaID(ByVal idproducto As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT TOP 1 idreceta  FROM Vista_Recetas WHERE idproducto = @idproducto", dbC)
        cmd.Parameters.AddWithValue("idproducto", idproducto)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(1)

            dsP(0) = rdr("idreceta").ToString

        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarReceta(ByVal idinsumo As Integer,
                                    ByVal cantidad As String, ByVal idreceta As Integer
                                    ) As String()
        Dim ans() As String
        If cantidad <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idinsumo FROM PartidasReceta WHERE idinsumo = @idinsumo AND idreceta = @idreceta", dbC)
            cmd.Parameters.AddWithValue("idinsumo", idinsumo)
            cmd.Parameters.AddWithValue("idreceta", idreceta)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe."
                rdr.Close()
            Else
                rdr.Close()
                cmd.CommandText = "insert into PartidasReceta(idinsumo,cantidad,idreceta) values('" & idinsumo & "','" & cantidad & "','" & idreceta & "')"

                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT idinsumo FROM PartidasReceta WHERE idinsumo = @idinsumo"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Agregado."
                ans(1) = rdr("idinsumo").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar."
        End If
        Return ans
    End Function
    Public Function actualizarReceta(ByVal idpartida As Integer,
                                      ByVal idinsumo As Integer,
                                    ByVal cantidad As String, ByVal idreceta As Integer) As String

        Dim err As String
        If cantidad = "" Then
            err = "Error: no se actualizó, es necesario capturar."
        Else
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idinsumo FROM PartidasReceta WHERE idreceta = @idreceta AND idinsumo = @idinsumo AND cantidad = @cantidad", dbC)
            cmd.Parameters.AddWithValue("idreceta", idreceta)
            cmd.Parameters.AddWithValue("idinsumo", idinsumo)
            cmd.Parameters.AddWithValue("cantidad", cantidad)
            'cmd.Parameters.AddWithValue("idreceta", idreceta)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                err = "Error: no se actualizó, ya existe."
            Else
                rdr.Close()
                cmd.CommandText = "UPDATE PartidasReceta SET cantidad = @cantidad, idinsumo = @idinsumo WHERE idpartida = @idpartida"
                cmd.Parameters.AddWithValue("idpartida", idpartida)

                cmd.ExecuteNonQuery()
                err = "Datos actualizados."
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        End If
        Return err
    End Function
    Public Function eliminarReceta(ByVal idpartida As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idpartida FROM PartidasReceta WHERE idpartida = @idpartida", dbC)
        cmd.Parameters.AddWithValue("idpartida", idpartida)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim ee As String
        If rdr.HasRows Then

            rdr.Close()
            cmd.CommandText = "DELETE FROM PartidasReceta WHERE idpartida = @idpartida"
            cmd.ExecuteNonQuery()
            ee = "Eliminado."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return ee
    End Function
    Public Function gvReceta(ByVal idproducto As Integer) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idpartida", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("producto", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("insumo", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("cantidad", System.Type.GetType("System.String")))

        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        'If activo = True And baja = False Then
        Dim cmd As New SqlCommand("SELECT  idpartida, producto, insumo, cantidad FROM Vista_Recetas WHERE idproducto = @idproducto ORDER BY producto", dbC)
        cmd.Parameters.AddWithValue("idproducto", idproducto)

        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow
            r(0) = rdr("idpartida").ToString
            r(1) = rdr("producto").ToString
            r(2) = rdr("insumo").ToString
            r(3) = rdr("cantidad").ToString

            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()

        Return dt
    End Function

    'ClasesProductos
    Public Function datosClasesProductos(ByVal idclase As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idclase, clase, clave, costo, cuenta_s1, cuenta_s2, cuenta_s3, cuenta_s4, cuenta_s5, cuenta_s6, cuenta_l, cuenta_i  FROM Clases WHERE idclase = @idclase", dbC)
        cmd.Parameters.AddWithValue("idclase", idclase)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(12)
            dsP(0) = rdr("idclase").ToString
            dsP(1) = rdr("clase").ToString
            dsP(2) = rdr("clave").ToString
            dsP(3) = rdr("costo").ToString

            dsP(4) = rdr("cuenta_s1").ToString
            dsP(5) = rdr("cuenta_s2").ToString
            dsP(6) = rdr("cuenta_s3").ToString
            dsP(7) = rdr("cuenta_s4").ToString

            dsP(8) = rdr("cuenta_s5").ToString
            dsP(9) = rdr("cuenta_s6").ToString
            dsP(10) = rdr("cuenta_l").ToString
            dsP(11) = rdr("cuenta_i").ToString

        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarClasesProductos(ByVal clase As String,
                                           ByVal clave As String,
                                           ByVal costo As String,
                                           ByVal cuenta_s1 As String,
                                           ByVal cuenta_s2 As String,
                                           ByVal cuenta_s3 As String,
                                           ByVal cuenta_s4 As String,
                                           ByVal cuenta_s5 As String,
                                           ByVal cuenta_s6 As String,
                                           ByVal cuenta_l As String,
                                           ByVal cuenta_i As String
                                    ) As String()
        Dim ans() As String
        If clase <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT clase FROM Clases WHERE clase = @clase", dbC)
            cmd.Parameters.AddWithValue("clase", clase)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe."
                rdr.Close()
            Else
                rdr.Close()
                cmd.CommandText = "insert into Clases(clase, clave, costo, cuenta_s1, cuenta_s2, cuenta_s3, cuenta_s4, cuenta_s5, cuenta_s6, cuenta_l, cuenta_i) values('" & clase & "','" & clave & "','" & costo & "','" & cuenta_s1 & "','" & cuenta_s2 & "','" & cuenta_s3 & "','" & cuenta_s4 & "','" & cuenta_s5 & "','" & cuenta_s6 & "','" & cuenta_l & "','" & cuenta_i & "')"

                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT idclase FROM Clases WHERE clase = @clase"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Agregado."
                ans(1) = rdr("idclase").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar."
        End If
        Return ans
    End Function
    Public Function actualizarClasesProductos(ByVal idclase As Integer,
                                           ByVal clase As String,
                                           ByVal clave As String,
                                           ByVal costo As String,
                                           ByVal cuenta_s1 As String,
                                           ByVal cuenta_s2 As String,
                                           ByVal cuenta_s3 As String,
                                           ByVal cuenta_s4 As String,
                                           ByVal cuenta_s5 As String,
                                           ByVal cuenta_s6 As String,
                                           ByVal cuenta_l As String,
                                           ByVal cuenta_i As String) As String

        Dim err As String
        If clase = "" Then
            err = "Error: no se actualizó, es necesario capturar."
        Else
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idclase FROM Clases WHERE clase = @clase AND idclase <> @idclase", dbC)
            cmd.Parameters.AddWithValue("clase", clase)
            cmd.Parameters.AddWithValue("idclase", idclase)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                err = "Error: no se actualizó, ya existe."
            Else
                rdr.Close()
                cmd.CommandText = "UPDATE Clases SET clase = @clase, clave = @clave, costo = @costo, cuenta_s1 = @cuenta_s1, cuenta_s2 = @cuenta_s2, cuenta_s3 = @cuenta_s3, cuenta_s4 = @cuenta_s4 , cuenta_s5 = @cuenta_s5, cuenta_s6 = @cuenta_s6, cuenta_l = @cuenta_l, cuenta_i = @cuenta_i WHERE idclase = @idclase"

                cmd.Parameters.AddWithValue("clave", clave)
                cmd.Parameters.AddWithValue("costo", costo)
                cmd.Parameters.AddWithValue("cuenta_s1", cuenta_s1)
                cmd.Parameters.AddWithValue("cuenta_s2", cuenta_s2)
                cmd.Parameters.AddWithValue("cuenta_s3", cuenta_s3)
                cmd.Parameters.AddWithValue("cuenta_s4", cuenta_s4)
                cmd.Parameters.AddWithValue("cuenta_s5", cuenta_s5)
                cmd.Parameters.AddWithValue("cuenta_s6", cuenta_s6)
                cmd.Parameters.AddWithValue("cuenta_l", cuenta_l)
                cmd.Parameters.AddWithValue("cuenta_i", cuenta_i)

                cmd.ExecuteNonQuery()
                err = "Datos actualizados."
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        End If
        Return err
    End Function
    Public Function eliminarClasesProductos(ByVal idclase As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idclase FROM Clases WHERE idclase = @idclase", dbC)
        cmd.Parameters.AddWithValue("idclase", idclase)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim ee As String
        If rdr.HasRows Then

            rdr.Close()
            cmd.CommandText = "DELETE FROM Clases WHERE idclase = @idclase"
            cmd.ExecuteNonQuery()
            ee = "Eliminado."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return ee
    End Function
    Public Function gvClasesProductos() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idclase", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("clase", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("clave", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("costo", System.Type.GetType("System.String")))

        dt.Columns.Add(New DataColumn("cuenta_s1", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("cuenta_s2", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("cuenta_s3", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("cuenta_s4", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("cuenta_s5", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("cuenta_s6", System.Type.GetType("System.String")))

        dt.Columns.Add(New DataColumn("cuenta_l", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("cuenta_i", System.Type.GetType("System.String")))

        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()

        Dim cmd As New SqlCommand("SELECT idclase, clase, clave, costo, cuenta_s1, cuenta_s2, cuenta_s3, cuenta_s4, cuenta_s5, cuenta_s6, cuenta_l, cuenta_i  FROM Clases  ORDER BY clase", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow

            r(0) = rdr("idclase").ToString
            r(1) = rdr("clase").ToString
            r(2) = rdr("clave").ToString
            r(3) = rdr("costo").ToString

            r(4) = rdr("cuenta_s1").ToString
            r(5) = rdr("cuenta_s2").ToString
            r(6) = rdr("cuenta_s3").ToString
            r(7) = rdr("cuenta_s4").ToString

            r(8) = rdr("cuenta_s5").ToString
            r(9) = rdr("cuenta_s6").ToString
            r(10) = rdr("cuenta_l").ToString
            r(11) = rdr("cuenta_i").ToString

            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()

        Return dt
    End Function

    'ClasesGastos
    Public Function datosClasesGastos(ByVal idclase As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idclase, clase, clave, grupo  FROM ClasesG WHERE idclase = @idclase", dbC)
        cmd.Parameters.AddWithValue("idclase", idclase)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(4)
            dsP(0) = rdr("idclase").ToString
            dsP(1) = rdr("clase").ToString
            dsP(2) = rdr("clave").ToString
            dsP(3) = rdr("grupo").ToString



        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarClasesGastos(ByVal clase As String,
                                    ByVal clave As String,
                                    ByVal grupo As String
                                    ) As String()
        Dim ans() As String
        If clase <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT clase FROM ClasesG WHERE clase = @clase", dbC)
            cmd.Parameters.AddWithValue("clase", clase)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe."
                rdr.Close()
            Else
                rdr.Close()
                cmd.CommandText = "insert into ClasesG (clase,clave,grupo) values('" & clase & "','" & clave & "'," & grupo & ")"

                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT idclase FROM ClasesG WHERE clase = @clase"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Agregado."
                ans(1) = rdr("idclase").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar."
        End If
        Return ans
    End Function
    Public Function actualizarClasesGastos(ByVal idclase As Integer,
                                       ByVal clase As String,
                                    ByVal clave As String,
                                    ByVal grupo As String) As String

        Dim err As String
        If clase = "" Then
            err = "Error: no se actualizó, es necesario capturar."
        Else
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idclase FROM ClasesG WHERE clase = @clase AND idclase <> @idclase", dbC)
            cmd.Parameters.AddWithValue("clase", clase)
            cmd.Parameters.AddWithValue("idclase", idclase)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                err = "Error: no se actualizó, ya existe."
            Else
                rdr.Close()
                cmd.CommandText = "UPDATE ClasesG SET clase = @clase, clave = @clave, grupo = @grupo WHERE idclase = @idclase"

                cmd.Parameters.AddWithValue("clave", clave)
                cmd.Parameters.AddWithValue("grupo", grupo)


                cmd.ExecuteNonQuery()
                err = "Datos actualizados."
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        End If
        Return err
    End Function
    Public Function eliminarClasesGastos(ByVal idclase As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idclase FROM ClasesG WHERE idclase = @idclase", dbC)
        cmd.Parameters.AddWithValue("idclase", idclase)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim ee As String
        If rdr.HasRows Then

            rdr.Close()
            cmd.CommandText = "DELETE FROM ClasesG WHERE idclase = @idclase"
            cmd.ExecuteNonQuery()
            ee = "Eliminado."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return ee
    End Function
    Public Function gvClasesGastos() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idclase", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("clase", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("clave", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("grupo", System.Type.GetType("System.String")))

        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        'If activo = True And baja = False Then
        Dim cmd As New SqlCommand("SELECT idclase, clase, clave, grupo FROM ClasesG  ORDER BY clase", dbC)


        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow
            r(0) = rdr("idclase").ToString
            r(1) = rdr("clase").ToString
            r(2) = rdr("clave").ToString
            r(3) = rdr("grupo").ToString

            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()

        Return dt
    End Function

    'ConceptosGastos


    'ClasesCajaChica
    Public Function datosClasesCajaChica(ByVal idclase As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idclase, clase, clave FROM ClasesCH WHERE idclase = @idclase", dbC)
        cmd.Parameters.AddWithValue("idclase", idclase)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(3)
            dsP(0) = rdr("idclase").ToString
            dsP(1) = rdr("clase").ToString
            dsP(2) = rdr("clave").ToString

        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarClasesCajaChica(ByVal clase As String,
                                           ByVal clave As String) As String()
        Dim ans() As String
        If clase <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT clase FROM ClasesCH WHERE clase = @clase", dbC)
            cmd.Parameters.AddWithValue("clase", clase)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe."
                rdr.Close()
            Else
                rdr.Close()
                cmd.CommandText = "insert into ClasesCH(clase, clave) values('" & clase & "','" & clave & "')"

                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT idclase FROM ClasesCH WHERE clase = @clase"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Agregado."
                ans(1) = rdr("idclase").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar."
        End If
        Return ans
    End Function
    Public Function actualizarClasesCajaChica(ByVal idclase As Integer,
                                           ByVal clase As String,
                                           ByVal clave As String) As String

        Dim err As String
        If clase = "" Then
            err = "Error: no se actualizó, es necesario capturar."
        Else
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idclase FROM ClasesCH WHERE clase = @clase AND idclase <> @idclase", dbC)
            cmd.Parameters.AddWithValue("clase", clase)
            cmd.Parameters.AddWithValue("idclase", idclase)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                err = "Error: no se actualizó, ya existe."
            Else
                rdr.Close()
                cmd.CommandText = "UPDATE ClasesCH SET clase = @clase, clave = @clave WHERE idclase = @idclase"

                cmd.Parameters.AddWithValue("clave", clave)

                cmd.ExecuteNonQuery()
                err = "Datos actualizados."
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        End If
        Return err
    End Function
    Public Function eliminarClasesCajaChica(ByVal idclase As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idclase FROM ClasesCH WHERE idclase = @idclase", dbC)
        cmd.Parameters.AddWithValue("idclase", idclase)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim ee As String
        If rdr.HasRows Then

            rdr.Close()
            cmd.CommandText = "DELETE FROM ClasesCH WHERE idclase = @idclase"
            cmd.ExecuteNonQuery()
            ee = "Eliminado."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return ee
    End Function
    Public Function gvClasesCajaChica() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idclase", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("clase", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("clave", System.Type.GetType("System.String")))

        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()

        Dim cmd As New SqlCommand("SELECT idclase, clase, clave FROM ClasesCH  ORDER BY clase", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow

            r(0) = rdr("idclase").ToString
            r(1) = rdr("clase").ToString
            r(2) = rdr("clave").ToString


            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()

        Return dt
    End Function

    'ConceptosdeVales
    Public Function datosConceptosdeVales(ByVal idvale As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idvale, concepto FROM ConceptosVales WHERE idvale = @idvale", dbC)
        cmd.Parameters.AddWithValue("idvale", idvale)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(2)
            dsP(0) = rdr("idvale").ToString
            dsP(1) = rdr("concepto").ToString


        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarConceptosdeVales(ByVal concepto As String) As String()
        Dim ans() As String
        If concepto <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT concepto FROM ConceptosVales WHERE concepto = @concepto", dbC)
            cmd.Parameters.AddWithValue("concepto", concepto)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe."
                rdr.Close()
            Else
                rdr.Close()
                cmd.CommandText = "insert into ConceptosVales(concepto) values('" & concepto & "')"

                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT idvale FROM ConceptosVales WHERE concepto = @concepto"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Agregado."
                ans(1) = rdr("idvale").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar."
        End If
        Return ans
    End Function
    Public Function actualizarConceptosdeVales(ByVal idvale As Integer,
                                           ByVal concepto As String) As String

        Dim err As String
        If concepto = "" Then
            err = "Error: no se actualizó, es necesario capturar."
        Else
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idvale FROM ConceptosVales WHERE concepto = @concepto AND idvale <> @idvale", dbC)
            cmd.Parameters.AddWithValue("concepto", concepto)
            cmd.Parameters.AddWithValue("idvale", idvale)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                err = "Error: no se actualizó, ya existe."
            Else
                rdr.Close()
                cmd.CommandText = "UPDATE ConceptosVales SET concepto = @concepto WHERE idvale = @idvale"


                cmd.ExecuteNonQuery()
                err = "Datos actualizados."
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        End If
        Return err
    End Function
    Public Function eliminarConceptosdeVales(ByVal idvale As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idvale FROM ConceptosVales WHERE idvale = @idvale", dbC)
        cmd.Parameters.AddWithValue("idvale", idvale)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim ee As String
        If rdr.HasRows Then

            rdr.Close()
            cmd.CommandText = "DELETE FROM ConceptosVales WHERE idvale = @idvale"
            cmd.ExecuteNonQuery()
            ee = "Eliminado."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return ee
    End Function
    Public Function gvConceptosdeVales() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idvale", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("concepto", System.Type.GetType("System.String")))


        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()

        Dim cmd As New SqlCommand("SELECT idvale, concepto FROM ConceptosVales  ORDER BY concepto", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow

            r(0) = rdr("idvale").ToString
            r(1) = rdr("concepto").ToString


            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()

        Return dt
    End Function

    'ConfiSucursales
    Public Function datosConfiSucursales(ByVal idsucursal As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idsucursal, sucursal FROM Sucursales WHERE idsucursal = @idsucursal", dbC)
        cmd.Parameters.AddWithValue("idsucursal", idsucursal)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(2)
            dsP(0) = rdr("idsucursal").ToString
            dsP(1) = rdr("sucursal").ToString


        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarConfiSucursales(ByVal sucursal As String) As String()
        Dim ans() As String
        If sucursal <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT sucursal FROM Sucursales WHERE sucursal = @sucursal", dbC)
            cmd.Parameters.AddWithValue("sucursal", sucursal)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe."
                rdr.Close()
            Else
                rdr.Close()
                cmd.CommandText = "insert into Sucursales (sucursal) values('" & sucursal & "')"

                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT idsucursal FROM Sucursales WHERE sucursal = @sucursal"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Agregado."
                ans(1) = rdr("idsucursal").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar."
        End If
        Return ans
    End Function
    Public Function actualizarConfiSucursales(ByVal idsucursal As Integer,
                                           ByVal sucursal As String) As String

        Dim err As String
        If sucursal = "" Then
            err = "Error: no se actualizó, es necesario capturar."
        Else
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idsucursal FROM Sucursales WHERE sucursal = @sucursal AND idsucursal <> @idsucursal", dbC)
            cmd.Parameters.AddWithValue("sucursal", sucursal)
            cmd.Parameters.AddWithValue("idsucursal", idsucursal)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                err = "Error: no se actualizó, ya existe."
            Else
                rdr.Close()
                cmd.CommandText = "UPDATE Sucursales SET sucursal = @sucursal WHERE idsucursal = @idsucursal"


                cmd.ExecuteNonQuery()
                err = "Datos actualizados."
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        End If
        Return err
    End Function
    Public Function eliminarConfiSucursales(ByVal idsucursal As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idsucursal FROM Sucursales WHERE idsucursal = @idsucursal", dbC)
        cmd.Parameters.AddWithValue("idsucursal", idsucursal)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim ee As String
        If rdr.HasRows Then

            rdr.Close()
            cmd.CommandText = "DELETE FROM Sucursales WHERE idsucursal = @idsucursal"
            cmd.ExecuteNonQuery()
            ee = "Eliminado."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return ee
    End Function
    Public Function gvConfiSucursales() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idsucursal", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("sucursal", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("prorrateo", System.Type.GetType("System.String")))

        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()

        Dim cmd As New SqlCommand("SELECT idsucursal, sucursal, prorrateo FROM Sucursales  ORDER BY sucursal", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow

            r(0) = rdr("idsucursal").ToString
            r(1) = rdr("sucursal").ToString
            r(2) = rdr("prorrateo").ToString

            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()

        Return dt
    End Function

    'ConfiIncidencias
    Public Function datosConfiIncidencias(ByVal idincidencia As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idincidencia, incidencia FROM Incidencias WHERE idincidencia = @idincidencia", dbC)
        cmd.Parameters.AddWithValue("idincidencia", idincidencia)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(2)
            dsP(0) = rdr("idincidencia").ToString
            dsP(1) = rdr("incidencia").ToString


        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarConfiIncidencias(ByVal incidencia As String) As String()
        Dim ans() As String
        If incidencia <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT incidencia FROM Incidencias WHERE incidencia = @incidencia", dbC)
            cmd.Parameters.AddWithValue("incidencia", incidencia)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe."
                rdr.Close()
            Else
                rdr.Close()
                cmd.CommandText = "insert into Incidencias (incidencia) values('" & incidencia & "')"

                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT idincidencia FROM Incidencias WHERE incidencia = @incidencia"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Agregado."
                ans(1) = rdr("idincidencia").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar."
        End If
        Return ans
    End Function
    Public Function actualizarConfiIncidencias(ByVal idincidencia As Integer,
                                           ByVal incidencia As String) As String

        Dim err As String
        If incidencia = "" Then
            err = "Error: no se actualizó, es necesario capturar."
        Else
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idincidencia FROM Incidencias WHERE incidencia = @incidencia AND idincidencia <> @idincidencia", dbC)
            cmd.Parameters.AddWithValue("incidencia", incidencia)
            cmd.Parameters.AddWithValue("idincidencia", idincidencia)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                err = "Error: no se actualizó, ya existe."
            Else
                rdr.Close()
                cmd.CommandText = "UPDATE Incidencias SET incidencia = @incidencia WHERE idincidencia = @idincidencia"


                cmd.ExecuteNonQuery()
                err = "Datos actualizados."
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        End If
        Return err
    End Function
    Public Function eliminarConfiIncidencias(ByVal idincidencia As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idincidencia FROM Incidencias WHERE idincidencia = @idincidencia", dbC)
        cmd.Parameters.AddWithValue("idincidencia", idincidencia)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim ee As String
        If rdr.HasRows Then

            rdr.Close()
            cmd.CommandText = "DELETE FROM Incidencias WHERE idincidencia = @idincidencia"
            cmd.ExecuteNonQuery()
            ee = "Eliminado."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return ee
    End Function
    Public Function gvConfiIncidencias() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idincidencia", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("incidencia", System.Type.GetType("System.String")))


        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()

        Dim cmd As New SqlCommand("SELECT idincidencia, incidencia FROM Incidencias  ORDER BY incidencia", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow

            r(0) = rdr("idincidencia").ToString
            r(1) = rdr("incidencia").ToString


            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()

        Return dt
    End Function

    'ConfiDiscrepancia
    Public Function datosConfiDiscrepancia(ByVal iddiscrepancia As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT iddiscrepancia, discrepancia FROM Discrepancias WHERE iddiscrepancia = @iddiscrepancia", dbC)
        cmd.Parameters.AddWithValue("iddiscrepancia", iddiscrepancia)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(2)
            dsP(0) = rdr("iddiscrepancia").ToString
            dsP(1) = rdr("discrepancia").ToString


        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarConfiDiscrepancia(ByVal discrepancia As String) As String()
        Dim ans() As String
        If discrepancia <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT discrepancia FROM Discrepancias WHERE discrepancia = @discrepancia", dbC)
            cmd.Parameters.AddWithValue("discrepancia", discrepancia)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe."
                rdr.Close()
            Else
                rdr.Close()
                cmd.CommandText = "insert into Discrepancias (discrepancia) values('" & discrepancia & "')"

                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT iddiscrepancia FROM Discrepancias WHERE discrepancia = @discrepancia"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Agregado."
                ans(1) = rdr("iddiscrepancia").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar."
        End If
        Return ans
    End Function
    Public Function actualizarConfiDiscrepancia(ByVal iddiscrepancia As Integer,
                                           ByVal discrepancia As String) As String

        Dim err As String
        If discrepancia = "" Then
            err = "Error: no se actualizó, es necesario capturar."
        Else
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT iddiscrepancia FROM Discrepancias WHERE discrepancia = @discrepancia AND iddiscrepancia <> @iddiscrepancia", dbC)
            cmd.Parameters.AddWithValue("discrepancia", discrepancia)
            cmd.Parameters.AddWithValue("iddiscrepancia", iddiscrepancia)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                err = "Error: no se actualizó, ya existe."
            Else
                rdr.Close()
                cmd.CommandText = "UPDATE Discrepancias SET discrepancia = @discrepancia WHERE iddiscrepancia = @iddiscrepancia"


                cmd.ExecuteNonQuery()
                err = "Datos actualizados."
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        End If
        Return err
    End Function
    Public Function eliminarConfiDiscrepancia(ByVal iddiscrepancia As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT iddiscrepancia FROM Discrepancias WHERE iddiscrepancia = @iddiscrepancia", dbC)
        cmd.Parameters.AddWithValue("iddiscrepancia", iddiscrepancia)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim ee As String
        If rdr.HasRows Then

            rdr.Close()
            cmd.CommandText = "DELETE FROM Discrepancias WHERE iddiscrepancia = @iddiscrepancia"
            cmd.ExecuteNonQuery()
            ee = "Eliminado."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return ee
    End Function
    Public Function gvConfiDiscrepancias() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("iddiscrepancia", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("discrepancia", System.Type.GetType("System.String")))


        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()

        Dim cmd As New SqlCommand("SELECT iddiscrepancia, discrepancia FROM Discrepancias  ORDER BY discrepancia", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow

            r(0) = rdr("iddiscrepancia").ToString
            r(1) = rdr("discrepancia").ToString


            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()

        Return dt
    End Function

    'EquipClase
    Public Function datosEquipClase(ByVal idclase As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idclase, clase FROM ClasesA WHERE idclase = @idclase", dbC)
        cmd.Parameters.AddWithValue("idclase", idclase)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(2)
            dsP(0) = rdr("idclase").ToString
            dsP(1) = rdr("clase").ToString


        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarEquipClase(ByVal clase As String) As String()
        Dim ans() As String
        If clase <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT clase FROM ClasesA WHERE clase = @clase", dbC)
            cmd.Parameters.AddWithValue("clase", clase)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe."
                rdr.Close()
            Else
                rdr.Close()
                cmd.CommandText = "insert into ClasesA (clase) values('" & clase & "')"

                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT idclase FROM ClasesA WHERE clase = @clase"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Agregado."
                ans(1) = rdr("idclase").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar."
        End If
        Return ans
    End Function
    Public Function actualizarEquipClase(ByVal idclase As Integer,
                                           ByVal clase As String) As String

        Dim err As String
        If clase = "" Then
            err = "Error: no se actualizó, es necesario capturar."
        Else
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idclase FROM ClasesA WHERE clase = @clase AND idclase <> @idclase", dbC)
            cmd.Parameters.AddWithValue("clase", clase)
            cmd.Parameters.AddWithValue("idclase", idclase)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                err = "Error: no se actualizó, ya existe."
            Else
                rdr.Close()
                cmd.CommandText = "UPDATE ClasesA SET clase = @clase WHERE idclase = @idclase"


                cmd.ExecuteNonQuery()
                err = "Datos actualizados."
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        End If
        Return err
    End Function
    Public Function eliminarEquipClase(ByVal idclase As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idclase FROM ClasesA WHERE idclase = @idclase", dbC)
        cmd.Parameters.AddWithValue("idclase", idclase)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim ee As String
        If rdr.HasRows Then

            rdr.Close()
            cmd.CommandText = "DELETE FROM ClasesA WHERE idclase = @idclase"
            cmd.ExecuteNonQuery()
            ee = "Eliminado."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return ee
    End Function
    Public Function gvEquipClase() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idclase", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("clase", System.Type.GetType("System.String")))


        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()

        Dim cmd As New SqlCommand("SELECT idclase, clase FROM ClasesA  ORDER BY clase", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow

            r(0) = rdr("idclase").ToString
            r(1) = rdr("clase").ToString


            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()

        Return dt
    End Function

    'EquipActivo
    Public Function datosEquipActivo(ByVal idactivo As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idactivo, activo, idclase FROM Activos WHERE idactivo = @idactivo", dbC)
        cmd.Parameters.AddWithValue("idactivo", idactivo)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(2)
            dsP(0) = rdr("idactivo").ToString
            dsP(1) = rdr("activo").ToString
            dsP(2) = rdr("idclase").ToString


        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarEquipActivo(ByVal activo As String, ByVal idclase As Integer) As String()
        Dim ans() As String
        If activo <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT activo FROM Activos WHERE activo = @activo", dbC)
            cmd.Parameters.AddWithValue("activo", activo)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe."
                rdr.Close()
            Else
                rdr.Close()
                cmd.CommandText = "insert into Activos (activo,idclase) values('" & activo & "','" & idclase & "')"

                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT idactivo FROM Activos WHERE activo = @activo"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Agregado."
                ans(1) = rdr("idactivo").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar."
        End If
        Return ans
    End Function
    Public Function actualizarEquipActivo(ByVal idactivo As Integer, ByVal activo As String,
                                           ByVal idclase As Integer) As String

        Dim err As String
        If activo = "" Then
            err = "Error: no se actualizó, es necesario capturar."
        Else
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idactivo FROM Activos WHERE activo = @activo AND idactivo <> @idactivo", dbC)
            cmd.Parameters.AddWithValue("activo", activo)
            cmd.Parameters.AddWithValue("idclase", idclase)
            cmd.Parameters.AddWithValue("idactivo", idactivo)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                err = "Error: no se actualizó, ya existe."
            Else
                rdr.Close()
                cmd.CommandText = "UPDATE Activos SET activo = @activo, idclase = @idclase  WHERE idactivo = @idactivo"


                cmd.ExecuteNonQuery()
                err = "Datos actualizados."
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        End If
        Return err
    End Function
    Public Function eliminarEquipActivo(ByVal idactivo As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idactivo FROM Activos WHERE idactivo = @idactivo", dbC)
        cmd.Parameters.AddWithValue("idactivo", idactivo)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim ee As String
        If rdr.HasRows Then

            rdr.Close()
            cmd.CommandText = "DELETE FROM Activos WHERE idactivo = @idactivo"
            cmd.ExecuteNonQuery()
            ee = "Eliminado."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return ee
    End Function
    Public Function gvEquipActivo() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idactivo", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("activo", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("clase", System.Type.GetType("System.String")))

        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()

        Dim cmd As New SqlCommand("SELECT idactivo, activo, clase FROM vm_EquipActivo  ORDER BY activo", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow

            r(0) = rdr("idactivo").ToString
            r(1) = rdr("activo").ToString
            r(2) = rdr("clase").ToString

            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()

        Return dt
    End Function

    'EquipRefac
    Public Function datosEquipRefac(ByVal idrefacc As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idrefacc, refacc, idactivo FROM Refacciones WHERE idrefacc = @idrefacc", dbC)
        cmd.Parameters.AddWithValue("idrefacc", idrefacc)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(2)
            dsP(0) = rdr("idrefacc").ToString
            dsP(1) = rdr("refacc").ToString
            dsP(2) = rdr("idactivo").ToString


        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarEquipRefac(ByVal refacc As String, ByVal idactivo As Integer) As String()
        Dim ans() As String
        If refacc <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT refacc FROM Refacciones WHERE refacc = @refacc", dbC)
            cmd.Parameters.AddWithValue("refacc", refacc)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe."
                rdr.Close()
            Else
                rdr.Close()
                cmd.CommandText = "insert into Refacciones (refacc,idactivo) values('" & refacc & "','" & idactivo & "')"

                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT idrefacc FROM Refacciones WHERE refacc = @refacc"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Agregado."
                ans(1) = rdr("idrefacc").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar."
        End If
        Return ans
    End Function
    Public Function actualizarEquipRefac(ByVal idrefacc As Integer, ByVal refacc As String,
                                           ByVal idactivo As Integer) As String

        Dim err As String
        If refacc = "" Then
            err = "Error: no se actualizó, es necesario capturar."
        Else
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idrefacc FROM Refacciones WHERE refacc = @refacc AND idrefacc <> @idrefacc", dbC)
            cmd.Parameters.AddWithValue("refacc", refacc)
            cmd.Parameters.AddWithValue("idrefacc", idrefacc)
            cmd.Parameters.AddWithValue("idactivo", idactivo)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                err = "Error: no se actualizó, ya existe."
            Else
                rdr.Close()
                cmd.CommandText = "UPDATE Refacciones SET refacc = @refacc, idactivo = @idactivo  WHERE idrefacc = @idrefacc"


                cmd.ExecuteNonQuery()
                err = "Datos actualizados."
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        End If
        Return err
    End Function
    Public Function eliminarEquipRefac(ByVal idrefacc As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idrefacc FROM Refacciones WHERE idrefacc = @idrefacc", dbC)
        cmd.Parameters.AddWithValue("idrefacc", idrefacc)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim ee As String
        If rdr.HasRows Then

            rdr.Close()
            cmd.CommandText = "DELETE FROM Refacciones WHERE idrefacc = @idrefacc"
            cmd.ExecuteNonQuery()
            ee = "Eliminado."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return ee
    End Function
    Public Function gvEquipRefac() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idrefacc", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("refacc", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("activo", System.Type.GetType("System.String")))

        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()

        Dim cmd As New SqlCommand("SELECT idrefacc, refacc, activo FROM vm_EquipRefac  ORDER BY refacc", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow

            r(0) = rdr("idrefacc").ToString
            r(1) = rdr("refacc").ToString
            r(2) = rdr("activo").ToString

            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()

        Return dt
    End Function

End Class

