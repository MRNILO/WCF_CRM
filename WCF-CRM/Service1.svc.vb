Imports System.Data.SqlClient
Imports System.Net
Imports System.Net.Mail
Imports System.ServiceModel.Activation
Imports System.Web.Script.Serialization
Imports Newtonsoft.Json
Imports SendGrid

Public Class Service1
    Implements IService1

    Dim ConexionStr As String = ConfigurationManager.ConnectionStrings("Conexion_SQL").ConnectionString()
    Dim Conexion As New SqlConnection(ConexionStr)

    Private ODBC_OBJ As New DirectConn
    Private GE_SQL As New SQL_Functions

#Region "Error login"
    Function Obtener_nombre_metodo() As String
        Dim st = New StackTrace()
        Dim sf = st.GetFrame(1)

        Return sf.GetMethod().Name
    End Function
    Function Inserta_error(ByVal Mensaje As String, ByVal Operacion As String) As Boolean Implements IService1.Inserta_error
        Dim cmd As New SqlCommand("Inserta_Error", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@PMensaje", Mensaje)
        cmd.Parameters.AddWithValue("@POperacion", Operacion)

        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception

            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
#End Region

    Public Sub New()
    End Sub

#Region "Adm"
    Function Obtener_AcumuladosAdm(ByVal FechaInicio As Date, ByVal FechaFinal As Date) As List(Of CAcumuladosSupervisor) Implements IService1.Obtener_AcumuladosAdm
        Dim Resultado As New List(Of CAcumuladosSupervisor)
        Dim cmd As New SqlCommand("Obtener_AcumuladosAdm", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        'cmd.Parameters.AddWithValue("@idSupervisor", id_supervisor)
        cmd.Parameters.AddWithValue("@FechaInicio", FechaInicio)
        cmd.Parameters.AddWithValue("@FechaFinal", FechaFinal)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CAcumuladosSupervisor
        While reader.Read
            Aux = New CAcumuladosSupervisor
            Aux.Cantidad = DirectCast(reader.Item("Cantidad"), Integer)
            Aux.NombreCliente = DirectCast(reader.Item("NombreCliente"), String)
            Aux.Producto = DirectCast(reader.Item("Producto"), String)
            Aux.Empresa = DirectCast(reader.Item("Empresa"), String)
            Aux.Etapa = DirectCast(reader.Item("Etapa"), String)
            Aux.Usuario = DirectCast(reader.Item("Usuario"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Actualiza_contraseña_Admin(ByVal id_usuario As Integer, ByVal Contraseña As String) As Boolean Implements IService1.Actualiza_contraseña_Admin

        Dim cmd As New SqlCommand("Actualiza_contraseña_Admin", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@id_Admin", id_usuario)
        cmd.Parameters.AddWithValue("@MD5Contra", Contraseña)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_nombreClientesAdm() As List(Of CCLientesSupervisor) Implements IService1.Obtener_nombreClientesAdm
        Dim Resultado As New List(Of CCLientesSupervisor)
        Dim cmd As New SqlCommand("Obtener_nombreClientesAdm", Conexion)
        ' cmd.Parameters.AddWithValue("@idSupervisor", id_supervisor)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CCLientesSupervisor
        While reader.Read
            Aux = New CCLientesSupervisor
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.Cliente = DirectCast(reader.Item("Cliente"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_nombresUsuariosAdm() As List(Of CSupervisorUsuarios) Implements IService1.Obtener_nombresUsuariosAdm
        Dim Resultado As New List(Of CSupervisorUsuarios)
        Dim cmd As New SqlCommand("Obtener_nombresUsuariosAdm", Conexion)
        'cmd.Parameters.AddWithValue("@idSupervisor", id_supervisor)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CSupervisorUsuarios
        While reader.Read
            Aux = New CSupervisorUsuarios
            Aux.id_usuario = DirectCast(reader.Item("id_usuario"), Integer)
            Aux.Usuario = DirectCast(reader.Item("Usuario"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_supervisor_Adm(ByVal id_supervisor As Integer) As CDetallesSupervisor Implements IService1.Obtener_supervisor_Adm

        Dim cmd As New SqlCommand("Obtener_supervisor_Adm", Conexion)
        cmd.Parameters.AddWithValue("@idAdmin", id_supervisor)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New CDetallesSupervisor
        While reader.Read

            Aux.is_supervisor = DirectCast(reader.Item("id_usuario"), Integer)
            Aux.nombre = DirectCast(reader.Item("nombre"), String)
            Aux.apellidoPaterno = DirectCast(reader.Item("apellidoPaterno"), String)
            Aux.apellidoMaterno = DirectCast(reader.Item("apellidoMaterno"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Aux.usuario = DirectCast(reader.Item("usuario"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)


        End While
        Conexion.Close()
        Return Aux
    End Function
    Function Obtener_clientesAdm() As List(Of ClientesSupervisor) Implements IService1.Obtener_clientesAdm
        Dim Resultado As New List(Of ClientesSupervisor)
        Dim cmd As New SqlCommand("Obtener_clientesAdm", Conexion)
        'cmd.Parameters.AddWithValue("@idSupervisor", idSupervisor)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As ClientesSupervisor
        While reader.Read
            Aux = New ClientesSupervisor
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.Nombre = DirectCast(reader.Item("Nombre"), String)
            Aux.ApellidoPaterno = DirectCast(reader.Item("ApellidoPaterno"), String)
            Aux.ApellidoMaterno = DirectCast(reader.Item("ApellidoMaterno"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Aux.Producto = DirectCast(reader.Item("Producto"), String)
            Aux.Empresa = DirectCast(reader.Item("Empresa"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.Descripcion = DirectCast(reader.Item("Descripcion"), String)
            Aux.Usuario = DirectCast(reader.Item("Usuario"), String)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Aux.fotografia = DirectCast(reader.Item("fotografia"), String)
            Aux.fotoTpresentacion = DirectCast(reader.Item("fotoTpresentacion"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "Status"
    Function Obtener_totalesUsuario(ByVal idUsuario As Integer) As CTotalesUsuario Implements IService1.Obtener_totalesUsuario
        Dim Resultado As New List(Of CidCliente)
        Dim cmd As New SqlCommand("Obtener_AcumuladosClientes", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idUsuario", idUsuario)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New CTotalesUsuario
        While reader.Read

            Aux.ClientesActivos = DirectCast(reader.Item("ClientesActivos"), Integer)
            Aux.ProspectosPorsemana = DirectCast(reader.Item("ProspectosSemana"), Integer)
            Aux.ClientesCancelados = DirectCast(reader.Item("ClientesCancelados"), Integer)
            Aux.ClientesTotal = DirectCast(reader.Item("ClientesTotal"), Integer)


        End While
        Conexion.Close()
        Return Aux
    End Function
#End Region
#Region "LogIn"
    Public Function LogIN(ByVal Usuario As String, ByVal ConstraseñaMD5 As String) As CUsuarios Implements IService1.LogIN
        Dim Resultado As New CUsuarios
        'Intenta ADM
        Try
            Resultado = LogIn_administrativos(Usuario, ConstraseñaMD5)
            If Resultado.id_usuario > 0 Then
                Return Resultado
            End If
        Catch ex As Exception

        End Try

        'Intenta Supervisor
        Try
            Resultado = LogIn_supervisores(Usuario, ConstraseñaMD5)
            If Resultado.id_usuario > 0 Then
                Return Resultado
            End If
        Catch ex As Exception

        End Try

        'Intenta ADM
        Try
            Resultado = LogIn_Usuario(Usuario, ConstraseñaMD5)
            If Resultado.id_usuario > 0 Then
                Return Resultado
            End If
        Catch ex As Exception

        End Try

        Return New CUsuarios
    End Function
    Function LogIn_administrativos(ByVal Usuario As String, ByVal Contraseña As String) As CUsuarios
        Dim Resultado As New CUsuarios
        Dim cmd As New SqlCommand("LogINADM", Conexion)
        cmd.Parameters.AddWithValue("@PUsuario", Usuario)
        cmd.Parameters.AddWithValue("@PContraseñaMD5", Contraseña)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New CUsuarios
        While reader.Read
            Aux.id_usuario = DirectCast(reader.Item("id_usuario"), Integer)
            Aux.nombre = DirectCast(reader.Item("nombre"), String)
            Aux.apellidoMaterno = DirectCast(reader.Item("apellidoMaterno"), String)
            Aux.apellidoPaterno = DirectCast(reader.Item("apellidoPaterno"), String)
            Aux.Email = DirectCast(reader.Item("email"), String)
            Aux.usuario = DirectCast(reader.Item("usuario"), String)
            Aux.contraseña = DirectCast(reader.Item("contraseña"), String)
            Aux.fotografia = DirectCast(reader.Item("fotografia"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.Nivel = DirectCast(reader.Item("Nivel"), Integer)
        End While
        Conexion.Close()
        Return Aux
    End Function
    Function LogIn_supervisores(ByVal Usuario As String, ByVal Contraseña As String) As CUsuarios
        Dim Resultado As New CUsuarios
        Dim cmd As New SqlCommand("LogINSup", Conexion)
        cmd.Parameters.AddWithValue("@PUsuario", Usuario)
        cmd.Parameters.AddWithValue("@PContraseñaMD5", Contraseña)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New CUsuarios
        While reader.Read
            Aux.id_usuario = DirectCast(reader.Item("id_supervisor"), Integer)
            Aux.nombre = DirectCast(reader.Item("nombre"), String)
            Aux.apellidoMaterno = DirectCast(reader.Item("apellidoMaterno"), String)
            Aux.apellidoPaterno = DirectCast(reader.Item("apellidoPaterno"), String)
            Aux.Email = DirectCast(reader.Item("email"), String)
            Aux.usuario = DirectCast(reader.Item("usuario"), String)
            Aux.contraseña = DirectCast(reader.Item("contraseña"), String)
            Aux.fotografia = DirectCast(reader.Item("fotografia"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.Nivel = DirectCast(reader.Item("Nivel"), Integer)
            Aux.BorraEk = Convert.ToUInt32(reader.Item("BorraEk").ToString)
        End While
        Conexion.Close()
        Return Aux
    End Function
    Function LogIn_Usuario(ByVal Usuario As String, ByVal Contraseña As String) As CUsuarios
        Dim Resultado As New CUsuarios
        Dim cmd As New SqlCommand("LogINUsr", Conexion)
        cmd.Parameters.AddWithValue("@PUsuario", Usuario)
        cmd.Parameters.AddWithValue("@PContraseñaMD5", Contraseña)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New CUsuarios
        While reader.Read
            Aux.id_usuario = DirectCast(reader.Item("id_usuario"), Integer)
            Aux.nombre = DirectCast(reader.Item("nombre"), String)
            Aux.apellidoMaterno = DirectCast(reader.Item("apellidoMaterno"), String)
            Aux.apellidoPaterno = DirectCast(reader.Item("apellidoPaterno"), String)
            Aux.Email = DirectCast(reader.Item("email"), String)
            Aux.usuario = DirectCast(reader.Item("usuario"), String)
            Aux.contraseña = DirectCast(reader.Item("contraseña"), String)
            Aux.fotografia = DirectCast(reader.Item("fotografia"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.Nivel = DirectCast(reader.Item("Nivel"), Integer)
        End While
        Conexion.Close()
        Return Aux
    End Function
#End Region
#Region "Clientes"
    Function CompletaCliente(ByVal id_cliente As Integer,
                                ByVal NPersona As String, ByVal NContrato As String,
                                ByVal RFC As String, ByVal NHijos As Integer,
                                ByVal IngresosPersonales As Integer,
                                ByVal Edo_Civil As String) As Boolean Implements IService1.CompletaCliente

        Dim cmd As New SqlCommand("CompletaCliente", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_cliente", id_cliente)
        cmd.Parameters.AddWithValue("@PNPersona", NPersona)
        cmd.Parameters.AddWithValue("@PNContrato", NContrato)
        cmd.Parameters.AddWithValue("@PRFC", RFC)
        cmd.Parameters.AddWithValue("@PNHijos", NHijos)
        cmd.Parameters.AddWithValue("@PIngresosPersonales", IngresosPersonales)
        cmd.Parameters.AddWithValue("@PEdo_Civil", Edo_Civil)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function

    Function ValidaCliente(ByVal nombre As String, ByVal app1 As String, ByVal app2 As String) As Boolean Implements IService1.ValidaCliente
        Dim Resultado As Boolean = False
        Dim cmd As New SqlCommand("ValidaCliente", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Nom", nombre)
        cmd.Parameters.AddWithValue("@app1", app1)
        cmd.Parameters.AddWithValue("@app2", app2)

        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader

        While reader.Read
            If DirectCast(reader.Item(0), Integer) > 0 Then
                Resultado = True
            End If
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function ValidaEmail(ByVal Email As String) As Boolean Implements IService1.ValidaEmail
        Dim Resultado As Boolean = False
        Dim cmd As New SqlCommand("ValidaEmail", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pemail", Email)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As Integer = 0
        While reader.Read

            Aux = reader.Item(0)


        End While
        If Aux > 0 Then
            Resultado = True
        End If
        Conexion.Close()
        Return Resultado
    End Function
    Function Inserta_ClientesCC(ByVal id_cliente As Integer, ByVal id_usuario As Integer, tipoCredito As String) As Boolean Implements IService1.Inserta_ClientesCC
        Dim Resultado As Boolean = False
        Dim cmd As New SqlCommand("inserta_ClientesCC", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idCliente", id_cliente)
        cmd.Parameters.AddWithValue("@idUsuario", id_usuario)
        cmd.Parameters.AddWithValue("@tipoCredito", tipoCredito)

        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_ids_cliente(ByVal id_cliente As Integer) As List(Of CidCliente) Implements IService1.Obtener_ids_cliente
        Dim Resultado As New List(Of CidCliente)
        Dim cmd As New SqlCommand("Obtener_ids_clientes", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@id_cliente", id_cliente)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CidCliente
        While reader.Read
            Aux = New CidCliente
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.Nombre = DirectCast(reader.Item("Nombre"), String)
            Aux.ApellidoPaterno = DirectCast(reader.Item("ApellidoPaterno"), String)
            Aux.ApellidoMaterno = DirectCast(reader.Item("ApellidoMaterno"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Aux.id_producto = DirectCast(reader.Item("id_producto"), Integer)
            Aux.id_nivel = DirectCast(reader.Item("id_nivel"), Integer)
            Aux.id_empresa = DirectCast(reader.Item("id_empresa"), Integer)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.id_etapaActual = DirectCast(reader.Item("id_etapaActual"), Integer)
            Aux.id_campaña = DirectCast(reader.Item("id_campaña"), Integer)
            Aux.id_usuarioOriginal = DirectCast(reader.Item("id_usuarioOriginal"), Integer)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Aux.fotografia = DirectCast(reader.Item("fotografia"), String)
            Aux.fotoTpresentacion = DirectCast(reader.Item("fotoTpresentacion"), String)
            Aux.NSS = DirectCast(reader.Item("NSS"), String)
            Aux.CURP = DirectCast(reader.Item("CURP"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.fechaNacimiento = DirectCast(reader.Item("fechaNacimiento"), Date)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_nombresClientesidUsuario(ByVal id_usuario As Integer) As List(Of CNombresCliente) Implements IService1.Obtener_nombresClientesidUsuario
        Dim Resultado As New List(Of CNombresCliente)
        Dim cmd As New SqlCommand("Obtener_nombresClientesidUsuario", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idUsuario", id_usuario)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CNombresCliente
        While reader.Read
            Aux = New CNombresCliente
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.Cliente = DirectCast(reader.Item("Cliente"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function VerificaCliente(ByVal idcliente As Integer, ByVal idusuario As Integer) As Boolean Implements IService1.VerificaCliente
        Dim Resultado As Boolean = False
        Dim cmd As New SqlCommand("Verifica_cliente_usuario", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idCliente", idcliente)
        cmd.Parameters.AddWithValue("@idUsuario", idusuario)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader

        While reader.Read
            If DirectCast(reader.Item(0), Integer) > 0 Then
                Resultado = True
            End If
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Actualiza_ultimafecha(ByVal idcliente As Integer) As Boolean Implements IService1.Actualiza_ultimafecha
        Dim Resultado As Boolean = False
        Dim cmd As New SqlCommand("Actualiza_ultimafecha", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idCliente", idcliente)

        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function ComprobarNSS(ByVal NSS As String) As Integer Implements IService1.ComprobarNSS
        Dim Resultado As New Integer
        Dim cmd As New SqlCommand("ComprobarNSS", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@NSS", NSS)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        While reader.Read
            Resultado = DirectCast(reader.Item("id_cliente"), Integer)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function ComprobarCURP(ByVal CURP As String) As Integer Implements IService1.ComprobarCURP
        Dim Resultado As New Integer
        Dim cmd As New SqlCommand("ComprobarCURP", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Curp", CURP)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        While reader.Read
            Resultado = DirectCast(reader.Item("id_cliente"), Integer)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Inserta_clientes(ByVal Nombre As String, ByVal ApellidoPaterno As String, ByVal ApellidoMaterno As String, ByVal Email As String, ByVal id_producto As Integer, ByVal id_nivel As Integer, ByVal id_empresa As Integer, ByVal id_etapaActual As Integer, ByVal id_campaña As Integer, ByVal id_usuarioOriginal As Integer, ByVal Observaciones As String, ByVal fotografia As String, ByVal fotoTpresentacion As String, ByVal NSS As String, ByVal CURP As String, ByVal FechaNacimiento As Date, ByVal id_cve_fracc As String) As Integer Implements IService1.Inserta_clientes

        Dim resultado As Integer = 0
        'If ValidaCliente(Nombre, ApellidoMaterno, ApellidoPaterno) Then
        'Else
        Dim cmd As New SqlCommand("Inserta_Cliente", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@PNombre", Nombre)
        cmd.Parameters.AddWithValue("@PApellidoPaterno", ApellidoPaterno)
        cmd.Parameters.AddWithValue("@PApellidoMaterno", ApellidoMaterno)
        cmd.Parameters.AddWithValue("@PEmail", Email)
        cmd.Parameters.AddWithValue("@Pid_producto", id_producto)
        cmd.Parameters.AddWithValue("@Pid_nivel", id_nivel)
        cmd.Parameters.AddWithValue("@Pid_empresa", id_empresa)
        cmd.Parameters.AddWithValue("@Pid_etapaActual", id_etapaActual)
        cmd.Parameters.AddWithValue("@Pid_campaña", id_campaña)
        cmd.Parameters.AddWithValue("@Pid_usuarioOriginal", id_usuarioOriginal)
        cmd.Parameters.AddWithValue("@PObservaciones", Observaciones)
        cmd.Parameters.AddWithValue("@Pfotografia", fotografia)
        cmd.Parameters.AddWithValue("@PfotoTpresentacion", fotoTpresentacion)
        cmd.Parameters.AddWithValue("@PNSS", NSS)
        cmd.Parameters.AddWithValue("@PCURP", CURP)
        cmd.Parameters.AddWithValue("@PfechaNacimiento", FechaNacimiento)
        cmd.Parameters.AddWithValue("@id_fracc", id_cve_fracc)
        Conexion.Close()
        Conexion.Open()
        Try
            Dim reader As SqlDataReader = cmd.ExecuteReader

            While reader.Read
                resultado = reader.Item(0)
            End While
        Catch ex As Exception
            resultado = 0
        End Try


        Conexion.Close()
        'End If

        Return resultado
    End Function
    Function Inserta_email_adjuntos(ByVal filename As String, ByVal MediaType As String, ByVal id_email As Integer, ByVal Body64Str As String) As Boolean Implements IService1.Inserta_email_adjuntos

        Dim cmd As New SqlCommand("Inserta_adjunto", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pfilename", filename)
        cmd.Parameters.AddWithValue("@PMediaType", MediaType)
        cmd.Parameters.AddWithValue("@Pid_email", id_email)
        cmd.Parameters.AddWithValue("@PBody64Str", Body64Str)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_clientes(ByVal id_cliente As Integer) As Boolean Implements IService1.Elimina_clientes

        Dim cmd As New SqlCommand("Elimina_clientes", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_cliente", id_cliente)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_clientes(ByVal id_cliente As Integer, ByVal Nombre As String, ByVal ApellidoPaterno As String, ByVal ApellidoMaterno As String, ByVal Email As String, ByVal id_producto As Integer, ByVal id_nivel As Integer, ByVal id_empresa As Integer, ByVal id_campaña As Integer, ByVal Observaciones As String, ByVal fotografia As String, ByVal fotoTpresentacion As String, ByVal Monto As Decimal, ByVal Id_Usr As Integer) As Boolean Implements IService1.Actualiza_clientes

        Dim cmd As New SqlCommand("Actualiza_clientes", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_cliente", id_cliente)
        cmd.Parameters.AddWithValue("@PNombre", Nombre)
        cmd.Parameters.AddWithValue("@PApellidoPaterno", ApellidoPaterno)
        cmd.Parameters.AddWithValue("@PApellidoMaterno", ApellidoMaterno)
        cmd.Parameters.AddWithValue("@PEmail", Email)
        cmd.Parameters.AddWithValue("@Pid_producto", id_producto)
        cmd.Parameters.AddWithValue("@Pid_nivel", id_nivel)
        cmd.Parameters.AddWithValue("@Pid_empresa", id_empresa)
        'cmd.Parameters.AddWithValue("@PfechaCreacion", fechaCreacion)
        'cmd.Parameters.AddWithValue("@Pid_etapaActual", id_etapaActual)
        cmd.Parameters.AddWithValue("@Pid_campaña", id_campaña)
        'cmd.Parameters.AddWithValue("@Pid_usuarioOriginal", id_usuarioOriginal)
        cmd.Parameters.AddWithValue("@PObservaciones", Observaciones)
        cmd.Parameters.AddWithValue("@Pfotografia", fotografia)
        cmd.Parameters.AddWithValue("@PfotoTpresentacion", fotoTpresentacion)
        cmd.Parameters.AddWithValue("@PMonto", Monto)
        cmd.Parameters.AddWithValue("@Id_Usr", Id_Usr)

        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_clientes_callcenter(ByVal id_cliente As Integer, ByVal Nombre As String, ByVal ApellidoPaterno As String, ByVal ApellidoMaterno As String, ByVal Email As String,
                                           ByVal id_producto As Integer, ByVal id_nivel As Integer, ByVal id_empresa As Integer, ByVal id_campaña As Integer, ByVal Observaciones As String,
                                           ByVal fotografia As String, ByVal fotoTpresentacion As String, ByVal Monto As Decimal, ByVal NSS As String, ByVal FechaNacimiento As Date, ByVal Id_Usr As Integer) As Boolean Implements IService1.Actualiza_clientes_callcenter

        Dim cmd As New SqlCommand("Actualiza_clientes_callcenter", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_cliente", id_cliente)
        cmd.Parameters.AddWithValue("@PNombre", Nombre)
        cmd.Parameters.AddWithValue("@PApellidoPaterno", ApellidoPaterno)
        cmd.Parameters.AddWithValue("@PApellidoMaterno", ApellidoMaterno)
        cmd.Parameters.AddWithValue("@PEmail", Email)
        cmd.Parameters.AddWithValue("@Pid_producto", id_producto)
        cmd.Parameters.AddWithValue("@Pid_nivel", id_nivel)
        cmd.Parameters.AddWithValue("@Pid_empresa", id_empresa)
        cmd.Parameters.AddWithValue("@Pid_campaña", id_campaña)
        cmd.Parameters.AddWithValue("@PObservaciones", Observaciones)
        cmd.Parameters.AddWithValue("@Pfotografia", fotografia)
        cmd.Parameters.AddWithValue("@PfotoTpresentacion", fotoTpresentacion)
        cmd.Parameters.AddWithValue("@PMonto", Monto)
        cmd.Parameters.AddWithValue("@PNSS", NSS)
        cmd.Parameters.AddWithValue("@PFechaNacimiento", FechaNacimiento)
        cmd.Parameters.AddWithValue("@Id_Usr", Id_Usr)

        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_ClienteObservaciones(ByVal idCliente As Integer) As List(Of CClienteObservaciones) Implements IService1.Obtener_ClienteObservaciones
        Dim Resultado As New List(Of CClienteObservaciones)
        Dim cmd As New SqlCommand("Obtener_Clientes_Observaciones_idCliente", Conexion)
        Dim Aux As CClienteObservaciones

        cmd.Parameters.AddWithValue("@id_Cliente", idCliente)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()

        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        While reader.Read
            Aux = New CClienteObservaciones
            Aux.Observacion = DirectCast(reader.Item("Observacion"), String)
            Aux.Fecha_Registro = DirectCast(reader.Item("Creacion"), DateTime)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado

    End Function
    Function Actualiza_rankingCliente(ByVal id_cliente As Integer, ByVal ranking As String) As Boolean Implements IService1.Actualiza_rankingCliente

        Dim cmd As New SqlCommand("Actualiza_rankingCliente", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idCliente", id_cliente)
        cmd.Parameters.AddWithValue("@ranking", ranking)

        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualizar_Ranking(ByVal id_cliente As Integer, ByVal id_usuario As Integer, ByVal ranking_org As String, ByVal ranking_nvo As String) As Boolean Implements IService1.Actualizar_Ranking
        Dim cmd As New SqlCommand("Actualizar_Ranking", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idCliente", id_cliente)
        cmd.Parameters.AddWithValue("@idUsuario", id_usuario)
        cmd.Parameters.AddWithValue("@rankingOrg", ranking_org)
        cmd.Parameters.AddWithValue("@rankingNvo", ranking_nvo)

        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualizar_Ranking_Visitas(ByVal id_cliente As Integer, ByVal id_usuario As Integer, ByVal ranking_org As String, ByVal ranking_nvo As String, ByVal id_Visita As Integer, ByVal id_Impedimento As Integer) As Boolean Implements IService1.Actualizar_Ranking_Visitas
        Dim cmd As New SqlCommand("Actualizar_Ranking_Visitas", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idCliente", id_cliente)
        cmd.Parameters.AddWithValue("@idUsuario", id_usuario)
        cmd.Parameters.AddWithValue("@rankingOrg", ranking_org)
        cmd.Parameters.AddWithValue("@rankingNvo", ranking_nvo)
        cmd.Parameters.AddWithValue("@id_Visita", id_Visita)
        cmd.Parameters.AddWithValue("@id_Impedimento", id_Impedimento)

        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function

    Function Obtener_clientes_Todos() As List(Of CClientes) Implements IService1.Obtener_clientes_Todos
        Dim Resultado As New List(Of CClientes)
        Dim cmd As New SqlCommand("Obtener_clientes_Todos", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CClientes
        While reader.Read
            Aux = New CClientes
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.Nombre = DirectCast(reader.Item("Nombre"), String)
            Aux.ApellidoPaterno = DirectCast(reader.Item("ApellidoPaterno"), String)
            Aux.ApellidoMaterno = DirectCast(reader.Item("ApellidoMaterno"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Aux.id_producto = DirectCast(reader.Item("id_producto"), Integer)
            Aux.id_nivel = DirectCast(reader.Item("id_nivel"), Integer)
            Aux.id_empresa = DirectCast(reader.Item("id_empresa"), Integer)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.id_etapaActual = DirectCast(reader.Item("id_etapaActual"), Integer)
            Aux.id_campaña = DirectCast(reader.Item("id_campaña"), Integer)
            Aux.id_usuarioOriginal = DirectCast(reader.Item("id_usuarioOriginal"), Integer)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Aux.fotografia = DirectCast(reader.Item("fotografia"), String)
            Aux.fotoTpresentacion = DirectCast(reader.Item("fotoTpresentacion"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_clientes_detalles_todos() As List(Of CClientesDetalles) Implements IService1.Obtener_clientes_detalles_todos
        Dim Resultado As New List(Of CClientesDetalles)
        Dim cmd As New SqlCommand("Obtener_clientes_detalles_todos", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CClientesDetalles
        While reader.Read
            Aux = New CClientesDetalles
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.Nombre = DirectCast(reader.Item("Nombre"), String)
            Aux.ApellidoPaterno = DirectCast(reader.Item("ApellidoPaterno"), String)
            Aux.ApellidoMaterno = DirectCast(reader.Item("ApellidoMaterno"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Aux.NombreCorto = DirectCast(reader.Item("NombreCorto"), String)
            Aux.nivelinteres = DirectCast(reader.Item("nivelinteres"), String)
            Aux.Empresa = DirectCast(reader.Item("Empresa"), String)
            Aux.campañaNombre = DirectCast(reader.Item("campañaNombre"), String)
            Aux.NombreAsesor = DirectCast(reader.Item("NombreAsesor"), String)
            Aux.ApellidoAsesor = DirectCast(reader.Item("ApellidoAsesor"), String)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Aux.fotografia = DirectCast(reader.Item("fotografia"), String)
            Aux.fotoTpresentacion = DirectCast(reader.Item("fotoTpresentacion"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.id_etapaActual = DirectCast(reader.Item("id_etapaActual"), Integer)
            Aux.id_producto = DirectCast(reader.Item("id_producto"), Integer)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_clientes_detalles_idUsuario(ByVal id_usuario As Integer) As List(Of CClientesDetalles) Implements IService1.Obtener_clientes_detalles_idUsuario
        Dim Resultado As New List(Of CClientesDetalles)
        Dim cmd As New SqlCommand("Obtener_Clientes_detalles_usuario", Conexion)
        cmd.Parameters.AddWithValue("@IdUsuario", id_usuario)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CClientesDetalles
        While reader.Read
            Aux = New CClientesDetalles
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.Nombre = DirectCast(reader.Item("Nombre"), String)
            Aux.ApellidoPaterno = DirectCast(reader.Item("ApellidoPaterno"), String)
            Aux.ApellidoMaterno = DirectCast(reader.Item("ApellidoMaterno"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Aux.NombreCorto = DirectCast(reader.Item("NombreCorto"), String)
            Aux.nivelinteres = DirectCast(reader.Item("nivelinteres"), String)
            Aux.Empresa = DirectCast(reader.Item("Empresa"), String)
            Aux.campañaNombre = DirectCast(reader.Item("campañaNombre"), String)
            Aux.NombreAsesor = DirectCast(reader.Item("NombreAsesor"), String)
            Aux.ApellidoAsesor = DirectCast(reader.Item("ApellidoAsesor"), String)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Aux.fotografia = DirectCast(reader.Item("fotografia"), String)
            Aux.fotoTpresentacion = DirectCast(reader.Item("fotoTpresentacion"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.id_etapaActual = DirectCast(reader.Item("id_etapaActual"), Integer)
            Aux.Etapa = DirectCast(reader.Item("Etapa"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_Clientes_Telefonos_idCliente(ByVal id_cliente As Integer) As List(Of CClienteTelefonos) Implements IService1.Obtener_Clientes_Telefonos_idCliente
        Dim Resultado As New List(Of CClienteTelefonos)
        Dim cmd As New SqlCommand("Obtener_Clientes_Telefonos_idCliente", Conexion)

        cmd.Parameters.AddWithValue("@idCliente", id_cliente)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()

        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CClienteTelefonos
        While reader.Read
            Aux = New CClienteTelefonos
            Aux.Telefono = DirectCast(reader.Item("Telefono"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()

        Return Resultado
    End Function
    Function Obtener_Clientes_detalles_idCliente(ByVal id_cliente As Integer) As List(Of CClientesDetalles) Implements IService1.Obtener_Clientes_detalles_idCliente
        Dim Resultado As New List(Of CClientesDetalles)
        Dim cmd As New SqlCommand("Obtener_Clientes_detalles_idCliente", Conexion)

        cmd.Parameters.AddWithValue("@idCliente", id_cliente)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CClientesDetalles
        While reader.Read
            Aux = New CClientesDetalles
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.Nombre = DirectCast(reader.Item("Nombre"), String)
            Aux.ApellidoPaterno = DirectCast(reader.Item("ApellidoPaterno"), String)
            Aux.ApellidoMaterno = DirectCast(reader.Item("ApellidoMaterno"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Aux.NombreCorto = DirectCast(reader.Item("NombreCorto"), String)
            Aux.nivelinteres = DirectCast(reader.Item("nivelinteres"), String)
            Aux.Empresa = DirectCast(reader.Item("Empresa"), String)
            Aux.Id_Campaña = DirectCast(reader.Item("id_campaña"), Integer)
            Aux.campañaNombre = DirectCast(reader.Item("campañaNombre"), String)
            Aux.tipoCampana = DirectCast(reader.Item("TipoCampaña"), String)
            Aux.NombreAsesor = DirectCast(reader.Item("NombreAsesor"), String)
            Aux.ApellidoAsesor = DirectCast(reader.Item("ApellidoAsesor"), String)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Aux.NSS = DirectCast(reader.Item("NSS"), String)
            Aux.CURP = DirectCast(reader.Item("CURP"), String)
            Aux.fotografia = DirectCast(reader.Item("fotografia"), String)
            Aux.fotoTpresentacion = DirectCast(reader.Item("fotoTpresentacion"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.fechaNacimiento = DirectCast(reader.Item("fechaNacimiento"), Date)
            Aux.id_etapaActual = DirectCast(reader.Item("id_etapaActual"), Integer)
            Aux.id_producto = DirectCast(reader.Item("id_producto"), Integer)
            Aux.ranking = DirectCast(reader.Item("ranking"), String)
            Aux.Numcte = DirectCast(reader.Item("Numcte"), Integer)
            Aux.Numcte2 = DirectCast(reader.Item("Numcte2"), Integer)
            Aux.id_Usuario = DirectCast(reader.Item("IdUsuario"), Integer)
            Aux.ModeloEk = DirectCast(reader.Item("ModeloEk"), String)

            If IsDBNull(reader.Item("fecha_cierre")) Then
                Aux.FechaCierre = "1900-01-01"
            Else
                Aux.FechaCierre = DirectCast(reader.Item("fecha_cierre"), Date)
            End If

            If IsDBNull(reader.Item("fecha_escritura")) Then
                Aux.FechaEscritura = "1900-01-01"
            Else
                Aux.FechaEscritura = DirectCast(reader.Item("fecha_escritura"), Date)
            End If

            If IsDBNull(reader.Item("fecha_cancelacion")) Then
                Aux.FechaCancelacion = "1900-01-01"
            Else
                Aux.FechaCancelacion = DirectCast(reader.Item("fecha_cancelacion"), Date)
            End If

            If (String.IsNullOrEmpty(reader.Item("empresaEK").ToString)) Then
                Aux.EmpresaEK = 0
            Else
                Aux.EmpresaEK = DirectCast(reader.Item("empresaEK"), Integer)
            End If

            If IsDBNull(reader.Item("Fecha_Recuperacion")) Then
                Aux.Fecha_Recuperacion = "1900-01-01"
            Else
                Aux.Fecha_Recuperacion = DirectCast(reader.Item("Fecha_Recuperacion"), Date)
            End If

            If IsDBNull(reader.Item("FechaOperacionEk")) Then
                Aux.Fecha_OperacionEK = "1900-01-01"
            Else
                Aux.Fecha_OperacionEK = DirectCast(reader.Item("FechaOperacionEk"), Date)
            End If


            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function

    Function Obtener_Clientes_AsesorCallCenter(ByVal id_cliente As Integer) As List(Of AsesorCallCenter) Implements IService1.Obtener_Clientes_AsesorCallCenter
        Dim Resultado As New List(Of AsesorCallCenter)
        Dim cmd As New SqlCommand("Obtener_Clientes_AsesorCallCenter", Conexion)

        cmd.Parameters.AddWithValue("@idCliente", id_cliente)
        cmd.CommandType = CommandType.StoredProcedure

        Conexion.Close()
        Conexion.Open()
        Dim Reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As AsesorCallCenter
        While Reader.Read
            Aux = New AsesorCallCenter
            Aux.id_usuario = DirectCast(Reader.Item("id_usuario"), Integer)
            Aux.nombre = DirectCast(Reader.Item("nombre"), String)
            Aux.apellidoPaterno = DirectCast(Reader.Item("apellidoPaterno"), String)
            Aux.apellidoMaterno = DirectCast(Reader.Item("apellidoMaterno"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()

        Return Resultado
    End Function
    Public Function Obtener_Clientes_TipoCredito_idCliente(ByVal id_cliente As Integer) As List(Of CClientesTipoCredito) Implements IService1.Obtener_Clientes_TipoCredito_idCliente
        Dim Resultado As New List(Of CClientesTipoCredito)
        Dim cmd As New SqlCommand("Obtener_Clientes_TipoCredito_idCliente", Conexion)

        cmd.Parameters.AddWithValue("@idCliente", id_cliente)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()

        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CClientesTipoCredito
        While reader.Read
            Aux = New CClientesTipoCredito
            Aux.TipoCredito = DirectCast(reader.Item("tipocredito"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()

        Return Resultado
    End Function
#End Region
#Region "Campañas"
    Function Obtener_combo_campañas() As List(Of CComboCampañas) Implements IService1.Obtener_combo_campañas
        Dim Resultado As New List(Of CComboCampañas)
        Dim cmd As New SqlCommand("Obtener_campañasCombo", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CComboCampañas
        While reader.Read
            Aux = New CComboCampañas
            Aux.id_campaña = DirectCast(reader.Item("id_campaña"), Integer)
            Aux.Campaña = DirectCast(reader.Item("Campaña"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Inserta_campañas(ByVal campañaNombre As String, ByVal id_tipoCampaña As Integer, ByVal id_MedioCampaña As Integer, ByVal fechaInicio As Date, ByVal fechaFinal As Date, ByVal Observaciones As String) As Boolean Implements IService1.Inserta_campañas

        Dim cmd As New SqlCommand("Inserta_Campaña", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@PcampañaNombre", campañaNombre)
        cmd.Parameters.AddWithValue("@Pid_tipoCampaña", id_tipoCampaña)
        cmd.Parameters.AddWithValue("@Pid_MedioCampaña", id_MedioCampaña)
        cmd.Parameters.AddWithValue("@PfechaInicio", fechaInicio)
        cmd.Parameters.AddWithValue("@PfechaFinal", fechaFinal)
        cmd.Parameters.AddWithValue("@PObservaciones", Observaciones)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_campañas(ByVal id_campaña As Integer) As Boolean Implements IService1.Elimina_campañas

        Dim cmd As New SqlCommand("Elimina_campañas", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_campaña", id_campaña)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_campañas(ByVal id_campaña As Integer, ByVal campañaNombre As String, ByVal id_tipoCampaña As Integer, ByVal fechaCreacion As Date, ByVal fechaInicio As Date, ByVal fechaFinal As Date, ByVal Observaciones As String) As Boolean Implements IService1.Actualiza_campañas

        Dim cmd As New SqlCommand("Actualiza_campañas", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_campaña", id_campaña)
        cmd.Parameters.AddWithValue("@PcampañaNombre", campañaNombre)
        cmd.Parameters.AddWithValue("@Pid_tipoCampaña", id_tipoCampaña)
        cmd.Parameters.AddWithValue("@PfechaCreacion", fechaCreacion)
        cmd.Parameters.AddWithValue("@PfechaInicio", fechaInicio)
        cmd.Parameters.AddWithValue("@PfechaFinal", fechaFinal)
        cmd.Parameters.AddWithValue("@PObservaciones", Observaciones)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_campañas() As List(Of CCampaña) Implements IService1.Obtener_campañas
        Dim Resultado As New List(Of CCampaña)
        Dim cmd As New SqlCommand("Obtener_campañas_todas", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CCampaña
        While reader.Read
            Aux = New CCampaña
            Aux.id_campaña = DirectCast(reader.Item("id_campaña"), Integer)
            Aux.campañaNombre = DirectCast(reader.Item("campañaNombre"), String)
            Aux.id_tipoCampaña = DirectCast(reader.Item("id_tipoCampaña"), Integer)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.fechaInicio = DirectCast(reader.Item("fechaInicio"), Date)
            Aux.fechaFinal = DirectCast(reader.Item("fechaFinal"), Date)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_campañasDetalles() As List(Of CCampañaDetalles) Implements IService1.Obtener_campañasDetalles
        Dim Resultado As New List(Of CCampañaDetalles)
        Dim cmd As New SqlCommand("Obtener_campañasDetalles", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CCampañaDetalles
        While reader.Read
            Aux = New CCampañaDetalles
            Aux.id_campaña = DirectCast(reader.Item("id_campaña"), Integer)
            Aux.campañaNombre = DirectCast(reader.Item("campañaNombre"), String)
            Aux.tipoCampaña = DirectCast(reader.Item("TipoCampaña"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.fechaInicio = DirectCast(reader.Item("FechaInic"), Date)
            Aux.fechaFinal = DirectCast(reader.Item("FechaFin"), Date)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_campañas_idCampaña(ByVal id_campaña As Integer) As CCampaña Implements IService1.Obtener_campañas_idCampaña
        Dim cmd As New SqlCommand("Obtener_campañas_todas", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New CCampaña
        While reader.Read
            Aux.id_campaña = DirectCast(reader.Item("id_campaña"), Integer)
            Aux.campañaNombre = DirectCast(reader.Item("campañaNombre"), String)
            Aux.id_tipoCampaña = DirectCast(reader.Item("id_tipoCampaña"), Integer)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.fechaInicio = DirectCast(reader.Item("fechaInicio"), Date)
            Aux.fechaFinal = DirectCast(reader.Item("fechaFinal"), Date)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)

        End While
        Conexion.Close()
        Return Aux
    End Function

#End Region
#Region "Categorias"
    Function Inserta_categoriasProductos(ByVal categoria As String) As Boolean Implements IService1.Inserta_categoriasProductos

        Dim cmd As New SqlCommand("Inserta_categoriaProducto", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pcategoria", categoria)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_categoriasProductos(ByVal id_categoria As Integer) As Boolean Implements IService1.Elimina_categoriasProductos

        Dim cmd As New SqlCommand("Elimina_categoriasProductos", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_categoria", id_categoria)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_categoriasProductos(ByVal id_categoria As Integer, ByVal categoria As String) As Boolean Implements IService1.Actualiza_categoriasProductos

        Dim cmd As New SqlCommand("Actualiza_categoriasProductos", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_categoria", id_categoria)
        cmd.Parameters.AddWithValue("@Pcategoria", categoria)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_categoriasProductos() As List(Of CCategoriasProducto) Implements IService1.Obtener_categoriasProductos
        Dim Resultado As New List(Of CCategoriasProducto)
        Dim cmd As New SqlCommand("Obtener_categoriasProductos", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CCategoriasProducto
        While reader.Read
            Aux = New CCategoriasProducto
            Aux.id_categoria = DirectCast(reader.Item("id_categoria"), Integer)
            Aux.categoria = DirectCast(reader.Item("categoria"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "Citas"
    Function Insertar_Cita(ByVal IdCliente As Integer, ByVal IdUsuario As Integer, ByVal IdUsuarioAsignado As Integer, ByVal IdCampana As Integer, ByVal TipoCampana As String,
                           ByVal Origen As String, ByVal LugarContacto As String, ByVal Proyecto As String, ByVal Modelo As Integer, ByVal VigenciaInicial As Date,
                           ByVal VigenciaFinal As Date, ByVal FechaCita As Date, ByVal Ranking As String, ByVal Status As Integer) As Boolean Implements IService1.Insertar_Cita

        Dim cmd As New SqlCommand("Insertar_Cita", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@pId_Cliente", IdCliente)
        cmd.Parameters.AddWithValue("@pId_Usuario", IdUsuario)
        cmd.Parameters.AddWithValue("@pId_UsuarioAsignado", IdUsuarioAsignado)
        cmd.Parameters.AddWithValue("@pId_Camapana", IdCampana)
        cmd.Parameters.AddWithValue("@pTipoCampana", TipoCampana)
        cmd.Parameters.AddWithValue("@pOrigen", Origen)
        cmd.Parameters.AddWithValue("@pLugar_Contacto", LugarContacto)
        cmd.Parameters.AddWithValue("@pProyectoVisita", Proyecto)
        cmd.Parameters.AddWithValue("@pModelo", Modelo)
        cmd.Parameters.AddWithValue("@pVigenciaInicio", VigenciaInicial)
        cmd.Parameters.AddWithValue("@pVigenciaFinal", VigenciaFinal)
        cmd.Parameters.AddWithValue("@pFechaCita", FechaCita)
        cmd.Parameters.AddWithValue("@pRanking", Ranking)
        cmd.Parameters.AddWithValue("@pEstatus", Status)

        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try

        Conexion.Close()
        Return False
    End Function

    Function Insertar_CitasCallCenter(ByVal IdCliente As Integer, ByVal IdUsuario As Integer, ByVal IdUsuarioAsignado As Integer, ByVal IdCampana As Integer, ByVal TipoCampana As String,
                           ByVal Origen As String, ByVal LugarContacto As String, ByVal Proyecto As String, ByVal Modelo As Integer, ByVal VigenciaInicial As Date,
                           ByVal VigenciaFinal As Date, ByVal FechaCita As Date, ByVal Ranking As String, ByVal Status As Integer) As Boolean Implements IService1.Insertar_CitasCallCenter

        Dim cmd As New SqlCommand("Insertar_CitasCallCenter", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@pId_Cliente", IdCliente)
        cmd.Parameters.AddWithValue("@pId_Usuario", IdUsuario)
        cmd.Parameters.AddWithValue("@pId_UsuarioAsignado", IdUsuarioAsignado)
        cmd.Parameters.AddWithValue("@pId_Camapana", IdCampana)
        cmd.Parameters.AddWithValue("@pTipoCampana", TipoCampana)
        cmd.Parameters.AddWithValue("@pOrigen", Origen)
        cmd.Parameters.AddWithValue("@pLugar_Contacto", LugarContacto)
        cmd.Parameters.AddWithValue("@pProyectoVisita", Proyecto)
        cmd.Parameters.AddWithValue("@pModelo", Modelo)
        cmd.Parameters.AddWithValue("@pVigenciaInicio", VigenciaInicial)
        cmd.Parameters.AddWithValue("@pVigenciaFinal", VigenciaFinal)
        cmd.Parameters.AddWithValue("@pFechaCita", FechaCita)
        cmd.Parameters.AddWithValue("@pRanking", Ranking)
        cmd.Parameters.AddWithValue("@pEstatus", Status)

        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try

        Conexion.Close()
        Return False
    End Function

    Function Insertar_CitasProspectador(ByVal IdCliente As Integer, ByVal IdUsuario As Integer, ByVal IdUsuarioAsignado As Integer, ByVal IdCampana As Integer, ByVal TipoCampana As String,
                           ByVal Origen As String, ByVal LugarContacto As String, ByVal Proyecto As String, ByVal Modelo As Integer, ByVal VigenciaInicial As Date,
                           ByVal VigenciaFinal As Date, ByVal FechaCita As Date, ByVal Ranking As String, ByVal Status As Integer) As Boolean Implements IService1.Insertar_CitasProspectador

        Dim cmd As New SqlCommand("Insertar_CitasProspectador", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@pId_Cliente", IdCliente)
        cmd.Parameters.AddWithValue("@pId_Usuario", IdUsuario)
        cmd.Parameters.AddWithValue("@pId_UsuarioAsignado", IdUsuarioAsignado)
        cmd.Parameters.AddWithValue("@pId_Camapana", IdCampana)
        cmd.Parameters.AddWithValue("@pTipoCampana", TipoCampana)
        cmd.Parameters.AddWithValue("@pOrigen", Origen)
        cmd.Parameters.AddWithValue("@pLugar_Contacto", LugarContacto)
        cmd.Parameters.AddWithValue("@pProyectoVisita", Proyecto)
        cmd.Parameters.AddWithValue("@pModelo", Modelo)
        cmd.Parameters.AddWithValue("@pVigenciaInicio", VigenciaInicial)
        cmd.Parameters.AddWithValue("@pVigenciaFinal", VigenciaFinal)
        cmd.Parameters.AddWithValue("@pFechaCita", FechaCita)
        cmd.Parameters.AddWithValue("@pRanking", Ranking)
        cmd.Parameters.AddWithValue("@pEstatus", Status)

        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try

        Conexion.Close()
        Return False
    End Function

    Function Insertar_CitasCaseta(ByVal IdCliente As Integer, ByVal IdUsuario As Integer, ByVal IdUsuarioAsignado As Integer, ByVal IdCampana As Integer, ByVal TipoCampana As String,
                           ByVal Origen As String, ByVal LugarContacto As String, ByVal Proyecto As String, ByVal Modelo As Integer, ByVal VigenciaInicial As Date,
                           ByVal VigenciaFinal As Date, ByVal FechaCita As Date, ByVal Ranking As String, ByVal Status As Integer) As Boolean Implements IService1.Insertar_CitasCaseta

        Dim cmd As New SqlCommand("Insertar_CitasCaseta", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@pId_Cliente", IdCliente)
        cmd.Parameters.AddWithValue("@pId_Usuario", IdUsuario)
        cmd.Parameters.AddWithValue("@pId_UsuarioAsignado", IdUsuarioAsignado)
        cmd.Parameters.AddWithValue("@pId_Camapana", IdCampana)
        cmd.Parameters.AddWithValue("@pTipoCampana", TipoCampana)
        cmd.Parameters.AddWithValue("@pOrigen", Origen)
        cmd.Parameters.AddWithValue("@pLugar_Contacto", LugarContacto)
        cmd.Parameters.AddWithValue("@pProyectoVisita", Proyecto)
        cmd.Parameters.AddWithValue("@pModelo", Modelo)
        cmd.Parameters.AddWithValue("@pVigenciaInicio", VigenciaInicial)
        cmd.Parameters.AddWithValue("@pVigenciaFinal", VigenciaFinal)
        cmd.Parameters.AddWithValue("@pFechaCita", FechaCita)
        cmd.Parameters.AddWithValue("@pRanking", Ranking)
        cmd.Parameters.AddWithValue("@pEstatus", Status)

        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try

        Conexion.Close()
        Return False
    End Function

    Function Insertar_CitaCallCenter(ByVal IdCliente As Integer, ByVal IdUsuario As Integer, ByVal IdUsuarioAsignado As Integer, ByVal Origen As String,
                                     ByVal LugarContacto As String, ByVal Proyecto As String, ByVal Modelo As Integer, ByVal VigenciaInicial As Date,
                                     ByVal VigenciaFinal As Date, ByVal FechaCita As Date, ByVal Estatus As String, ByVal IdCampana As Integer,
                                     ByVal TipoCampana As String, ByVal Activa As Integer) As Boolean Implements IService1.Insertar_CitaCallCenter

        Dim cmd As New SqlCommand("Insertar_CitaCallCenter", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@pId_Cliente", IdCliente)
        cmd.Parameters.AddWithValue("@pId_UsuarioCC", IdUsuario)
        cmd.Parameters.AddWithValue("@pId_UsuarioAsesor", IdUsuarioAsignado)
        cmd.Parameters.AddWithValue("@pId_Camapana", IdCampana)
        cmd.Parameters.AddWithValue("@pTipoCampana", TipoCampana)
        cmd.Parameters.AddWithValue("@pOrigen", Origen)
        cmd.Parameters.AddWithValue("@pLugar_Contacto", LugarContacto)
        cmd.Parameters.AddWithValue("@pProyectoVisita", Proyecto)
        cmd.Parameters.AddWithValue("@pModelo", Modelo)
        cmd.Parameters.AddWithValue("@pVigenciaInicio", VigenciaInicial)
        cmd.Parameters.AddWithValue("@pVigenciaFinal", VigenciaFinal)
        cmd.Parameters.AddWithValue("@pFechaCita", FechaCita)
        cmd.Parameters.AddWithValue("@pEstatus", Estatus)
        cmd.Parameters.AddWithValue("@pActivo", Activa)

        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try

        Conexion.Close()
        Return False
    End Function

    Function Insertar_ObservacionesCitas(ByVal IdCita As Integer, ByVal IdUsuario As Integer, ByVal Completada As Integer, ByVal Observaciones As String) As Boolean Implements IService1.Insertar_ObservacionesCitas
        Dim cmd As New SqlCommand("Insertar_ObservacionesCitas", Conexion)

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@pId_Cita", IdCita)
        cmd.Parameters.AddWithValue("@pId_Usuario", IdUsuario)
        cmd.Parameters.AddWithValue("@pCompletada", Completada)
        cmd.Parameters.AddWithValue("@pObservaciones", Observaciones)

        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try

        Conexion.Close()
        Return False
    End Function

    Function Inserta_CitasCall(ByVal id_cliente As Integer, ByVal id_usuarioCC As Integer, ByVal id_usuarioAsesor As Integer, ByVal Origen As String, ByVal Lugar_Contacto As String, ByVal ProyectoVisita As String, ByVal Modelo As String, ByVal VigenciaInicio As Date, ByVal VigenciaFinal As Date, ByVal FechaCita As Date, ByVal Estatus As String) As Boolean Implements IService1.Inserta_CitasCall

        Dim cmd As New SqlCommand("Inserta_CitaCC", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_cliente", id_cliente)
        cmd.Parameters.AddWithValue("@Pid_usuarioCC", id_usuarioCC)
        cmd.Parameters.AddWithValue("@Pid_usuarioAsesor", id_usuarioAsesor)
        cmd.Parameters.AddWithValue("@POrigen", Origen)
        cmd.Parameters.AddWithValue("@PLugar_Contacto", Lugar_Contacto)
        cmd.Parameters.AddWithValue("@PProyectoVisita", ProyectoVisita)
        cmd.Parameters.AddWithValue("@PModelo", Modelo)
        cmd.Parameters.AddWithValue("@PVigenciaInicio", VigenciaInicio)
        cmd.Parameters.AddWithValue("@PVigenciaFinal", VigenciaFinal)
        cmd.Parameters.AddWithValue("@PFechaCita", FechaCita)
        cmd.Parameters.AddWithValue("@PEstatus", Estatus)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function

    Function Verificar_VigenciaCitas(ByVal Id_Cliente As Integer) As List(Of VigenciaCitas) Implements IService1.Verificar_VigenciaCitas
        Dim DTA As New DataTable
        Dim DTB As New DataTable

        Dim cmd As New SqlCommand
        Dim Aux As VigenciaCitas
        Dim Resultado As New List(Of VigenciaCitas)

        Try
            Conexion.Close()

            Conexion.Open()
            Dim DA As New SqlDataAdapter("SELECT Id_Cita FROM CitasClientes WHERE Id_Cliente = " & Id_Cliente, Conexion)
            DA.Fill(DTA)

            For Each Row As DataRow In DTA.Rows
                cmd = New SqlCommand("Verificar_VigenciaCitas", Conexion)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@pId_Cita", Row("Id_Cita"))

                cmd.ExecuteNonQuery()
            Next

            Dim DB As New SqlDataAdapter("SELECT Id_Cita, US.id_usuario, CONCAT(US.nombre, ' ', US.apellidoPaterno, ' ', US.apellidoMaterno) AgenteAsignado,
                                                 UPPER(TU.TipoUsuario) TipoUsuario
                                          FROM CitasClientes CC
                                          INNER JOIN usuarios US ON US.id_usuario = CC.Id_Usuario
                                          INNER JOIN TipoUsuarios TU ON TU.id_tipoUsuario = US.id_tipoUsuario
                                          WHERE CC.Id_Cliente = " & Id_Cliente & " AND Status = 1", Conexion)
            DB.Fill(DTB)
            Conexion.Close()

            Aux = New VigenciaCitas
            Aux.TotalCitas = DTA.Rows.Count
            Aux.CitasVigentes = DTB.Rows.Count
            If DTB.Rows.Count = 0 Then
                Aux.UsuarioVigente = "-"
            Else
                Aux.Id_Usuario = DTB.Rows(0).Item("id_usuario")
                Aux.UsuarioVigente = DTB.Rows(0).Item("AgenteAsignado")
                Aux.TipoUsuario = DTB.Rows(0).Item("TipoUsuario")
            End If

            Resultado.Add(Aux)
        Catch ex As Exception
            Conexion.Close()

            Aux = New VigenciaCitas
            Aux.TotalCitas = -1
            Aux.CitasVigentes = -1
            Aux.UsuarioVigente = "$ERR"
            Aux.TipoUsuario = "$ERR"

            Resultado.Add(Aux)
        End Try

        Return Resultado
    End Function

    Function Verificar_VigenciaCita(ByVal Id_Cliente As Integer) As List(Of VigenciaCitas) Implements IService1.Verificar_VigenciaCita
        Dim cmd As SqlCommand
        Dim DT As New DataTable

        Dim Resultado As New List(Of VigenciaCitas)
        Dim Aux As VigenciaCitas

        Try
            Conexion.Close()

            Conexion.Open()
            Dim DA As New SqlDataAdapter("SELECT id_cita, CONCAT(nombre, ' ', apellidoPaterno, ' ', apellidoMaterno) AgenteCallCenter
                                          FROM CitasCall CC
                                          INNER JOIN usuarios US ON US.id_usuario = CC.id_usuarioCC
                                          WHERE id_cliente = " & Id_Cliente & " AND Activa = 1", Conexion)
            DA.Fill(DT)
            Conexion.Close()

            Aux = New VigenciaCitas
            Aux.TotalCitas = DT.Rows.Count
            Aux.CitasVigentes = DT.Rows.Count

            If DT.Rows.Count = 0 Then
                Aux.UsuarioVigente = "-"
            Else
                Aux.UsuarioVigente = DT.Rows(0).Item("AgenteCallCenter")
            End If

            For Each Row As DataRow In DT.Rows
                cmd = New SqlCommand("Verificar_VigenciaCita", Conexion)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@pId_Cita", Row("id_cita"))

                Conexion.Open()
                If cmd.ExecuteNonQuery() > 0 Then Aux.CitasVigentes -= 1
                Conexion.Close()
            Next

            Resultado.Add(Aux)
        Catch ex As Exception
            Conexion.Close()
            Aux = New VigenciaCitas
            Aux.TotalCitas = -1
            Aux.CitasVigentes = -1

            Resultado.Add(Aux)
        End Try

        Return Resultado
    End Function

    Function Obtener_citasCliente(ByVal idCliente As Integer) As List(Of CCitasDetallesCliente) Implements IService1.Obtener_citasCliente
        Dim Resultado As New List(Of CCitasDetallesCliente)
        Dim cmd As New SqlCommand("Obtener_citasCliente", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idCliente", idCliente)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CCitasDetallesCliente
        While reader.Read
            Aux = New CCitasDetallesCliente
            Aux.id_cita = DirectCast(reader.Item("id_cita"), Integer)
            Aux.FechaProgramada = DirectCast(reader.Item("FechaProgramada"), Date)
            Aux.HoraProgramacion = reader.Item("HoraProgramacion")
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.Programada = DirectCast(reader.Item("Programada"), String)
            Aux.AvisoCliente = DirectCast(reader.Item("AvisoCliente"), String)
            Aux.AvisoUsuario = DirectCast(reader.Item("AvisoUsuario"), String)
            Aux.realizada = DirectCast(reader.Item("realizada"), String)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Aux.HoraTermino = reader.Item("HoraTermino")
            Aux.Lugar = DirectCast(reader.Item("Lugar"), String)
            Aux.Calificacion = DirectCast(reader.Item("Calificacion"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function

    Function Obtener_observacionesCita(ByVal idCita As Integer) As CObservacionesCita Implements IService1.Obtener_observacionesCita

        Dim cmd As New SqlCommand("Obtener_observacionesCita", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idCita", idCita)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New CObservacionesCita
        While reader.Read

            Aux.Realizada = reader.Item("realizada")
            Aux.Observaciones = DirectCast(reader.Item("ObservacionUsuario"), String)

        End While
        Conexion.Close()
        Return Aux
    End Function

    Function ObservacionesCita(ByVal idCita As Integer, ByVal Observaciones As String, ByVal Realizada As Integer) As Boolean Implements IService1.ObservacionesCita

        Dim cmd As New SqlCommand("ObservacionesCita", Conexion)
        cmd.CommandType = CommandType.StoredProcedure

        cmd.Parameters.AddWithValue("@idCita", idCita)
        cmd.Parameters.AddWithValue("@obs", Observaciones)
        cmd.Parameters.AddWithValue("@Reali", Realizada)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False

    End Function

    Function VerificaSiYaSeCalificoCita(ByVal idCita As Integer) As Boolean
        Dim Resultado As Boolean = False
        Dim cmd As New SqlCommand("Obtener_CalifCita", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idCita", idCita)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As Integer = 0
        While reader.Read

            Aux = reader.Item("calificacionCliente")

            If Aux > 0 Then
                Resultado = True
            End If
        End While
        Conexion.Close()
        Return Resultado

    End Function

    Function CalificaCita(ByVal idCita As Integer, ByVal Calificacion As Integer) As Boolean Implements IService1.CalificaCita
        If VerificaSiYaSeCalificoCita(idCita) Then
        Else
            Dim cmd As New SqlCommand("califica_llamadaidCita", Conexion)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@califCliente", Calificacion)
            cmd.Parameters.AddWithValue("@idCita", idCita)
            Conexion.Close()
            Try
                Conexion.Open()
                If cmd.ExecuteNonQuery() > 0 Then
                    Conexion.Close()
                    Return True
                End If
            Catch ex As Exception
                Conexion.Close()
                Return False
            End Try
            Conexion.Close()
            Return False
        End If
        Return False
    End Function

    Function Inserta_citas(ByVal id_cliente As Integer, ByVal id_usuario As Integer, ByVal Fecha As Date, ByVal HoraProgramacion As String, ByVal Programada As String, ByVal AvisoCliente As String, ByVal AvisoUsuario As String, ByVal realizada As String, ByVal ObservacionUsuario As String, ByVal ObservacionCliente As String, ByVal HoraTermino As String, ByVal Lugar As String, ByVal ConfimacionCliente As String) As Integer Implements IService1.Inserta_citas
        Dim cmd As New SqlCommand("Inserta_Cita", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_cliente", id_cliente)
        cmd.Parameters.AddWithValue("@Pid_usuario", id_usuario)
        cmd.Parameters.AddWithValue("@PFecha", Fecha)
        cmd.Parameters.AddWithValue("@PHoraProgramacion", HoraProgramacion)
        cmd.Parameters.AddWithValue("@PProgramada", Programada)
        cmd.Parameters.AddWithValue("@PAvisoCliente", AvisoCliente)
        cmd.Parameters.AddWithValue("@PAvisoUsuario", AvisoUsuario)
        cmd.Parameters.AddWithValue("@Prealizada", realizada)
        cmd.Parameters.AddWithValue("@PObservacionUsuario", ObservacionUsuario)
        cmd.Parameters.AddWithValue("@PObservacionCliente", ObservacionCliente)
        cmd.Parameters.AddWithValue("@PHoraTermino", HoraTermino)
        cmd.Parameters.AddWithValue("@PLugar", Lugar)
        cmd.Parameters.AddWithValue("@PConfimacionCliente", ConfimacionCliente)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As Integer = 0
        While reader.Read

            Aux = reader.Item(0)


        End While
        Conexion.Close()
        Return Aux
    End Function

    Function Elimina_citas(ByVal id_cita As Integer) As Boolean Implements IService1.Elimina_citas

        Dim cmd As New SqlCommand("Elimina_citas", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_cita", id_cita)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function

    Function Actualiza_citas(ByVal id_cita As Integer, ByVal id_cliente As Integer, ByVal id_usuario As Integer, ByVal Fecha As Date, ByVal fechaCreacion As Date, ByVal HoraProgramacion As String, ByVal Programada As String, ByVal AvisoCliente As String, ByVal AvisoUsuario As String, ByVal realizada As String, ByVal ObservacionUsuario As String, ByVal ObservacionCliente As String, ByVal HoraTermino As String, ByVal Lugar As String, ByVal ConfimacionCliente As String) As Boolean Implements IService1.Actualiza_citas

        Dim cmd As New SqlCommand("Actualiza_citas", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_cita", id_cita)
        cmd.Parameters.AddWithValue("@Pid_cliente", id_cliente)
        cmd.Parameters.AddWithValue("@Pid_usuario", id_usuario)
        cmd.Parameters.AddWithValue("@PFecha", Fecha)
        cmd.Parameters.AddWithValue("@PfechaCreacion", fechaCreacion)
        cmd.Parameters.AddWithValue("@PHoraProgramacion", HoraProgramacion)
        cmd.Parameters.AddWithValue("@PProgramada", Programada)
        cmd.Parameters.AddWithValue("@PAvisoCliente", AvisoCliente)
        cmd.Parameters.AddWithValue("@PAvisoUsuario", AvisoUsuario)
        cmd.Parameters.AddWithValue("@Prealizada", realizada)
        cmd.Parameters.AddWithValue("@PObservacionUsuario", ObservacionUsuario)
        cmd.Parameters.AddWithValue("@PObservacionCliente", ObservacionCliente)
        cmd.Parameters.AddWithValue("@PHoraTermino", HoraTermino)
        cmd.Parameters.AddWithValue("@PLugar", Lugar)
        cmd.Parameters.AddWithValue("@PConfimacionCliente", ConfimacionCliente)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function

    Function Obtener_citas_id_usuario(ByVal id_usuario As Integer) As List(Of CCitas) Implements IService1.Obtener_citas_id_usuario
        Dim Resultado As New List(Of CCitas)
        Dim cmd As New SqlCommand("Obtener_citas_detalles_idCliente", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CCitas
        While reader.Read
            Aux = New CCitas
            Aux.id_cita = DirectCast(reader.Item("id_cita"), Integer)
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.id_usuario = DirectCast(reader.Item("id_usuario"), Integer)
            Aux.Fecha = DirectCast(reader.Item("Fecha"), Date)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.HoraProgramacion = reader.Item("HoraProgramacion")
            Aux.Programada = reader.Item("Programada")
            Aux.AvisoCliente = reader.Item("AvisoCliente")
            Aux.AvisoUsuario = reader.Item("AvisoUsuario")
            Aux.realizada = reader.Item("realizada")
            Aux.ObservacionUsuario = DirectCast(reader.Item("ObservacionUsuario"), String)
            Aux.ObservacionCliente = DirectCast(reader.Item("ObservacionCliente"), String)
            Aux.HoraTermino = reader.Item("HoraTermino")
            Aux.Lugar = DirectCast(reader.Item("Lugar"), String)
            Aux.ConfimacionCliente = reader.Item("ConfimacionCliente")
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function

    Function Obtener_citas_detalles_idusuario(ByVal id_usuario As Integer) As List(Of CDetallesCitaUsuario) Implements IService1.Obtener_citas_detalles_idusuario
        Dim Resultado As New List(Of CDetallesCitaUsuario)
        Dim cmd As New SqlCommand("Obtener_citas_detalles_idusuario", Conexion)
        cmd.Parameters.AddWithValue("@Pid_usuario", id_usuario)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CDetallesCitaUsuario
        While reader.Read
            Aux = New CDetallesCitaUsuario
            Aux.id_cita = DirectCast(reader.Item("id_cita"), Integer)
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.Nombre = DirectCast(reader.Item("Nombre"), String)
            Aux.ApellidoPaterno = DirectCast(reader.Item("ApellidoPaterno"), String)
            Aux.ApellidoMaterno = DirectCast(reader.Item("ApellidoMaterno"), String)
            Aux.Fecha = DirectCast(reader.Item("Fecha"), Date)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.HoraProgramacion = reader.Item("HoraProgramacion")
            Aux.Programada = reader.Item("Programada")
            Aux.AvisoCliente = reader.Item("AvisoCliente")
            Aux.AvisoUsuario = reader.Item("AvisoUsuario")
            Aux.realizada = reader.Item("realizada")
            Aux.ObservacionUsuario = DirectCast(reader.Item("ObservacionUsuario"), String)
            Aux.ObservacionCliente = DirectCast(reader.Item("ObservacionCliente"), String)
            Aux.HoraTermino = reader.Item("HoraTermino")
            Aux.Lugar = DirectCast(reader.Item("Lugar"), String)
            Aux.ConfimacionCliente = reader.Item("ConfimacionCliente")
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function

    Function Obtener_citas_detalles_idCliente(ByVal id_cliente As Integer) As List(Of CDetallesCitaUsuario) Implements IService1.Obtener_citas_detalles_idCliente
        Dim Resultado As New List(Of CDetallesCitaUsuario)
        Dim cmd As New SqlCommand("Obtener_citas_detalles_idCliente", Conexion)
        cmd.Parameters.AddWithValue("@PidCliente", id_cliente)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CDetallesCitaUsuario
        While reader.Read
            Aux = New CDetallesCitaUsuario
            Aux.id_cita = DirectCast(reader.Item("id_cita"), Integer)
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.Nombre = DirectCast(reader.Item("Nombre"), String)
            Aux.ApellidoPaterno = DirectCast(reader.Item("ApellidoPaterno"), String)
            Aux.ApellidoMaterno = DirectCast(reader.Item("ApellidoMaterno"), String)
            Aux.Fecha = DirectCast(reader.Item("Fecha"), Date)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.HoraProgramacion = reader.Item("HoraProgramacion")
            Aux.Programada = reader.Item("Programada")
            Aux.AvisoCliente = reader.Item("AvisoCliente")
            Aux.AvisoUsuario = reader.Item("AvisoUsuario")
            Aux.realizada = reader.Item("realizada")
            Aux.ObservacionUsuario = DirectCast(reader.Item("ObservacionUsuario"), String)
            Aux.ObservacionCliente = DirectCast(reader.Item("ObservacionCliente"), String)
            Aux.HoraTermino = reader.Item("HoraTermino")
            Aux.Lugar = DirectCast(reader.Item("Lugar"), String)
            Aux.ConfimacionCliente = reader.Item("ConfimacionCliente")
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "Visitas"
    Function Insertar_VisitasClientes(ByVal IdCita As Integer, ByVal IdCliente As Integer, ByVal IdUsuario As Integer, ByVal IdUsuarioAsignado As Integer, ByVal IdUsuarioVisita As Integer,
                                      ByVal IdCampana As Integer, ByVal IdImpedimento As Integer, ByVal TipoCredito As String, ByVal Monto As Double, ByVal Ranking As String,
                                      ByVal Origen As String, ByVal Proyecto As String, ByVal Modelo As Integer, ByVal TipoCampana As String, ByVal VigenciaIncial As Date,
                                      ByVal VigenciaFinal As Date, ByVal FechaVisita As Date, ByVal Status As Integer) As Boolean Implements IService1.Insertar_VisitasClientes

        Dim cmd As New SqlCommand("Insertar_VisitasClientes", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@pIdCita", IdCita)
        cmd.Parameters.AddWithValue("@pIdCliente", IdCliente)
        cmd.Parameters.AddWithValue("@pIdUsuario", IdUsuario)
        cmd.Parameters.AddWithValue("@pIdUsuarioAsignado", IdUsuarioAsignado)
        cmd.Parameters.AddWithValue("@pIdUsuarioVisita", IdUsuarioVisita)
        cmd.Parameters.AddWithValue("@pIdCampana", IdCampana)
        cmd.Parameters.AddWithValue("@pIdImpedimento", IdImpedimento)
        cmd.Parameters.AddWithValue("@pTipoCredito", TipoCredito)
        cmd.Parameters.AddWithValue("@pMonto", Monto)
        cmd.Parameters.AddWithValue("@pRanking", Ranking)
        cmd.Parameters.AddWithValue("@pOrigen", Origen)
        cmd.Parameters.AddWithValue("@pProyecto", Proyecto)
        cmd.Parameters.AddWithValue("@pModelo", Modelo)
        cmd.Parameters.AddWithValue("@pTipoCampana", TipoCampana)
        cmd.Parameters.AddWithValue("@pVigenciaInicial", VigenciaIncial)
        cmd.Parameters.AddWithValue("@pVigenciaFinal", VigenciaFinal)
        cmd.Parameters.AddWithValue("@pFechaVisita", FechaVisita)
        cmd.Parameters.AddWithValue("@pStatus", Status)

        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try

        Conexion.Close()
        Return False
    End Function

    Function Obtener_TipoVisita() As DataSet Implements IService1.Obtener_TipoVisita
        Dim Query As String = "EXEC [dbo].[spObtener_TipoVisita]"
        Return GE_SQL.SQlGetDataset(Query)
    End Function
#End Region
#Region "Configuraciones"
    Function Actualiza_configuraciones(ByVal id_configuracion As Integer, ByVal diasDeGracias As Integer, ByVal emailSistema As String, ByVal contraseñaEmail As String, ByVal smtpServer As String, ByVal puertoEmail As Integer, ByVal SSL As String, ByVal EnviarEmails As String) As Boolean Implements IService1.Actualiza_configuraciones

        Dim cmd As New SqlCommand("Actualiza_configuraciones", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_configuracion", id_configuracion)
        cmd.Parameters.AddWithValue("@PdiasDeGracias", diasDeGracias)
        cmd.Parameters.AddWithValue("@PemailSistema", emailSistema)
        cmd.Parameters.AddWithValue("@PcontraseñaEmail", contraseñaEmail)
        cmd.Parameters.AddWithValue("@PsmtpServer", smtpServer)
        cmd.Parameters.AddWithValue("@PpuertoEmail", puertoEmail)
        cmd.Parameters.AddWithValue("@PSSL", SSL)
        'cmd.Parameters.AddWithValue("@PusuarioAdmin", usuarioAdmin)
        'cmd.Parameters.AddWithValue("@Pcontraseña", contraseña)
        cmd.Parameters.AddWithValue("@PEnviarEmails", EnviarEmails)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_configuraciones() As CConfiguraciones Implements IService1.Obtener_configuraciones

        Dim cmd As New SqlCommand("Obtener_configuraciones", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New CConfiguraciones
        While reader.Read

            Aux.id_configuracion = DirectCast(reader.Item("id_configuracion"), Integer)
            Aux.diasDeGracias = DirectCast(reader.Item("diasDeGracias"), Integer)
            Aux.emailSistema = DirectCast(reader.Item("emailSistema"), String)
            Aux.contraseñaEmail = DirectCast(reader.Item("contraseñaEmail"), String)
            Aux.smtpServer = DirectCast(reader.Item("smtpServer"), String)
            Aux.puertoEmail = DirectCast(reader.Item("puertoEmail"), Integer)
            Aux.SSL = reader.Item("SSL")
            Aux.EnviarEmails = reader.Item("EnviarEmails")

        End While
        Conexion.Close()
        Return Aux
    End Function

#End Region
#Region "Contacto Empresa"
    Function Inserta_ContactoEmpresa(ByVal id_empresa As Integer, ByVal Nombre As String, ByVal ApellidoPaterno As String, ByVal ApellidoMaterno As String, ByVal Email As String, ByVal Observaciones As String, ByVal fotografia As String) As Boolean Implements IService1.Inserta_ContactoEmpresa

        Dim cmd As New SqlCommand("Inserta_contacto_Empresa", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_empresa", id_empresa)
        cmd.Parameters.AddWithValue("@PNombre", Nombre)
        cmd.Parameters.AddWithValue("@PApellidoPaterno", ApellidoPaterno)
        cmd.Parameters.AddWithValue("@PApellidoMaterno", ApellidoMaterno)
        cmd.Parameters.AddWithValue("@PEmail", Email)
        cmd.Parameters.AddWithValue("@PObservaciones", Observaciones)
        cmd.Parameters.AddWithValue("@Pfotografia", fotografia)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_ContactoEmpresa(ByVal id_contactoEmpresa As Integer) As Boolean Implements IService1.Elimina_ContactoEmpresa

        Dim cmd As New SqlCommand("Elimina_ContactoEmpresa", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_contactoEmpresa", id_contactoEmpresa)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_ContactoEmpresa(ByVal id_contactoEmpresa As Integer, ByVal id_empresa As Integer, ByVal Nombre As String, ByVal ApellidoPaterno As String, ByVal ApellidoMaterno As String, ByVal Email As String, ByVal Observaciones As String, ByVal fotografia As String) As Boolean Implements IService1.Actualiza_ContactoEmpresa

        Dim cmd As New SqlCommand("Actualiza_ContactoEmpresa", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_contactoEmpresa", id_contactoEmpresa)
        cmd.Parameters.AddWithValue("@Pid_empresa", id_empresa)
        cmd.Parameters.AddWithValue("@PNombre", Nombre)
        cmd.Parameters.AddWithValue("@PApellidoPaterno", ApellidoPaterno)
        cmd.Parameters.AddWithValue("@PApellidoMaterno", ApellidoMaterno)
        cmd.Parameters.AddWithValue("@PEmail", Email)
        cmd.Parameters.AddWithValue("@PObservaciones", Observaciones)
        cmd.Parameters.AddWithValue("@Pfotografia", fotografia)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_ContactoEmpresa(ByVal id_empresa As Integer) As List(Of CContactoEmpresa) Implements IService1.Obtener_ContactoEmpresa
        Dim Resultado As New List(Of CContactoEmpresa)
        Dim cmd As New SqlCommand("Obtener_contactos_idempresa", Conexion)
        cmd.Parameters.AddWithValue("@PidEmpresa", id_empresa)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CContactoEmpresa
        While reader.Read
            Aux = New CContactoEmpresa
            Aux.id_contactoEmpresa = DirectCast(reader.Item("id_contactoEmpresa"), Integer)

            Aux.Nombre = DirectCast(reader.Item("Nombre"), String)
            Aux.ApellidoPaterno = DirectCast(reader.Item("ApellidoPaterno"), String)
            Aux.ApellidoMaterno = DirectCast(reader.Item("ApellidoMaterno"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Aux.fotografia = DirectCast(reader.Item("fotografia"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_detalles_empresa_idEmpresa(ByVal id_contactoEmpresa As Integer) As CContactoEmpresa Implements IService1.Obtener_detalles_empresa_idEmpresa

        Dim cmd As New SqlCommand("Obtener_detalles_contactoEmpresa", Conexion)
        cmd.Parameters.AddWithValue("@idContactoEmpresa", id_contactoEmpresa)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New CContactoEmpresa
        While reader.Read

            Aux.id_contactoEmpresa = DirectCast(reader.Item("id_contactoEmpresa"), Integer)
            Aux.Nombre = DirectCast(reader.Item("Nombre"), String)
            Aux.ApellidoPaterno = DirectCast(reader.Item("ApellidoPaterno"), String)
            Aux.ApellidoMaterno = DirectCast(reader.Item("ApellidoMaterno"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Aux.fotografia = DirectCast(reader.Item("fotografia"), String)

        End While
        Conexion.Close()
        Return Aux
    End Function
#End Region
#Region "Empresas"
    Function Obtener_empresasComboBusqueda(ByVal Query As String) As List(Of CComboEmpresas) Implements IService1.Obtener_empresasComboBusqueda
        Dim Resultado As New List(Of CComboEmpresas)
        Dim cmd As New SqlCommand("Obtener_empresasComboBusqueda", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@query", Query)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CComboEmpresas
        While reader.Read
            Aux = New CComboEmpresas
            Aux.id_empresa = DirectCast(reader.Item("id_empresa"), Integer)
            Aux.Empresa = DirectCast(reader.Item("Empresa"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_combo_empresas() As List(Of CComboEmpresas) Implements IService1.Obtener_combo_empresas
        Dim Resultado As New List(Of CComboEmpresas)
        Dim cmd As New SqlCommand("Obtener_empresasCombo", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CComboEmpresas
        While reader.Read
            Aux = New CComboEmpresas
            Aux.id_empresa = DirectCast(reader.Item("id_empresa"), Integer)
            Aux.Empresa = DirectCast(reader.Item("Empresa"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Inserta_empresas(ByVal Empresa As String, ByVal Razon_Social As String, ByVal Direccion As String, ByVal PaginaWEb As String, ByVal Horario As String, ByVal id_rubro As Integer, ByVal id_ciudad As Integer, ByVal email As String, ByVal Observaciones As String, ByVal logotipo As String) As Boolean Implements IService1.Inserta_empresas

        Dim cmd As New SqlCommand("Inserta_empresas", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@PEmpresa", Empresa)
        cmd.Parameters.AddWithValue("@PRazon_Social", Razon_Social)
        cmd.Parameters.AddWithValue("@PDireccion", Direccion)
        cmd.Parameters.AddWithValue("@PPaginaWEb", PaginaWEb)
        'cmd.Parameters.AddWithValue("@PfechaCreacion", fechaCreacion)
        cmd.Parameters.AddWithValue("@PHorario", Horario)
        cmd.Parameters.AddWithValue("@Pid_rubro", id_rubro)
        cmd.Parameters.AddWithValue("@Pid_ciudad", id_ciudad)
        cmd.Parameters.AddWithValue("@Pemail", email)
        cmd.Parameters.AddWithValue("@PObservaciones", Observaciones)
        cmd.Parameters.AddWithValue("@Plogotipo", logotipo)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_empresas(ByVal id_empresa As Integer) As Boolean Implements IService1.Elimina_empresas

        Dim cmd As New SqlCommand("Elimina_empresas", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_empresa", id_empresa)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_empresas(ByVal id_empresa As Integer, ByVal Empresa As String, ByVal Razon_Social As String, ByVal Direccion As String, ByVal PaginaWEb As String, ByVal fechaCreacion As Date, ByVal Horario As String, ByVal id_rubro As Integer, ByVal id_ciudad As Integer, ByVal email As String, ByVal Observaciones As String, ByVal logotipo As String) As Boolean Implements IService1.Actualiza_empresas

        Dim cmd As New SqlCommand("Actualiza_empresas", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_empresa", id_empresa)
        cmd.Parameters.AddWithValue("@PEmpresa", Empresa)
        cmd.Parameters.AddWithValue("@PRazon_Social", Razon_Social)
        cmd.Parameters.AddWithValue("@PDireccion", Direccion)
        cmd.Parameters.AddWithValue("@PPaginaWEb", PaginaWEb)
        cmd.Parameters.AddWithValue("@PfechaCreacion", fechaCreacion)
        cmd.Parameters.AddWithValue("@PHorario", Horario)
        cmd.Parameters.AddWithValue("@Pid_rubro", id_rubro)
        cmd.Parameters.AddWithValue("@Pid_ciudad", id_ciudad)
        cmd.Parameters.AddWithValue("@Pemail", email)
        cmd.Parameters.AddWithValue("@PObservaciones", Observaciones)
        cmd.Parameters.AddWithValue("@Plogotipo", logotipo)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_empresas() As List(Of CEmpresas) Implements IService1.Obtener_empresas
        Dim Resultado As New List(Of CEmpresas)
        Dim cmd As New SqlCommand("Obtener_empresas", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CEmpresas
        While reader.Read
            Aux = New CEmpresas
            Aux.id_empresa = DirectCast(reader.Item("id_empresa"), Integer)
            Aux.Empresa = DirectCast(reader.Item("Empresa"), String)
            Aux.Razon_Social = DirectCast(reader.Item("Razon_Social"), String)
            Aux.Direccion = DirectCast(reader.Item("Direccion"), String)
            Aux.PaginaWEb = DirectCast(reader.Item("PaginaWEb"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.Horario = DirectCast(reader.Item("Horario"), String)
            Aux.id_rubro = DirectCast(reader.Item("id_rubro"), Integer)
            Aux.id_ciudad = DirectCast(reader.Item("id_ciudad"), Integer)
            Aux.email = DirectCast(reader.Item("email"), String)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Aux.logotipo = DirectCast(reader.Item("logotipo"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_detalles_empresas() As List(Of CEmpresasDetalles) Implements IService1.Obtener_detalles_empresas
        Dim Resultado As New List(Of CEmpresasDetalles)
        Dim cmd As New SqlCommand("Obtener_empresas_detalles", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CEmpresasDetalles
        While reader.Read
            Aux = New CEmpresasDetalles
            Aux.id_empresa = DirectCast(reader.Item("id_empresa"), Integer)
            Aux.Empresa = DirectCast(reader.Item("Empresa"), String)
            Aux.Razon_Social = DirectCast(reader.Item("Razon_Social"), String)
            Aux.Direccion = DirectCast(reader.Item("Direccion"), String)
            Aux.PaginaWEb = DirectCast(reader.Item("PaginaWEb"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.Horario = DirectCast(reader.Item("Horario"), String)
            Aux.rubro = DirectCast(reader.Item("rubro"), String)
            Aux.Ciudad = DirectCast(reader.Item("Ciudad"), String)
            Aux.Estado = DirectCast(reader.Item("Estado"), String)
            Aux.email = DirectCast(reader.Item("email"), String)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Aux.logotipo = DirectCast(reader.Item("logotipo"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_detalles_empresas_idEmpresa(ByVal id_empresa As Integer) As CEmpresasDetalles Implements IService1.Obtener_detalles_empresas_idEmpresa

        Dim cmd As New SqlCommand("Obtener_detalles_Empresa", Conexion)
        cmd.Parameters.AddWithValue("@idEmpresa", id_empresa)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New CEmpresasDetalles
        While reader.Read

            Aux.id_empresa = DirectCast(reader.Item("id_empresa"), Integer)
            Aux.Empresa = DirectCast(reader.Item("Empresa"), String)
            Aux.Razon_Social = DirectCast(reader.Item("Razon_Social"), String)
            Aux.Direccion = DirectCast(reader.Item("Direccion"), String)
            Aux.PaginaWEb = DirectCast(reader.Item("PaginaWEb"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.Horario = DirectCast(reader.Item("Horario"), String)
            Aux.rubro = DirectCast(reader.Item("rubro"), String)
            Aux.Ciudad = DirectCast(reader.Item("Ciudad"), String)
            Aux.Estado = DirectCast(reader.Item("Estado"), String)
            Aux.email = DirectCast(reader.Item("email"), String)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Aux.logotipo = DirectCast(reader.Item("logotipo"), String)

        End While
        Conexion.Close()
        Return Aux
    End Function
    Function Obtener_detallesEmpresa_idCliente(ByVal id_cliente As Integer) As CEmpresasDetalles Implements IService1.Obtener_detallesEmpresa_idCliente

        Dim cmd As New SqlCommand("Obtener_detallesEmpresa_idCliente", Conexion)
        cmd.Parameters.AddWithValue("@idCliente", id_cliente)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New CEmpresasDetalles
        While reader.Read

            Aux.id_empresa = DirectCast(reader.Item("id_empresa"), Integer)
            Aux.Empresa = DirectCast(reader.Item("Empresa"), String)
            Aux.Razon_Social = DirectCast(reader.Item("Razon_Social"), String)
            Aux.Direccion = DirectCast(reader.Item("Direccion"), String)
            Aux.PaginaWEb = DirectCast(reader.Item("PaginaWEb"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.Horario = DirectCast(reader.Item("Horario"), String)
            Aux.rubro = DirectCast(reader.Item("rubro"), String)
            Aux.Ciudad = DirectCast(reader.Item("Ciudad"), String)
            Aux.Estado = DirectCast(reader.Item("Estado"), String)
            Aux.email = DirectCast(reader.Item("email"), String)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Aux.logotipo = DirectCast(reader.Item("logotipo"), String)

        End While
        Conexion.Close()
        Return Aux
    End Function
#End Region
#Region "Operaciones"
    Function Obtener_operacionesIdCliente(ByVal idCliente As Integer) As List(Of COperacionesCliente) Implements IService1.Obtener_operacionesIdCliente
        Dim Resultado As New List(Of COperacionesCliente)
        Dim cmd As New SqlCommand("Obtener_operacionesIdCliente", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idCLiente", idCliente)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As COperacionesCliente
        While reader.Read
            Aux = New COperacionesCliente
            Aux.id_etapa = DirectCast(reader.Item("id_etapa"), Integer)
            Aux.Etapa = DirectCast(reader.Item("Etapa"), String)
            Aux.usuario = DirectCast(reader.Item("usuario"), String)
            Aux.FechaInicio = DirectCast(reader.Item("FechaInicio"), Date)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Aux.Producto = DirectCast(reader.Item("Producto"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Inserta_operaciones(ByVal id_cliente As Integer, ByVal id_usuario As Integer, ByVal id_etapa As Integer, ByVal FechaInicio As Date, ByVal FechaFinal As Date, ByVal Observaciones As String) As Boolean Implements IService1.Inserta_operaciones

        Dim cmd As New SqlCommand("Inserta_operacion", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_cliente", id_cliente)
        cmd.Parameters.AddWithValue("@Pid_usuario", id_usuario)
        cmd.Parameters.AddWithValue("@Pid_etapa", id_etapa)
        cmd.Parameters.AddWithValue("@PFechaInicio", FechaInicio)
        cmd.Parameters.AddWithValue("@PFechaFinal", FechaFinal)
        cmd.Parameters.AddWithValue("@PObservaciones", Observaciones)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_operaciones(ByVal id_operacion As Integer) As Boolean Implements IService1.Elimina_operaciones

        Dim cmd As New SqlCommand("Elimina_operaciones", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_operacion", id_operacion)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_operaciones(ByVal id_operacion As Integer, ByVal id_cliente As Integer, ByVal id_usuario As Integer, ByVal id_etapa As Integer, ByVal FechaInicio As Date, ByVal FechaFinal As Date, ByVal Observaciones As String) As Boolean Implements IService1.Actualiza_operaciones

        Dim cmd As New SqlCommand("Actualiza_operaciones", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_operacion", id_operacion)
        cmd.Parameters.AddWithValue("@Pid_cliente", id_cliente)
        cmd.Parameters.AddWithValue("@Pid_usuario", id_usuario)
        cmd.Parameters.AddWithValue("@Pid_etapa", id_etapa)
        cmd.Parameters.AddWithValue("@PFechaInicio", FechaInicio)
        cmd.Parameters.AddWithValue("@PFechaFinal", FechaFinal)
        cmd.Parameters.AddWithValue("@PObservaciones", Observaciones)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_operaciones(ByVal id_cliente As Integer) As List(Of COperaciones) Implements IService1.Obtener_operaciones
        Dim Resultado As New List(Of COperaciones)
        Dim cmd As New SqlCommand("Obtener_operaciones_idcliente", Conexion)
        cmd.Parameters.AddWithValue("@IdCliente", id_cliente)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As COperaciones
        While reader.Read
            Aux = New COperaciones
            Aux.id_operacion = DirectCast(reader.Item("id_operacion"), Integer)
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.id_usuario = DirectCast(reader.Item("id_usuario"), Integer)
            Aux.id_etapa = DirectCast(reader.Item("id_etapa"), Integer)
            Aux.FechaInicio = DirectCast(reader.Item("FechaInicio"), Date)
            Aux.FechaFinal = DirectCast(reader.Item("FechaFinal"), Date)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "Etapas Cliente"
    Function Inserta_etapasCliente(ByVal nEtapa As Integer, ByVal Descripcion As String) As Boolean Implements IService1.Inserta_etapasCliente

        Dim cmd As New SqlCommand("Inserta_etapasCliente", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@PnEtapa", nEtapa)
        cmd.Parameters.AddWithValue("@PDescripcion", Descripcion)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_etapasCliente() As List(Of CEtapasCliente) Implements IService1.Obtener_etapasCliente
        Dim Resultado As New List(Of CEtapasCliente)
        Dim cmd As New SqlCommand("Obtener_etapasCliente", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CEtapasCliente
        While reader.Read
            Aux = New CEtapasCliente
            Aux.id_etapa = DirectCast(reader.Item("id_etapa"), Integer)
            Aux.nEtapa = DirectCast(reader.Item("nEtapa"), Integer)
            Aux.Descripcion = DirectCast(reader.Item("Descripcion"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Elimina_etapasCliente(ByVal id_etapa As Integer) As Boolean Implements IService1.Elimina_etapasCliente

        Dim cmd As New SqlCommand("Elimina_etapasCliente", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_etapa", id_etapa)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_etapasCliente(ByVal id_etapa As Integer, ByVal nEtapa As Integer, ByVal Descripcion As String) As Boolean Implements IService1.Actualiza_etapasCliente

        Dim cmd As New SqlCommand("Actualiza_etapasCliente", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_etapa", id_etapa)
        cmd.Parameters.AddWithValue("@PnEtapa", nEtapa)
        cmd.Parameters.AddWithValue("@PDescripcion", Descripcion)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
#End Region
#Region "llamadas"
    Function VerificaSiYaSeCalifico(ByVal id_llamada As Integer) As Boolean
        Dim Resultado As Boolean = False
        Dim cmd As New SqlCommand("Obtener_CalifLlamada", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idLlam", id_llamada)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As Integer = 0
        While reader.Read

            Aux = reader.Item("calificacionCliente")

            If Aux > 0 Then
                Resultado = True
            End If
        End While
        Conexion.Close()
        Return Resultado

    End Function
    Function CalificaLlamada(ByVal id_llamada As Integer, ByVal Calificacion As Integer) As Boolean Implements IService1.CalificaLlamada
        If VerificaSiYaSeCalifico(id_llamada) Then
        Else
            Dim cmd As New SqlCommand("califica_llamadaidLlamada", Conexion)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@califCliente", Calificacion)
            cmd.Parameters.AddWithValue("@idLlam", id_llamada)
            Conexion.Close()
            Try
                Conexion.Open()
                If cmd.ExecuteNonQuery() > 0 Then
                    Conexion.Close()
                    Return True
                End If
            Catch ex As Exception
                Conexion.Close()
                Return False
            End Try
            Conexion.Close()
            Return False
        End If
        Return False
    End Function
    Function Obtener_llamadasPendientesHoyUsuario(ByVal id_usuario As Integer) As List(Of CLlamadasPendientesHoyUsuario) Implements IService1.Obtener_llamadasPendientesHoyUsuario
        Dim Resultado As New List(Of CLlamadasPendientesHoyUsuario)
        Dim cmd As New SqlCommand("Obtener_llamadasPendientesHoyUsuario", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idUsuario", id_usuario)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CLlamadasPendientesHoyUsuario
        While reader.Read
            Aux = New CLlamadasPendientesHoyUsuario
            Aux.Nombre = DirectCast(reader.Item("Nombre"), String)
            Aux.ApellidoPaterno = DirectCast(reader.Item("ApellidoPaterno"), String)
            Aux.ApellidoMaterno = DirectCast(reader.Item("ApellidoMaterno"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Aux.Producto = DirectCast(reader.Item("Producto"), String)
            Aux.Fecha = DirectCast(reader.Item("Fecha"), Date)
            Aux.HORA = reader.Item("HORA")
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_llamadasFechaUsuario(ByVal idUsuario As Integer, ByVal FechaInicio As Date, ByVal FechaFinal As Date) As List(Of CLlamadasFechas) Implements IService1.Obtener_llamadasFechaUsuario
        Dim Resultado As New List(Of CLlamadasFechas)
        Dim cmd As New SqlCommand("Obtener_llamadasFechaUsuario", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idUsuario", idUsuario)
        cmd.Parameters.AddWithValue("@fechainicio", FechaInicio)
        cmd.Parameters.AddWithValue("@Fechafinal", FechaFinal)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CLlamadasFechas
        While reader.Read
            Aux = New CLlamadasFechas
            Aux.Cliente = DirectCast(reader.Item("Cliente"), String)
            Aux.id_llamada = DirectCast(reader.Item("id_llamada"), Integer)
            Aux.Fecha = DirectCast(reader.Item("Fecha"), Date)
            Aux.HoraProgramacion = reader.Item("HoraProgramacion")
            Aux.realizada = reader.Item("realizada")
            Aux.ObservacionUsuario = DirectCast(reader.Item("ObservacionUsuario"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_llamadasCliente(ByVal idCliente As Integer) As List(Of CLlamadasCliente) Implements IService1.Obtener_llamadasCliente
        Dim Resultado As New List(Of CLlamadasCliente)
        Dim cmd As New SqlCommand("Obtener_llamadasCliente", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idCliente", idCliente)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CLlamadasCliente
        While reader.Read
            Aux = New CLlamadasCliente
            Aux.id_llamada = DirectCast(reader.Item("id_llamada"), Integer)
            Aux.usuario = DirectCast(reader.Item("usuario"), String)
            Aux.Fecha = DirectCast(reader.Item("Fecha"), Date)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.HoraProgramacion = reader.Item("HoraProgramacion")
            Aux.Programada = DirectCast(reader.Item("Programada"), String)
            Aux.AvisoCliente = DirectCast(reader.Item("AvisoCliente"), String)
            Aux.AvisoUsuario = DirectCast(reader.Item("AvisoUsuario"), String)
            Aux.realizada = DirectCast(reader.Item("realizada"), String)
            Aux.ObservacionUsuario = DirectCast(reader.Item("ObservacionUsuario"), String)
            Aux.ObservacionCliente = DirectCast(reader.Item("ObservacionCliente"), String)
            Aux.Calificacion = DirectCast(reader.Item("Calificacion"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Inserta_llamadas(ByVal id_cliente As Integer, ByVal id_usuario As Integer, ByVal Fecha As Date, ByVal HoraProgramacion As Date, ByVal Programada As String, ByVal AvisoCliente As String, ByVal AvisoUsuario As String, ByVal realizada As String, ByVal ObservacionUsuario As String, ByVal ObservacionCliente As String) As Integer Implements IService1.Inserta_llamadas

        Dim cmd As New SqlCommand("Inserta_llamadas", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_cliente", id_cliente)
        cmd.Parameters.AddWithValue("@Pid_usuario", id_usuario)
        cmd.Parameters.AddWithValue("@PFecha", Fecha)
        cmd.Parameters.AddWithValue("@PHoraProgramacion", HoraProgramacion)
        cmd.Parameters.AddWithValue("@PProgramada", Programada)
        cmd.Parameters.AddWithValue("@PAvisoCliente", AvisoCliente)
        cmd.Parameters.AddWithValue("@PAvisoUsuario", AvisoUsuario)
        cmd.Parameters.AddWithValue("@Prealizada", realizada)
        cmd.Parameters.AddWithValue("@PObservacionUsuario", ObservacionUsuario)
        cmd.Parameters.AddWithValue("@PObservacionCliente", ObservacionCliente)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As Integer = 0
        While reader.Read

            Aux = reader.Item(0)

        End While
        Conexion.Close()
        Return Aux
    End Function
    Function Cambia_realizadaLlamada(ByVal id_llamada As Integer) As Boolean Implements IService1.Cambia_realizadaLlamada

        Dim cmd As New SqlCommand("Cambia_realizadaLlamada", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idLlamda", id_llamada)

        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Try
                    Enviar_CorreoLlamadaCliente(id_llamada)
                Catch ex As Exception

                End Try

                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_llamadas(ByVal id_llamada As Integer, ByVal id_cliente As Integer, ByVal id_usuario As Integer, ByVal Fecha As Date, ByVal fechaCreacion As Date, ByVal HoraProgramacion As String, ByVal Programada As String, ByVal AvisoCliente As String, ByVal AvisoUsuario As String, ByVal realizada As String, ByVal ObservacionUsuario As String, ByVal ObservacionCliente As String) As Boolean Implements IService1.Actualiza_llamadas

        Dim cmd As New SqlCommand("Actualiza_llamadas", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_llamada", id_llamada)
        cmd.Parameters.AddWithValue("@Pid_cliente", id_cliente)
        cmd.Parameters.AddWithValue("@Pid_usuario", id_usuario)
        cmd.Parameters.AddWithValue("@PFecha", Fecha)
        cmd.Parameters.AddWithValue("@PfechaCreacion", fechaCreacion)
        cmd.Parameters.AddWithValue("@PHoraProgramacion", HoraProgramacion)
        cmd.Parameters.AddWithValue("@PProgramada", Programada)
        cmd.Parameters.AddWithValue("@PAvisoCliente", AvisoCliente)
        cmd.Parameters.AddWithValue("@PAvisoUsuario", AvisoUsuario)
        cmd.Parameters.AddWithValue("@Prealizada", realizada)
        cmd.Parameters.AddWithValue("@PObservacionUsuario", ObservacionUsuario)
        cmd.Parameters.AddWithValue("@PObservacionCliente", ObservacionCliente)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_llamadas(ByVal id_llamada As Integer) As Boolean Implements IService1.Elimina_llamadas

        Dim cmd As New SqlCommand("Elimina_llamadas", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_llamada", id_llamada)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_llamadas_usuario(ByVal id_usuario As Integer) As List(Of CLlamadas) Implements IService1.Obtener_llamadas_usuario
        Dim Resultado As New List(Of CLlamadas)
        Dim cmd As New SqlCommand("Obtener_llamadas_usuario", Conexion)
        cmd.Parameters.AddWithValue("@idUsuario", id_usuario)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CLlamadas
        While reader.Read
            Aux = New CLlamadas
            Aux.id_llamada = DirectCast(reader.Item("id_llamada"), Integer)
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.id_usuario = DirectCast(reader.Item("id_usuario"), Integer)
            Aux.Fecha = DirectCast(reader.Item("Fecha"), Date)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.HoraProgramacion = reader.Item("HoraProgramacion")
            Aux.Programada = reader.Item("Programada")
            Aux.AvisoCliente = reader.Item("AvisoCliente")
            Aux.AvisoUsuario = reader.Item("AvisoUsuario")
            Aux.realizada = reader.Item("realizada")
            Aux.ObservacionUsuario = DirectCast(reader.Item("ObservacionUsuario"), String)
            Aux.ObservacionCliente = DirectCast(reader.Item("ObservacionCliente"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_llamadas_cliente(ByVal id_cliente As Integer) As List(Of CLlamadas) Implements IService1.Obtener_llamadas_cliente
        Dim Resultado As New List(Of CLlamadas)
        Dim cmd As New SqlCommand("Obtener_llamadas_Cliente", Conexion)
        cmd.Parameters.AddWithValue("@idCliente", id_cliente)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CLlamadas
        While reader.Read
            Aux = New CLlamadas
            Aux.id_llamada = DirectCast(reader.Item("id_llamada"), Integer)
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.id_usuario = DirectCast(reader.Item("id_usuario"), Integer)
            Aux.Fecha = DirectCast(reader.Item("Fecha"), Date)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.HoraProgramacion = reader.Item("HoraProgramacion")
            Aux.Programada = reader.Item("Programada")
            Aux.AvisoCliente = reader.Item("AvisoCliente")
            Aux.AvisoUsuario = reader.Item("AvisoUsuario")
            Aux.realizada = reader.Item("realizada")
            Aux.ObservacionUsuario = DirectCast(reader.Item("ObservacionUsuario"), String)
            Aux.ObservacionCliente = DirectCast(reader.Item("ObservacionCliente"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "nivelInteres"
    Function Inserta_nivelinteres(ByVal nivelinteres As String) As Boolean Implements IService1.Inserta_nivelinteres

        Dim cmd As New SqlCommand("Inserta_nivelInteres", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pnivelinteres", nivelinteres)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_nivelinteres(ByVal id_nivelInteres As Integer) As Boolean Implements IService1.Elimina_nivelinteres

        Dim cmd As New SqlCommand("Elimina_nivelinteres", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_nivelInteres", id_nivelInteres)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_nivelinteres(ByVal id_nivelInteres As Integer, ByVal nivelinteres As String) As Boolean Implements IService1.Actualiza_nivelinteres

        Dim cmd As New SqlCommand("Actualiza_nivelinteres", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_nivelInteres", id_nivelInteres)
        cmd.Parameters.AddWithValue("@Pnivelinteres", nivelinteres)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_nivelinteres() As List(Of CNivelInteres) Implements IService1.Obtener_nivelinteres
        Dim Resultado As New List(Of CNivelInteres)
        Dim cmd As New SqlCommand("Obtener_nivelInteres", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CNivelInteres
        While reader.Read
            Aux = New CNivelInteres
            Aux.id_nivelInteres = DirectCast(reader.Item("id_nivelInteres"), Integer)
            Aux.nivelinteres = DirectCast(reader.Item("nivelinteres"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "Productos"
    Function Obtener_detallesProducto(ByVal idProducto As Integer) As CDetallesProducto Implements IService1.Obtener_detallesProducto

        Dim cmd As New SqlCommand("Obtener_detallesProducto", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idProducto", idProducto)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New CDetallesProducto
        While reader.Read

            Aux.id_producto = DirectCast(reader.Item("id_producto"), Integer)
            Aux.NombreCorto = DirectCast(reader.Item("NombreCorto"), String)
            Aux.NombreCompleto = DirectCast(reader.Item("NombreCompleto"), String)
            Aux.Descripcion = DirectCast(reader.Item("Descripcion"), String)
            Aux.PrecioNormal = DirectCast(reader.Item("PrecioNormal"), Integer)
            Aux.PrecioDescuento = DirectCast(reader.Item("PrecioDescuento"), Integer)
            Aux.id_categoria = DirectCast(reader.Item("id_categoria"), Integer)
            Aux.categoria = DirectCast(reader.Item("categoria"), String)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)

        End While
        Conexion.Close()
        Return Aux
    End Function
    Function Obtener_datos_comboProductos() As List(Of CComboProductos) Implements IService1.Obtener_datos_comboProductos
        Dim Resultado As New List(Of CComboProductos)
        Dim cmd As New SqlCommand("Obtener_productos_idNombreCorto", Conexion)
        'cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CComboProductos
        While reader.Read
            Aux = New CComboProductos
            Aux.id_producto = DirectCast(reader.Item("id_producto"), Integer)
            Aux.NombreCorto = DirectCast(reader.Item("NombreCorto"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Inserta_productos(ByVal NombreCorto As String, ByVal NombreCompleto As String, ByVal Descripcion As String, ByVal PrecioNormal As Integer, ByVal PrecioDescuento As Integer, ByVal id_categoria As Integer, ByVal fechaCreacion As Date, ByVal Observaciones As String, ByVal Fotografia As String) As Boolean Implements IService1.Inserta_productos

        Dim cmd As New SqlCommand("Inserta_producto", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@PNombreCorto", NombreCorto)
        cmd.Parameters.AddWithValue("@PNombreCompleto", NombreCompleto)
        cmd.Parameters.AddWithValue("@PDescripcion", Descripcion)
        cmd.Parameters.AddWithValue("@PPrecioNormal", PrecioNormal)
        cmd.Parameters.AddWithValue("@PPrecioDescuento", PrecioDescuento)
        cmd.Parameters.AddWithValue("@Pid_categoria", id_categoria)
        cmd.Parameters.AddWithValue("@PfechaCreacion", fechaCreacion)
        cmd.Parameters.AddWithValue("@PObservaciones", Observaciones)
        'cmd.Parameters.AddWithValue("@Pfotografia", Fotografia)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_productos(ByVal id_producto As Integer) As Boolean Implements IService1.Elimina_productos

        Dim cmd As New SqlCommand("Elimina_productos", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_producto", id_producto)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_productos(ByVal id_producto As Integer, ByVal NombreCorto As String, ByVal NombreCompleto As String, ByVal Descripcion As String, ByVal PrecioNormal As Integer, ByVal PrecioDescuento As Integer, ByVal id_Categoria As Integer, ByVal Observaciones As String) As Boolean Implements IService1.Actualiza_productos

        Dim cmd As New SqlCommand("Actualiza_productos", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_producto", id_producto)
        cmd.Parameters.AddWithValue("@PNombreCorto", NombreCorto)
        cmd.Parameters.AddWithValue("@PNombreCompleto", NombreCompleto)
        cmd.Parameters.AddWithValue("@PDescripcion", Descripcion)
        cmd.Parameters.AddWithValue("@PPrecioNormal", PrecioNormal)
        cmd.Parameters.AddWithValue("@PPrecioDescuento", PrecioDescuento)
        'cmd.Parameters.AddWithValue("@Pid_categoria", Obtener_idCategoria_DescripcionCat(DescCat))
        cmd.Parameters.AddWithValue("@Pid_categoria", id_Categoria)
        'cmd.Parameters.AddWithValue("@PfechaCreacion", fechaCreacion)
        cmd.Parameters.AddWithValue("@PObservaciones", Observaciones)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_idCategoria_DescripcionCat(ByVal descCat As String) As Integer
        Dim Resultado As Integer = 0
        Dim cmd As New SqlCommand("Obtener_idCategoria_DescripcionCat", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Cat", descCat)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader

        While reader.Read
            Resultado = DirectCast(reader.Item(0), Integer)

        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_productos() As List(Of CProductos) Implements IService1.Obtener_productos
        Dim Resultado As New List(Of CProductos)
        Dim cmd As New SqlCommand("Obtener_productos", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CProductos
        While reader.Read
            Aux = New CProductos
            Aux.id_producto = DirectCast(reader.Item("id_producto"), Integer)
            Aux.NombreCorto = DirectCast(reader.Item("NombreCorto"), String)
            Aux.NombreCompleto = DirectCast(reader.Item("NombreCompleto"), String)
            Aux.Descripcion = DirectCast(reader.Item("Descripcion"), String)
            Aux.PrecioNormal = DirectCast(reader.Item("PrecioNormal"), Integer)
            Aux.PrecioDescuento = DirectCast(reader.Item("PrecioDescuento"), Integer)
            Aux.id_categoria = DirectCast(reader.Item("id_categoria"), Integer)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_Productos_Detalles() As List(Of CProductosDetalles) Implements IService1.Obtener_Productos_Detalles
        Dim Resultado As New List(Of CProductosDetalles)
        Dim cmd As New SqlCommand("Obtener_productos_detalles", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CProductosDetalles
        While reader.Read
            Aux = New CProductosDetalles
            Aux.id_producto = DirectCast(reader.Item("id_producto"), Integer)
            Aux.NombreCorto = DirectCast(reader.Item("NombreCorto"), String)
            Aux.Descripcion = DirectCast(reader.Item("Descripcion"), String)
            Aux.PrecioNormal = DirectCast(reader.Item("PrecioNormal"), Integer)
            Aux.PrecioDescuento = DirectCast(reader.Item("PrecioDescuento"), Integer)
            Aux.categoria = DirectCast(reader.Item("categoria"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            'Aux.fotografia = DirectCast(reader.Item("fotografia"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "Referencias"
    Function Inserta_referencias(ByVal id_cliente As Integer, ByVal Nombre As String, ByVal ApellidoPaterno As String, ByVal ApellidoMaterno As String, ByVal email As String, ByVal fechaCreacion As Date, ByVal id_tiporeferencia As Integer, ByVal id_usuario As Integer, ByVal Observaciones As String, ByVal fotografia As String, ByVal fotoTPresentacion As String) As Boolean Implements IService1.Inserta_referencias

        Dim cmd As New SqlCommand("Inserta_referencia", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_cliente", id_cliente)
        cmd.Parameters.AddWithValue("@PNombre", Nombre)
        cmd.Parameters.AddWithValue("@PApellidoPaterno", ApellidoPaterno)
        cmd.Parameters.AddWithValue("@PApellidoMaterno", ApellidoMaterno)
        cmd.Parameters.AddWithValue("@Pemail", email)
        cmd.Parameters.AddWithValue("@PfechaCreacion", fechaCreacion)
        cmd.Parameters.AddWithValue("@Pid_tiporeferencia", id_tiporeferencia)
        cmd.Parameters.AddWithValue("@Pid_usuario", id_usuario)
        cmd.Parameters.AddWithValue("@PObservaciones", Observaciones)
        cmd.Parameters.AddWithValue("@Pfotografia", fotografia)
        cmd.Parameters.AddWithValue("@PfotoTPresentacion", fotoTPresentacion)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_referencias(ByVal id_referencia As Integer) As Boolean Implements IService1.Elimina_referencias

        Dim cmd As New SqlCommand("Elimina_referencias", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_referencia", id_referencia)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_referencias(ByVal id_referencia As Integer, ByVal id_cliente As Integer, ByVal Nombre As String, ByVal ApellidoPaterno As String, ByVal ApellidoMaterno As String, ByVal email As String, ByVal fechaCreacion As Date, ByVal id_tiporeferencia As Integer, ByVal id_usuario As Integer, ByVal Observaciones As String, ByVal fotografia As String, ByVal fotoTPresentacion As String) As Boolean Implements IService1.Actualiza_referencias

        Dim cmd As New SqlCommand("Actualiza_referencias", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_referencia", id_referencia)
        cmd.Parameters.AddWithValue("@Pid_cliente", id_cliente)
        cmd.Parameters.AddWithValue("@PNombre", Nombre)
        cmd.Parameters.AddWithValue("@PApellidoPaterno", ApellidoPaterno)
        cmd.Parameters.AddWithValue("@PApellidoMaterno", ApellidoMaterno)
        cmd.Parameters.AddWithValue("@Pemail", email)
        cmd.Parameters.AddWithValue("@PfechaCreacion", fechaCreacion)
        cmd.Parameters.AddWithValue("@Pid_tiporeferencia", id_tiporeferencia)
        cmd.Parameters.AddWithValue("@Pid_usuario", id_usuario)
        cmd.Parameters.AddWithValue("@PObservaciones", Observaciones)
        cmd.Parameters.AddWithValue("@Pfotografia", fotografia)
        cmd.Parameters.AddWithValue("@PfotoTPresentacion", fotoTPresentacion)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_referencias_cliente(ByVal id_cliente As Integer) As List(Of CReferenciasCliente) Implements IService1.Obtener_referencias_cliente
        Dim Resultado As New List(Of CReferenciasCliente)
        Dim cmd As New SqlCommand("Obtener_referencias_cliente", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CReferenciasCliente
        While reader.Read
            Aux = New CReferenciasCliente
            Aux.id_referencia = DirectCast(reader.Item("id_referencia"), Integer)
            Aux.Nombre = DirectCast(reader.Item("Nombre"), String)
            Aux.ApellidoPaterno = DirectCast(reader.Item("ApellidoPaterno"), String)
            Aux.ApellidoMaterno = DirectCast(reader.Item("ApellidoMaterno"), String)
            Aux.email = DirectCast(reader.Item("email"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.tiporeferencia = DirectCast(reader.Item("tiporeferencia"), String)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Aux.fotografia = DirectCast(reader.Item("fotografia"), String)
            Aux.fotoTPresentacion = DirectCast(reader.Item("fotoTPresentacion"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "rubros"
    Function Inserta_rubros(ByVal rubro As String) As Boolean Implements IService1.Inserta_rubros

        Dim cmd As New SqlCommand("Inserta_rubros", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Prubro", rubro)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_rubros(ByVal id_rubro As Integer) As Boolean Implements IService1.Elimina_rubros

        Dim cmd As New SqlCommand("Elimina_rubros", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_rubro", id_rubro)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_rubros(ByVal id_rubro As Integer, ByVal rubro As String) As Boolean Implements IService1.Actualiza_rubros

        Dim cmd As New SqlCommand("Actualiza_rubros", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_rubro", id_rubro)
        cmd.Parameters.AddWithValue("@Prubro", rubro)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_rubros() As List(Of CRubros) Implements IService1.Obtener_rubros
        Dim Resultado As New List(Of CRubros)
        Dim cmd As New SqlCommand("Obtener_rubros", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CRubros
        While reader.Read
            Aux = New CRubros
            Aux.id_rubro = DirectCast(reader.Item("id_rubro"), Integer)
            Aux.rubro = DirectCast(reader.Item("rubro"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "Supervisores"
    Function Obtener_reporte_clientesFechas(ByVal FInicio As Date, ByVal FFinal As Date) As List(Of Obtener_reporte_clientesFechas) Implements IService1.Obtener_reporte_clientesFechas
        Dim Resultado As New List(Of Obtener_reporte_clientesFechas)
        Dim cmd As New SqlCommand("Obtener_reporte_clientesFechas", Conexion)
        cmd.Parameters.AddWithValue("@fechaInicio", FInicio)
        cmd.Parameters.AddWithValue("@fechaFinal", FFinal)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As Obtener_reporte_clientesFechas
        While reader.Read
            Aux = New Obtener_reporte_clientesFechas
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.Nombre = DirectCast(reader.Item("Nombre"), String)
            Aux.ApellidoPaterno = DirectCast(reader.Item("ApellidoPaterno"), String)
            Aux.ApellidoMaterno = DirectCast(reader.Item("ApellidoMaterno"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Aux.Producto = DirectCast(reader.Item("Producto"), String)
            Aux.nivelinteres = DirectCast(reader.Item("nivelinteres"), String)
            Aux.id_empresa = DirectCast(reader.Item("id_empresa"), Integer)
            Aux.FechaDe = DirectCast(reader.Item("FechaDe"), Date)
            Aux.Etapa = DirectCast(reader.Item("Etapa"), String)
            Aux.Campaña = DirectCast(reader.Item("Campaña"), String)
            Aux.usuario = DirectCast(reader.Item("usuario"), String)
            Aux.monto = DirectCast(reader.Item("monto"), Decimal)
            Aux.UltimoMovimiento = DirectCast(reader.Item("UltimoMovimiento"), String)
            Aux.UltimaObservacion = DirectCast(reader.Item("UltimaObservacion"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Actualiza_supervisores(ByVal id_usuario As Integer, ByVal nombre As String, ByVal apellidoPaterno As String, ByVal apellidoMaterno As String, ByVal Email As String, ByVal Activo As Integer) As Boolean Implements IService1.Actualiza_supervisores

        Dim cmd As New SqlCommand("Actualiza_supervisores", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_supervisor", id_usuario)
        cmd.Parameters.AddWithValue("@Pnombre", nombre)
        cmd.Parameters.AddWithValue("@PapellidoPaterno", apellidoPaterno)
        cmd.Parameters.AddWithValue("@PapellidoMaterno", apellidoMaterno)
        cmd.Parameters.AddWithValue("@PEmail", Email)
        cmd.Parameters.AddWithValue("@activo", Activo)
        'cmd.Parameters.AddWithValue("@Pusuario", usuario)
        'cmd.Parameters.AddWithValue("@Pcontraseña", contraseña)
        'cmd.Parameters.AddWithValue("@PfechaCreacion", fechaCreacion)
        'cmd.Parameters.AddWithValue("@Pfotografia", fotografia)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_contraseaSupervisor(ByVal id_usuario As Integer, ByVal Contraseña As String) As Boolean Implements IService1.Actualiza_contraseaSupervisor

        Dim cmd As New SqlCommand("Actualiza_contraseña_supervisor", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@id_supervisor", id_usuario)
        cmd.Parameters.AddWithValue("@MD5Contra", Contraseña)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_supervisor_Detalles(ByVal id_supervisor As Integer) As CDetallesSupervisor Implements IService1.Obtener_supervisor_Detalles

        Dim cmd As New SqlCommand("Obtener_supervisor_Detalles", Conexion)
        cmd.Parameters.AddWithValue("@idSupervisor", id_supervisor)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New CDetallesSupervisor
        While reader.Read

            Aux.is_supervisor = DirectCast(reader.Item("id_supervisor"), Integer)
            Aux.nombre = DirectCast(reader.Item("nombre"), String)
            Aux.apellidoPaterno = DirectCast(reader.Item("apellidoPaterno"), String)
            Aux.apellidoMaterno = DirectCast(reader.Item("apellidoMaterno"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Aux.usuario = DirectCast(reader.Item("usuario"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)


        End While
        Conexion.Close()
        Return Aux
    End Function
    Function DiasSinTrabajar(ByVal id_supervisor As Integer) As List(Of DiasSinTrabajar) Implements IService1.DiasSinTrabajar
        Dim Resultado As New List(Of DiasSinTrabajar)
        Dim cmd As New SqlCommand("DiasSinTrabajar", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idSupervisor", id_supervisor)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As DiasSinTrabajar
        While reader.Read
            Aux = New DiasSinTrabajar
            Aux.ID = DirectCast(reader.Item("ID"), Integer)
            Aux.Cliente = DirectCast(reader.Item("Cliente"), String)
            Aux.Ultima = reader.Item("Ultima")
            Aux.Dias = DirectCast(reader.Item("Dias"), Integer)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function DiasSinTrabajarEtapa(ByVal id_supervisor As Integer, ByVal Etapa As Integer) As List(Of DiasSinTrabajar) Implements IService1.DiasSinTrabajarEtapa
        Dim Resultado As New List(Of DiasSinTrabajar)
        Dim cmd As New SqlCommand("DiasSinTrabajarEtapa", Conexion)

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idSupervisor", id_supervisor)
        cmd.Parameters.AddWithValue("@Etapa", Etapa)

        Conexion.Close()
        Conexion.Open()

        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As DiasSinTrabajar
        While reader.Read
            Aux = New DiasSinTrabajar
            Aux.ID = DirectCast(reader.Item("ID"), Integer)
            Aux.Cliente = DirectCast(reader.Item("Cliente"), String)
            Aux.Ultima = reader.Item("Ultima")
            Aux.Dias = DirectCast(reader.Item("Dias"), Integer)
            Resultado.Add(Aux)
        End While
        Conexion.Close()

        Return Resultado
    End Function

    Function DiasSinTrabajarEtapaFiltro(ByVal id_supervisor As Integer, ByVal Etapa As Integer, ByVal Dias As Integer, ByVal FechaInicio As Date, ByVal FechaFinal As Date) As List(Of DiasSinTrabajar) Implements IService1.DiasSinTrabajarEtapaFiltro
        Dim Resultado As New List(Of DiasSinTrabajar)
        Dim cmd As New SqlCommand("DiasSinTrabajarEtapaFiltro", Conexion)

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idSupervisor", id_supervisor)
        cmd.Parameters.AddWithValue("@Etapa", Etapa)
        cmd.Parameters.AddWithValue("@Dias", Dias)
        cmd.Parameters.AddWithValue("@FechaInicio", FechaInicio)
        cmd.Parameters.AddWithValue("@FechaFin", FechaFinal)

        Conexion.Close()
        Conexion.Open()

        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As DiasSinTrabajar
        While reader.Read
            Aux = New DiasSinTrabajar
            Aux.ID = DirectCast(reader.Item("ID"), Integer)
            Aux.Cliente = DirectCast(reader.Item("Cliente"), String)
            Aux.Ultima = reader.Item("Ultima")
            Aux.Dias = DirectCast(reader.Item("Dias"), Integer)
            Resultado.Add(Aux)
        End While
        Conexion.Close()

        Return Resultado
    End Function

    Function DiasSinTrabajarFiltro(ByVal id_supervisor As Integer, ByVal Filtro As String) As List(Of DiasSinTrabajar) Implements IService1.DiasSinTrabajarFiltro
        Dim Resultado As New List(Of DiasSinTrabajar)
        Dim cmd As New SqlCommand("DiasSinTrabajarFiltro", Conexion)

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idSupervisor", id_supervisor)
        cmd.Parameters.AddWithValue("@Filtro", Filtro)

        Conexion.Close()
        Conexion.Open()

        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As DiasSinTrabajar
        While reader.Read
            Aux = New DiasSinTrabajar
            Aux.ID = DirectCast(reader.Item("ID"), Integer)
            Aux.Cliente = DirectCast(reader.Item("Cliente"), String)
            Aux.Ultima = reader.Item("Ultima")
            Aux.Dias = DirectCast(reader.Item("Dias"), Integer)
            Resultado.Add(Aux)
        End While
        Conexion.Close()

        Return Resultado
    End Function
    Function Obtener_nombresClientesidSupervisor(ByVal id_supervisor As Integer) As List(Of CCLientesSupervisor) Implements IService1.Obtener_nombresClientesidSupervisor
        Dim Resultado As New List(Of CCLientesSupervisor)
        Dim cmd As New SqlCommand("Obtener_nombresClientesidSupervisor", Conexion)
        cmd.Parameters.AddWithValue("@idSupervisor", id_supervisor)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CCLientesSupervisor
        While reader.Read
            Aux = New CCLientesSupervisor
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.Cliente = DirectCast(reader.Item("Cliente"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function

    Function Obtener_nombresUsuariosSupervisor(ByVal id_supervisor As Integer) As List(Of CSupervisorUsuarios) Implements IService1.Obtener_nombresUsuariosSupervisor
        Dim Resultado As New List(Of CSupervisorUsuarios)
        Dim cmd As New SqlCommand("Obtener_nombresUsuariosSupervisor", Conexion)
        cmd.Parameters.AddWithValue("@idSupervisor", id_supervisor)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CSupervisorUsuarios
        While reader.Read
            Aux = New CSupervisorUsuarios
            Aux.id_usuario = DirectCast(reader.Item("id_usuario"), Integer)
            Aux.Usuario = DirectCast(reader.Item("Usuario"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Cambia_usuarioCliente(ByVal id_usuario As Integer, ByVal idCliente As Integer) As Boolean Implements IService1.Cambia_usuarioCliente

        Dim cmd As New SqlCommand("Cambia_usuarioCliente", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@id_usuario", id_usuario)
        cmd.Parameters.AddWithValue("@idCliente", idCliente)

        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Cambia_usuarioClienteSupervisor(ByVal id_usuario As Integer, ByVal idCliente As Integer, ByVal idSupervisor As Integer) As Boolean Implements IService1.Cambia_usuarioClienteSupervisor

        Dim cmd As New SqlCommand("Cambia_usuarioClienteSupervisor", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@id_usuario", id_usuario)
        cmd.Parameters.AddWithValue("@idCliente", idCliente)
        cmd.Parameters.AddWithValue("@idSupervisor", idSupervisor)

        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function

    Function Obtener_AcumuladosSupervisor(ByVal id_supervisor As Integer, ByVal FechaInicio As Date, ByVal FechaFinal As Date) As List(Of CAcumuladosSupervisor) Implements IService1.Obtener_AcumuladosSupervisor
        Dim Resultado As New List(Of CAcumuladosSupervisor)
        Dim cmd As New SqlCommand("Obtener_AcumuladosSupervisor", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idSupervisor", id_supervisor)
        cmd.Parameters.AddWithValue("@FechaInicio", FechaInicio)
        cmd.Parameters.AddWithValue("@FechaFinal", FechaFinal)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CAcumuladosSupervisor
        While reader.Read
            Aux = New CAcumuladosSupervisor
            Aux.Cantidad = DirectCast(reader.Item("Cantidad"), Integer)
            Aux.NombreCliente = DirectCast(reader.Item("NombreCliente"), String)
            Aux.Producto = DirectCast(reader.Item("Producto"), String)
            Aux.Empresa = DirectCast(reader.Item("Empresa"), String)
            Aux.Etapa = DirectCast(reader.Item("Etapa"), String)
            Aux.Usuario = DirectCast(reader.Item("Usuario"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_UsuarioDetalleSupervisor(ByVal id_supervisor As Integer) As List(Of CUsuariosDetalleSup) Implements IService1.Obtener_UsuarioDetalleSupervisor
        Dim Resultado As New List(Of CUsuariosDetalleSup)
        Dim cmd As New SqlCommand("Obtener_UsuarioDetalleSupervisor", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idSupervisor", id_supervisor)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CUsuariosDetalleSup
        While reader.Read
            Aux = New CUsuariosDetalleSup
            Aux.id_usuario = DirectCast(reader.Item("id_usuario"), Integer)
            Aux.nombre = DirectCast(reader.Item("nombre"), String)
            Aux.apellidoPaterno = DirectCast(reader.Item("apellidoPaterno"), String)
            Aux.apellidoMaterno = DirectCast(reader.Item("apellidoMaterno"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Aux.usuario = DirectCast(reader.Item("usuario"), String)
            Aux.contraseña = DirectCast(reader.Item("contraseña"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.activo = reader.Item("activo")
            Aux.id_TipoUsuario = DirectCast(reader.Item("id_tipoUsuario"), Integer)
            Aux.TipousuarioDes = DirectCast(reader.Item("TipoUsuario"), String)
            Aux.id_supervisor = If(IsDBNull(reader("id_supervisor")), 0, Convert.ToInt32(reader("id_supervisor")))
            Aux.SupervisorDes = DirectCast(reader.Item("SupervisorDes"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_DetalleSupervisor() As List(Of CDetalleSupervisor) Implements IService1.Obtener_DetalleSupervisor
        Dim Resultado As New List(Of CDetalleSupervisor)
        Dim cmd As New SqlCommand("Obtener_DetalleSupervisor", Conexion)

        cmd.CommandType = CommandType.StoredProcedure

        Conexion.Close()
        Conexion.Open()
        Dim Reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CDetalleSupervisor
        While Reader.Read
            Aux = New CDetalleSupervisor
            Aux.id_supervisor = DirectCast(Reader.Item("id_supervisor"), Integer)
            Aux.nombre = DirectCast(Reader.Item("nombre"), String)
            Aux.apellidoPaterno = DirectCast(Reader.Item("apellidoPaterno"), String)
            Aux.apellidoMaterno = DirectCast(Reader.Item("apellidoMaterno"), String)
            Aux.Email = DirectCast(Reader.Item("Email"), String)
            Aux.usuario = DirectCast(Reader.Item("usuario"), String)
            Aux.contraseña = DirectCast(Reader.Item("contraseña"), String)
            Aux.fechaCreacion = DirectCast(Reader.Item("fechaCreacion"), Date)
            Aux.activo = Reader.Item("activo")
            Aux.BorraEK = Reader.Item("BorraEk")

            Resultado.Add(Aux)
        End While

        Conexion.Close()

        Return Resultado
    End Function
    Function Inserta_supervisores(ByVal nombre As String, ByVal apellidoPaterno As String, ByVal apellidoMaterno As String, ByVal email As String, ByVal usuario As String, ByVal contraseña As String, ByVal fechaCreacion As Date, ByVal fotografia As String) As Boolean Implements IService1.Inserta_supervisores
        Dim cmd As New SqlCommand("Inserta_supervisor", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pnombre", nombre)
        cmd.Parameters.AddWithValue("@PapellidoPaterno", apellidoPaterno)
        cmd.Parameters.AddWithValue("@PapellidoMaterno", apellidoMaterno)
        cmd.Parameters.AddWithValue("@Pemail", email)
        cmd.Parameters.AddWithValue("@Pusuario", usuario)
        cmd.Parameters.AddWithValue("@Pcontraseña", contraseña)
        cmd.Parameters.AddWithValue("@PfechaCreacion", fechaCreacion)
        cmd.Parameters.AddWithValue("@Pfotografia", fotografia)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_supervisores(ByVal id_supervisor As Integer) As Boolean Implements IService1.Elimina_supervisores
        Dim cmd As New SqlCommand("Elimina_supervisores", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_supervisor", id_supervisor)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function

    Function Obtener_supervisores() As List(Of CSupervisores) Implements IService1.Obtener_supervisores
        Dim Resultado As New List(Of CSupervisores)
        Dim cmd As New SqlCommand("Obtener_supervisores", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CSupervisores
        While reader.Read
            Aux = New CSupervisores
            Aux.id_supervisor = DirectCast(reader.Item("id_supervisor"), Integer)
            Aux.NombreCompleto = DirectCast(reader.Item("supervisor"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Public Function Obtener_TipoUsuario() As List(Of TipoUsuario) Implements IService1.Obtener_TipoUsuario
        Dim Resultado As New List(Of TipoUsuario)
        Dim cmd As New SqlCommand("SELECT * FROM TipoUsuarios", Conexion)

        Conexion.Close()
        Conexion.Open()

        Dim Reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As TipoUsuario
        While Reader.Read
            Aux = New TipoUsuario
            Aux.id_tipoUsuario = Reader.Item("id_tipoUsuario")
            Aux.Tipo = Reader.Item("TipoUsuario")

            Resultado.Add(Aux)
        End While

        Conexion.Close()

        Return Resultado
    End Function

    Public Function Actualizar_Coordinador(ByVal NumEmpleado As Integer, ByVal NumCordinador As Integer, ByVal Nombre_Cordinador As String) As Boolean Implements IService1.Actualizar_Coordinador
        Dim cmd As New SqlCommand("Actualiza_Coordinador", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@PNumEmpleado", NumEmpleado)
        cmd.Parameters.AddWithValue("@PNumCordinador", NumCordinador)
        cmd.Parameters.AddWithValue("@PNombre_Cordinador", Nombre_Cordinador)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
#End Region
#Region "Oportunidades"
    Function Obtener_oportunidades() As List(Of COportunidades) Implements IService1.Obtener_oportunidades
        Dim Resultado As New List(Of COportunidades)
        Dim cmd As New SqlCommand("Obtener_oportunidades", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As COportunidades
        While reader.Read
            Aux = New COportunidades
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.Nombre = DirectCast(reader.Item("Nombre"), String)
            Aux.ApellidoPaterno = DirectCast(reader.Item("ApellidoPaterno"), String)
            Aux.ApellidoMaterno = DirectCast(reader.Item("ApellidoMaterno"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Aux.NombreCorto = DirectCast(reader.Item("NombreCorto"), String)
            Aux.nivelinteres = DirectCast(reader.Item("nivelinteres"), String)
            Aux.Empresa = DirectCast(reader.Item("Empresa"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.Etapa = DirectCast(reader.Item("Etapa"), String)
            Aux.campañaNombre = DirectCast(reader.Item("campañaNombre"), String)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "supervisorUsuario"
    Function Obtener_usuariosSupervisor(ByVal id_supervisor As Integer) As List(Of CUsuariosSupervisor) Implements IService1.Obtener_usuariosSupervisor
        Dim Resultado As New List(Of CUsuariosSupervisor)
        Dim cmd As New SqlCommand("Obtener_usuariosSupervisor", Conexion)
        cmd.Parameters.AddWithValue("@idSupervisor", id_supervisor)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CUsuariosSupervisor
        While reader.Read
            Aux = New CUsuariosSupervisor
            Aux.id_usuario = DirectCast(reader.Item("id_usuario"), Integer)
            Aux.nombre = DirectCast(reader.Item("nombre"), String)
            Aux.apellidoPaterno = DirectCast(reader.Item("apellidoPaterno"), String)
            Aux.apellidoMaterno = DirectCast(reader.Item("apellidoMaterno"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Aux.usuario = DirectCast(reader.Item("usuario"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Inserta_supervisorUsuario(ByVal id_usuario As Integer, ByVal id_supervisor As Integer) As Boolean Implements IService1.Inserta_supervisorUsuario

        Dim cmd As New SqlCommand("Inserta_supervisorUsuario", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_usuario", id_usuario)
        cmd.Parameters.AddWithValue("@Pid_supervisor", id_supervisor)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_supervisorUsuario(ByVal id_supervisorusuario As Integer) As Boolean Implements IService1.Elimina_supervisorUsuario

        Dim cmd As New SqlCommand("Elimina_supervisorUsuario", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_supervisorusuario", id_supervisorusuario)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_supervisorUsuario(ByVal id_supervisorusuario As Integer, ByVal id_usuario As Integer, ByVal id_supervisor As Integer) As Boolean Implements IService1.Actualiza_supervisorUsuario

        Dim cmd As New SqlCommand("Actualiza_supervisorUsuario", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_supervisorusuario", id_supervisorusuario)
        cmd.Parameters.AddWithValue("@Pid_usuario", id_usuario)
        cmd.Parameters.AddWithValue("@Pid_supervisor", id_supervisor)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_relacion_supervisorUsuario() As List(Of CRelacionSupervisorUsuario) Implements IService1.Obtener_relacion_supervisorUsuario
        Dim Resultado As New List(Of CRelacionSupervisorUsuario)
        Dim cmd As New SqlCommand("Obtener_supervisorUsuario", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CRelacionSupervisorUsuario
        While reader.Read
            Aux = New CRelacionSupervisorUsuario
            Aux.id_supervisor = DirectCast(reader.Item("id_supervisorusuario"), Integer)
            Aux.id_supervisor = DirectCast(reader.Item("id_supervisor"), Integer)
            Aux.nombre = DirectCast(reader.Item("nombre"), String)
            Aux.apellidoPaterno = DirectCast(reader.Item("apellidoPaterno"), String)
            Aux.apellidoMaterno = DirectCast(reader.Item("apellidoMaterno"), String)
            Aux.email = DirectCast(reader.Item("email"), String)
            Aux.id_usuario = DirectCast(reader.Item("id_usuario"), Integer)
            Aux.nombre1 = DirectCast(reader.Item("nombreUsuario"), String)
            Aux.apellidoPaterno1 = DirectCast(reader.Item("apellidoPaternoUsuario"), String)
            Aux.apellidoMaterno1 = DirectCast(reader.Item("apellidoMaternoUsuario"), String)
            Aux.Email1 = DirectCast(reader.Item("EmailUsuario"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_clientesSupervisor(ByVal idSupervisor As Integer) As List(Of ClientesSupervisor) Implements IService1.Obtener_clientesSupervisor
        Dim Resultado As New List(Of ClientesSupervisor)
        Dim cmd As New SqlCommand("Obtener_clientesSupervisor", Conexion)
        cmd.Parameters.AddWithValue("@idSupervisor", idSupervisor)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As ClientesSupervisor
        While reader.Read
            Aux = New ClientesSupervisor
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.Nombre = DirectCast(reader.Item("Nombre"), String)
            Aux.ApellidoPaterno = DirectCast(reader.Item("ApellidoPaterno"), String)
            Aux.ApellidoMaterno = DirectCast(reader.Item("ApellidoMaterno"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Aux.Producto = DirectCast(reader.Item("Producto"), String)
            Aux.Empresa = DirectCast(reader.Item("Empresa"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.Descripcion = DirectCast(reader.Item("Descripcion"), String)
            Aux.Usuario = DirectCast(reader.Item("Usuario"), String)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Aux.fotografia = DirectCast(reader.Item("fotografia"), String)
            Aux.fotoTpresentacion = DirectCast(reader.Item("fotoTpresentacion"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "TelefonoCliente"
    Function Obtener_telefonosModificaCliente(ByVal idCliente As Integer) As List(Of CTelefonosmodifica) Implements IService1.Obtener_telefonosModificaCliente
        Dim Resultado As New List(Of CTelefonosmodifica)
        Dim cmd As New SqlCommand("Obtener_telefonosClienteModifica", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idCliente", idCliente)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CTelefonosmodifica
        While reader.Read
            Aux = New CTelefonosmodifica
            Aux.id_telefonoCliente = DirectCast(reader.Item("id_telefonoCliente"), Integer)
            Aux.Principal = reader.Item("Principal")
            Aux.Telefono = DirectCast(reader.Item("Telefono"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Inserta_telefonoCliente(ByVal Principal As String, ByVal id_cliente As Integer, ByVal Telefono As String) As Boolean Implements IService1.Inserta_telefonoCliente

        Dim cmd As New SqlCommand("Inserta_telefonoCliente", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@PPrincipal", Principal)
        cmd.Parameters.AddWithValue("@Pid_cliente", id_cliente)
        cmd.Parameters.AddWithValue("@PTelefono", Telefono)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_telefonoCliente(ByVal id_telefonoCliente As Integer) As Boolean Implements IService1.Elimina_telefonoCliente

        Dim cmd As New SqlCommand("Elimina_telefonoCliente", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_telefonoCliente", id_telefonoCliente)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_telefonoCliente(ByVal id_telefonoCliente As Integer, ByVal Principal As String, ByVal id_cliente As Integer, ByVal Telefono As String) As Boolean Implements IService1.Actualiza_telefonoCliente

        Dim cmd As New SqlCommand("Actualiza_telefonoCliente", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_telefonoCliente", id_telefonoCliente)
        cmd.Parameters.AddWithValue("@PPrincipal", Principal)
        cmd.Parameters.AddWithValue("@Pid_cliente", id_cliente)
        cmd.Parameters.AddWithValue("@PTelefono", Telefono)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_telefonoCliente(ByVal id_cliente As Integer) As List(Of CTelefonoCliente) Implements IService1.Obtener_telefonoCliente
        Dim Resultado As New List(Of CTelefonoCliente)
        Dim cmd As New SqlCommand("Obtener_telefonosCliente", Conexion)
        cmd.Parameters.AddWithValue("@idCliente", id_cliente)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CTelefonoCliente
        While reader.Read
            Aux = New CTelefonoCliente
            Aux.id_telefonoCliente = DirectCast(reader.Item("id_telefonoCliente"), Integer)
            Aux.Principal = reader.Item("Principal")
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.Telefono = DirectCast(reader.Item("Telefono"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "telefonoContactoEmpresa"
    Function Inserta_telefonoContactoEmpresa(ByVal id_contactoEmpresa As Integer, ByVal principal As String, ByVal Telefono As String) As Boolean Implements IService1.Inserta_telefonoContactoEmpresa

        Dim cmd As New SqlCommand("Inserta_telefonoContactoEmpresa", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_contactoEmpresa", id_contactoEmpresa)
        cmd.Parameters.AddWithValue("@Pprincipal", principal)
        cmd.Parameters.AddWithValue("@PTelefono", Telefono)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_telefonoContactoEmpresa(ByVal id_telefonoContacto As Integer, ByVal id_contactoEmpresa As Integer, ByVal principal As String, ByVal Telefono As String) As Boolean Implements IService1.Actualiza_telefonoContactoEmpresa

        Dim cmd As New SqlCommand("Actualiza_telefonoContactoEmpresa", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_telefonoContacto", id_telefonoContacto)
        cmd.Parameters.AddWithValue("@Pid_contactoEmpresa", id_contactoEmpresa)
        cmd.Parameters.AddWithValue("@Pprincipal", principal)
        cmd.Parameters.AddWithValue("@PTelefono", Telefono)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_telefonoContactoEmpresa(ByVal id_telefonoContacto As Integer) As Boolean Implements IService1.Elimina_telefonoContactoEmpresa

        Dim cmd As New SqlCommand("Elimina_telefonoContactoEmpresa", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_telefonoContacto", id_telefonoContacto)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_telefonoContactoEmpresa(ByVal id_contactoempresa As Integer) As List(Of CTelefonoContactoEmpresa) Implements IService1.Obtener_telefonoContactoEmpresa
        Dim Resultado As New List(Of CTelefonoContactoEmpresa)
        Dim cmd As New SqlCommand("Obtener_telefonosContactoEmpresa", Conexion)
        cmd.Parameters.AddWithValue("@idContactoEmpresa", id_contactoempresa)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CTelefonoContactoEmpresa
        While reader.Read
            Aux = New CTelefonoContactoEmpresa
            Aux.id_telefonoContacto = DirectCast(reader.Item("id_telefonoContacto"), Integer)
            Aux.id_contactoEmpresa = DirectCast(reader.Item("id_contactoEmpresa"), Integer)
            Aux.principal = reader.Item("principal")
            Aux.Telefono = DirectCast(reader.Item("Telefono"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "TelefonoEmpresa"
    Function Inserta_TelefonoEmpresa(ByVal Principal As String, ByVal id_empresa As Integer, ByVal telefono As String) As Boolean Implements IService1.Inserta_TelefonoEmpresa

        Dim cmd As New SqlCommand("Inserta_telefonoEmpresa", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@PPrincipal", Principal)
        cmd.Parameters.AddWithValue("@Pid_empresa", id_empresa)
        cmd.Parameters.AddWithValue("@Ptelefono", telefono)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_TelefonoEmpresa(ByVal id_telefonoEmpresa As Integer) As Boolean Implements IService1.Elimina_TelefonoEmpresa

        Dim cmd As New SqlCommand("Elimina_TelefonoEmpresa", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_telefonoEmpresa", id_telefonoEmpresa)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function

    Function Actualiza_TelefonoEmpresa(ByVal id_telefonoEmpresa As Integer, ByVal Principal As String, ByVal id_empresa As Integer, ByVal telefono As String) As Boolean Implements IService1.Actualiza_TelefonoEmpresa

        Dim cmd As New SqlCommand("Actualiza_TelefonoEmpresa", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_telefonoEmpresa", id_telefonoEmpresa)
        cmd.Parameters.AddWithValue("@PPrincipal", Principal)
        cmd.Parameters.AddWithValue("@Pid_empresa", id_empresa)
        cmd.Parameters.AddWithValue("@Ptelefono", telefono)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_TelefonoEmpresa(ByVal id_empresa As Integer) As List(Of CTelefonoEmpresa) Implements IService1.Obtener_TelefonoEmpresa
        Dim Resultado As New List(Of CTelefonoEmpresa)
        Dim cmd As New SqlCommand("Obtener_telefonosEmpresas", Conexion)
        cmd.Parameters.AddWithValue("@idEmpresa", id_empresa)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CTelefonoEmpresa
        While reader.Read
            Aux = New CTelefonoEmpresa
            Aux.id_telefonoEmpresa = DirectCast(reader.Item("id_telefonoEmpresa"), Integer)
            Aux.Principal = reader.Item("Principal")
            Aux.id_empresa = DirectCast(reader.Item("id_empresa"), Integer)
            Aux.telefono = DirectCast(reader.Item("telefono"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "Telefono Referencia"
    Function Inserta_telefonoReferencia(ByVal Principal As String, ByVal id_referencia As Integer, ByVal Telefono As String) As Boolean Implements IService1.Inserta_telefonoReferencia

        Dim cmd As New SqlCommand("Inserta_telefonoReferencia", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@PPrincipal", Principal)
        cmd.Parameters.AddWithValue("@Pid_referencia", id_referencia)
        cmd.Parameters.AddWithValue("@PTelefono", Telefono)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_telefonoReferencia(ByVal id_telefonoReferencia As Integer) As Boolean Implements IService1.Elimina_telefonoReferencia

        Dim cmd As New SqlCommand("Elimina_telefonoReferencia", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_telefonoReferencia", id_telefonoReferencia)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_telefonoReferencia(ByVal id_telefonoReferencia As Integer, ByVal Principal As String, ByVal id_referencia As Integer, ByVal Telefono As String) As Boolean Implements IService1.Actualiza_telefonoReferencia

        Dim cmd As New SqlCommand("Actualiza_telefonoReferencia", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_telefonoReferencia", id_telefonoReferencia)
        cmd.Parameters.AddWithValue("@PPrincipal", Principal)
        cmd.Parameters.AddWithValue("@Pid_referencia", id_referencia)
        cmd.Parameters.AddWithValue("@PTelefono", Telefono)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_telefonoReferencia(ByVal id_referencia As Integer) As List(Of CTelefonoReferencia) Implements IService1.Obtener_telefonoReferencia
        Dim Resultado As New List(Of CTelefonoReferencia)
        Dim cmd As New SqlCommand("Obtener_telefonos_referencia", Conexion)
        cmd.Parameters.AddWithValue("@idReferencia", id_referencia)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CTelefonoReferencia
        While reader.Read
            Aux = New CTelefonoReferencia
            Aux.id_telefonoReferencia = DirectCast(reader.Item("id_telefonoReferencia"), Integer)
            Aux.Principal = reader.Item("Principal")
            Aux.id_referencia = DirectCast(reader.Item("id_referencia"), Integer)
            Aux.Telefono = DirectCast(reader.Item("Telefono"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "Telefono Supervisor"
    Function Inserta_telefonoSupervisor(ByVal Principal As String, ByVal id_supervisor As Integer, ByVal Telefono As String) As Boolean Implements IService1.Inserta_telefonoSupervisor

        Dim cmd As New SqlCommand("Inserta_telefonoSupervisor", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@PPrincipal", Principal)
        cmd.Parameters.AddWithValue("@Pid_supervisor", id_supervisor)
        cmd.Parameters.AddWithValue("@PTelefono", Telefono)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_telefonoSupervisor(ByVal id_telefonoSupervisor As Integer) As Boolean Implements IService1.Elimina_telefonoSupervisor

        Dim cmd As New SqlCommand("Elimina_telefonoSupervisor", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_telefonoSupervisor", id_telefonoSupervisor)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_telefonoSupervisor(ByVal id_telefonoSupervisor As Integer, ByVal Principal As String, ByVal id_supervisor As Integer, ByVal Telefono As String) As Boolean Implements IService1.Actualiza_telefonoSupervisor

        Dim cmd As New SqlCommand("Actualiza_telefonoSupervisor", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_telefonoSupervisor", id_telefonoSupervisor)
        cmd.Parameters.AddWithValue("@PPrincipal", Principal)
        cmd.Parameters.AddWithValue("@Pid_supervisor", id_supervisor)
        cmd.Parameters.AddWithValue("@PTelefono", Telefono)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_telefonoSupervisor(ByVal id_supervisor As Integer) As List(Of CTelefonoSupervisor) Implements IService1.Obtener_telefonoSupervisor
        Dim Resultado As New List(Of CTelefonoSupervisor)
        Dim cmd As New SqlCommand("Obtener_telefonoSupervisor", Conexion)
        cmd.Parameters.AddWithValue("@idSupervisor", id_supervisor)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CTelefonoSupervisor
        While reader.Read
            Aux = New CTelefonoSupervisor
            Aux.id_telefonoSupervisor = DirectCast(reader.Item("id_telefonoSupervisor"), Integer)
            Aux.Principal = reader.Item("Principal")
            Aux.id_supervisor = DirectCast(reader.Item("id_supervisor"), Integer)
            Aux.Telefono = DirectCast(reader.Item("Telefono"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "Telefono Usuario"
    Function Inserta_telefonoUsuario(ByVal Principal As String, ByVal id_usuario As Integer, ByVal Telefono As String) As Boolean Implements IService1.Inserta_telefonoUsuario

        Dim cmd As New SqlCommand("Inserta_telefonoUsuario", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@PPrincipal", Principal)
        cmd.Parameters.AddWithValue("@Pid_usuario", id_usuario)
        cmd.Parameters.AddWithValue("@PTelefono", Telefono)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_telefonoUsuario(ByVal id_telefonoUsuario As Integer) As Boolean Implements IService1.Elimina_telefonoUsuario

        Dim cmd As New SqlCommand("Elimina_telefonoUsuario", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_telefonoUsuario", id_telefonoUsuario)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_telefonoUsuario(ByVal id_telefonoUsuario As Integer, ByVal Principal As String, ByVal id_usuario As Integer, ByVal Telefono As String) As Boolean Implements IService1.Actualiza_telefonoUsuario

        Dim cmd As New SqlCommand("Actualiza_telefonoUsuario", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_telefonoUsuario", id_telefonoUsuario)
        cmd.Parameters.AddWithValue("@PPrincipal", Principal)
        cmd.Parameters.AddWithValue("@Pid_usuario", id_usuario)
        cmd.Parameters.AddWithValue("@PTelefono", Telefono)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_telefonoUsuario(ByVal id_usuario As Integer) As List(Of CTelefonoUsuario) Implements IService1.Obtener_telefonoUsuario
        Dim Resultado As New List(Of CTelefonoUsuario)
        Dim cmd As New SqlCommand("Obtener_telefonosUsuario", Conexion)
        cmd.Parameters.AddWithValue("@idUsuario", id_usuario)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CTelefonoUsuario
        While reader.Read
            Aux = New CTelefonoUsuario
            Aux.id_telefonoUsuario = DirectCast(reader.Item("id_telefonoUsuario"), Integer)
            Aux.Principal = reader.Item("Principal")
            Aux.id_usuario = DirectCast(reader.Item("id_usuario"), Integer)
            Aux.Telefono = DirectCast(reader.Item("Telefono"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "Tipo Campaña"
    Function Inserta_tipocampaña(ByVal TipoCampaña As String) As Boolean Implements IService1.Inserta_tipocampaña

        Dim cmd As New SqlCommand("Inserta_campañas", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@PTipoCampaña", TipoCampaña)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_tipocampaña(ByVal id_tipoCampaña As Integer) As Boolean Implements IService1.Elimina_tipocampaña

        Dim cmd As New SqlCommand("Elimina_tipocampaña", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_tipoCampaña", id_tipoCampaña)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_tipocampaña(ByVal id_tipoCampaña As Integer, ByVal TipoCampaña As String) As Boolean Implements IService1.Actualiza_tipocampaña

        Dim cmd As New SqlCommand("Actualiza_tipocampaña", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_tipoCampaña", id_tipoCampaña)
        cmd.Parameters.AddWithValue("@PTipoCampaña", TipoCampaña)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_tipocampaña() As List(Of CTipoCampaña) Implements IService1.Obtener_tipocampaña
        Dim Resultado As New List(Of CTipoCampaña)
        Dim cmd As New SqlCommand("Obtener_tiposCampaña", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CTipoCampaña
        While reader.Read
            Aux = New CTipoCampaña
            Aux.id_tipoCampaña = DirectCast(reader.Item("id_tipoCampaña"), Integer)
            Aux.TipoCampaña = DirectCast(reader.Item("TipoCampaña"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "tipo Referencia"
    Function Inserta_tiporeferencia(ByVal tiporeferencia As String) As Boolean Implements IService1.Inserta_tiporeferencia

        Dim cmd As New SqlCommand("Inserta_TipoReferencia", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Ptiporeferencia", tiporeferencia)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_tiporeferencia(ByVal id_tiporeferencia As Integer) As Boolean Implements IService1.Elimina_tiporeferencia

        Dim cmd As New SqlCommand("Elimina_tiporeferencia", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_tiporeferencia", id_tiporeferencia)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_tiporeferencia(ByVal id_tiporeferencia As Integer, ByVal tiporeferencia As String) As Boolean Implements IService1.Actualiza_tiporeferencia

        Dim cmd As New SqlCommand("Actualiza_tiporeferencia", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_tiporeferencia", id_tiporeferencia)
        cmd.Parameters.AddWithValue("@Ptiporeferencia", tiporeferencia)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_tiporeferencia() As List(Of CTipoReferencia) Implements IService1.Obtener_tiporeferencia
        Dim Resultado As New List(Of CTipoReferencia)
        Dim cmd As New SqlCommand("Obtener_tiposReferencia", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CTipoReferencia
        While reader.Read
            Aux = New CTipoReferencia
            Aux.id_tiporeferencia = DirectCast(reader.Item("id_tiporeferencia"), Integer)
            Aux.tiporeferencia = DirectCast(reader.Item("tiporeferencia"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "Usuarios"
    Function Inserta_usuarios(ByVal nombre As String, ByVal apellidoPaterno As String, ByVal apellidoMaterno As String, ByVal Email As String, ByVal usuario As String, ByVal contraseña As String, ByVal TipoUsuario As Integer, ByVal fotografia As String, ByVal Usuario_Coordinador As Integer, ByVal Coordinador As String) As Integer Implements IService1.Inserta_usuarios
        Dim cmd As New SqlCommand("Inserta_usuario", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pnombre", nombre)
        cmd.Parameters.AddWithValue("@PapellidoPaterno", apellidoPaterno)
        cmd.Parameters.AddWithValue("@PapellidoMaterno", apellidoMaterno)
        cmd.Parameters.AddWithValue("@PEmail", Email)
        cmd.Parameters.AddWithValue("@Pusuario", usuario)
        cmd.Parameters.AddWithValue("@Pcontraseña", contraseña)
        cmd.Parameters.AddWithValue("@TipoUsuario", TipoUsuario)
        cmd.Parameters.AddWithValue("@Pfotografia", fotografia)
        cmd.Parameters.AddWithValue("@PUsuario_Coordinador", Usuario_Coordinador)
        cmd.Parameters.AddWithValue("@PCoordinador", Coordinador)

        Conexion.Close()
        Conexion.Open()

        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New Integer
        While reader.Read
            Aux = reader.Item(0)
        End While

        Conexion.Close()

        Return Aux
    End Function
    Public Function ObtenerAgentes_CallCenter(ByVal TipoUsuario As Integer) As String Implements IService1.ObtenerAgentes_CallCenter
        Try
            Dim cmd As New SqlCommand("Obtener_usuarios_Tipo", Conexion)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@PTipo", TipoUsuario)
            Conexion.Close()
            Conexion.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader
            Dim Aux As String = ""

            While reader.Read
                Aux = Aux & "'" & reader.Item("usuario") & "',"
            End While
            Conexion.Close()

            Aux = Aux.Substring(0, Aux.Length - 1)

            Return Aux
        Catch ex As Exception
            Throw New FaultException(ex.Message)
        End Try
    End Function

    Function Inserta_Supervisor()

    End Function

    Function Elimina_usuarios(ByVal id_usuario As Integer) As Boolean Implements IService1.Elimina_usuarios
        Dim cmd As New SqlCommand("Elimina_usuarios", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_usuario", id_usuario)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_usuarios(ByVal id_usuario As Integer, ByVal nombre As String, ByVal apellidoPaterno As String, ByVal apellidoMaterno As String, ByVal Email As String, ByVal activo As Integer) As Boolean Implements IService1.Actualiza_usuarios

        Dim cmd As New SqlCommand("Actualiza_usuarios", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_usuario", id_usuario)
        cmd.Parameters.AddWithValue("@Pnombre", nombre)
        cmd.Parameters.AddWithValue("@PapellidoPaterno", apellidoPaterno)
        cmd.Parameters.AddWithValue("@PapellidoMaterno", apellidoMaterno)
        cmd.Parameters.AddWithValue("@PEmail", Email)
        cmd.Parameters.AddWithValue("@activo", activo)
        'cmd.Parameters.AddWithValue("@Pusuario", usuario)
        'cmd.Parameters.AddWithValue("@Pcontraseña", contraseña)
        'cmd.Parameters.AddWithValue("@PfechaCreacion", fechaCreacion)
        'cmd.Parameters.AddWithValue("@Pfotografia", fotografia)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_usuariosPass(ByVal id_usuario As Integer, ByVal nombre As String, ByVal apellidoPaterno As String, ByVal apellidoMaterno As String, ByVal Email As String, ByVal usuario As String, ByVal contraseña As String,
                                    ByVal activo As Integer, ByVal TipoUsuario As Integer, ByVal id_Supervisor As Integer) As Boolean Implements IService1.Actualiza_usuariosPass

        Dim cmd As New SqlCommand("Actualiza_usuariosPass", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_usuario", id_usuario)
        cmd.Parameters.AddWithValue("@Pnombre", nombre)
        cmd.Parameters.AddWithValue("@PapellidoPaterno", apellidoPaterno)
        cmd.Parameters.AddWithValue("@PapellidoMaterno", apellidoMaterno)
        cmd.Parameters.AddWithValue("@PEmail", Email)
        cmd.Parameters.AddWithValue("@activo", activo)
        cmd.Parameters.AddWithValue("@Pusuario", usuario)
        cmd.Parameters.AddWithValue("@Pcontraseña", contraseña)
        cmd.Parameters.AddWithValue("@PTipoUsuario", TipoUsuario)
        cmd.Parameters.AddWithValue("@PIdSupervisor", id_Supervisor)
        'cmd.Parameters.AddWithValue("@Pfotografia", fotografia)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function

    Function Obtener_usuarios_todos() As List(Of CUsuarios) Implements IService1.Obtener_usuarios_todos
        Dim Resultado As New List(Of CUsuarios)
        Dim cmd As New SqlCommand("Obtener_usuarios_todos", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CUsuarios
        While reader.Read
            Aux = New CUsuarios
            Aux.id_usuario = DirectCast(reader.Item("id_usuario"), Integer)
            Aux.nombre = DirectCast(reader.Item("nombre"), String)
            Aux.apellidoPaterno = DirectCast(reader.Item("apellidoPaterno"), String)
            Aux.apellidoMaterno = DirectCast(reader.Item("apellidoMaterno"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Aux.usuario = DirectCast(reader.Item("usuario"), String)
            Aux.contraseña = DirectCast(reader.Item("contraseña"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.fotografia = DirectCast(reader.Item("fotografia"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_usuarios_detalles(ByVal id_usuario As Integer) As CUsuarios Implements IService1.Obtener_usuarios_detalles

        Dim cmd As New SqlCommand("Obtener_usuario_detalles", Conexion)
        cmd.Parameters.AddWithValue("@idUsuario", id_usuario)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New CUsuarios
        While reader.Read

            Aux.id_usuario = DirectCast(reader.Item("id_usuario"), Integer)
            Aux.nombre = DirectCast(reader.Item("nombre"), String)
            Aux.apellidoPaterno = DirectCast(reader.Item("apellidoPaterno"), String)
            Aux.apellidoMaterno = DirectCast(reader.Item("apellidoMaterno"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Aux.usuario = DirectCast(reader.Item("usuario"), String)
            Aux.contraseña = DirectCast(reader.Item("contraseña"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.fotografia = DirectCast(reader.Item("fotografia"), String)
        End While
        Conexion.Close()
        Return Aux
    End Function
    Function Actualiza_contraseaUsuario(ByVal id_usuario As Integer, ByVal Contraseña As String) As Boolean Implements IService1.Actualiza_contraseaUsuario

        Dim cmd As New SqlCommand("Actualiza_contraseña", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idUsuario", id_usuario)
        cmd.Parameters.AddWithValue("@Contraseña", Contraseña)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function VerificaUsuario(ByVal Usuario As String) As Boolean Implements IService1.VerificaUsuario
        Dim Resultado As Boolean = False
        Dim cmd As New SqlCommand("VerificaUsuario", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Usuario", Usuario)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New Cids
        While reader.Read
            Aux = New Cids
            Aux.id_supervisor = DirectCast(reader.Item("id_supervisor"), Integer)
            Aux.id_usuario = DirectCast(reader.Item("id_usuario"), Integer)
            Aux.id_usuario1 = DirectCast(reader.Item("id_usuario1"), Integer)
        End While

        If Aux.id_supervisor > 0 Or Aux.id_usuario > 0 Or Aux.id_usuario1 > 0 Then
            Resultado = True
        Else
            Resultado = False
        End If

        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "Estados"
    Function Obtener_estados() As List(Of CEstados) Implements IService1.Obtener_estados
        Dim Resultado As New List(Of CEstados)
        Dim cmd As New SqlCommand("Obtener_estados", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CEstados
        While reader.Read
            Aux = New CEstados
            Aux.id = DirectCast(reader.Item("id"), Integer)
            Aux.nombre = DirectCast(reader.Item("nombre"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_ciudad(ByVal id_estado As Integer) As List(Of CCiudades) Implements IService1.Obtener_ciudad
        Dim Resultado As New List(Of CCiudades)
        Dim cmd As New SqlCommand("Obtener_ciudades_idestado", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idEstado", id_estado)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CCiudades
        While reader.Read
            Aux = New CCiudades
            Aux.id = DirectCast(reader.Item("id"), Integer)
            Aux.nombre = DirectCast(reader.Item("nombre"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "Avance etapas"
    Function Valida_EtapaCliente(ByVal IdCliente As Integer, ByVal IdEtapa As Integer, ByVal IdUsuario As Integer, ByVal IdProducto As Integer) As Boolean Implements IService1.Valida_EtapaCliente
        Dim cmd As New SqlCommand("ValidaEtapaCliente", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@IdCliente", IdCliente)
        cmd.Parameters.AddWithValue("@IdEtapa", IdEtapa)
        cmd.Parameters.AddWithValue("@IdUsuario", IdUsuario)
        cmd.Parameters.AddWithValue("@IdProducto", IdProducto)
        Conexion.Close()

        Try
            Conexion.Open()
            If cmd.ExecuteScalar = 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        Finally
            Conexion.Close()
        End Try
    End Function
    Function Avanza_EtapaCliente(ByVal id_cliente As Integer, ByVal id_usuario As Integer, ByVal id_etapa As Integer, ByVal Observaciones As String, ByVal id_productoRegistro As Integer) As Boolean Implements IService1.Avanza_EtapaCliente

        Dim cmd As New SqlCommand("Avanza_EtapaCliente", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idCliente", id_cliente)
        cmd.Parameters.AddWithValue("@idUsuario", id_usuario)
        cmd.Parameters.AddWithValue("@idEtapa", id_etapa)
        cmd.Parameters.AddWithValue("@Observaciones", Observaciones)
        cmd.Parameters.AddWithValue("@idProducto", id_productoRegistro)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_etapasClienteDetalles(ByVal id_cliente As Integer) As List(Of CEtapasDetalles) Implements IService1.Obtener_etapasClienteDetalles
        Dim Resultado As New List(Of CEtapasDetalles)
        Dim cmd As New SqlCommand("Obtener_etapasClienteDetalle", Conexion)
        cmd.Parameters.AddWithValue("@idCliente", id_cliente)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CEtapasDetalles
        While reader.Read
            Aux = New CEtapasDetalles
            Aux.id_operacion = DirectCast(reader.Item("id_operacion"), Integer)
            Aux.FechaInicio = DirectCast(reader.Item("FechaInicio"), Date)
            Aux.Observaciones = DirectCast(reader.Item("Observaciones"), String)
            Aux.usuario = DirectCast(reader.Item("usuario"), String)
            Aux.nombre = DirectCast(reader.Item("nombre"), String)
            Aux.Descripcion = DirectCast(reader.Item("Descripcion"), String)
            Aux.NombreCorto = DirectCast(reader.Item("NombreCorto"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
#End Region
#Region "Tareas"
    Function Inserta_tareas(ByVal descripcion As String, ByVal id_prioridad As Integer, ByVal id_usuario As Integer, ByVal avisado As String, ByVal fechaCreacion As Date, ByVal fechaProgramada As Date, ByVal HoraProgramada As String) As Boolean Implements IService1.Inserta_tareas

        Dim cmd As New SqlCommand("Insertar_tarea", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pdescripcion", descripcion)
        cmd.Parameters.AddWithValue("@Pid_prioridad", id_prioridad)
        cmd.Parameters.AddWithValue("@Pid_usuario", id_usuario)
        cmd.Parameters.AddWithValue("@Pavisado", avisado)
        cmd.Parameters.AddWithValue("@PfechaCreacion", fechaCreacion)
        cmd.Parameters.AddWithValue("@PfechaProgramada", fechaProgramada)
        cmd.Parameters.AddWithValue("@PHoraProgramada", HoraProgramada)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_tareas(ByVal id_tarea As Integer, ByVal descripcion As String, ByVal id_prioridad As Integer, ByVal id_usuario As Integer, ByVal avisado As String, ByVal fechaCreacion As Date, ByVal fechaProgramada As Date, ByVal HoraProgramada As String) As Boolean Implements IService1.Actualiza_tareas

        Dim cmd As New SqlCommand("Actualiza_tarea", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_tarea", id_tarea)
        cmd.Parameters.AddWithValue("@Pdescripcion", descripcion)
        cmd.Parameters.AddWithValue("@Pid_prioridad", id_prioridad)
        cmd.Parameters.AddWithValue("@Pid_usuario", id_usuario)
        cmd.Parameters.AddWithValue("@Pavisado", avisado)
        cmd.Parameters.AddWithValue("@PfechaCreacion", fechaCreacion)
        cmd.Parameters.AddWithValue("@PfechaProgramada", fechaProgramada)
        cmd.Parameters.AddWithValue("@PHoraProgramada", HoraProgramada)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Elimina_tareas(ByVal id_tarea As Integer) As Boolean Implements IService1.Elimina_tareas

        Dim cmd As New SqlCommand("Elimina_tarea", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Pid_tarea", id_tarea)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_tareas_prioridad() As List(Of CTareasPrioridad) Implements IService1.Obtener_tareas_prioridad
        Dim Resultado As New List(Of CTareasPrioridad)
        Dim cmd As New SqlCommand("Obtener_tareasPrioridad", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CTareasPrioridad
        While reader.Read
            Aux = New CTareasPrioridad
            Aux.id_prioridad = DirectCast(reader.Item("id_prioridad"), Integer)
            Aux.prioridad = DirectCast(reader.Item("prioridad"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_tareasPendientesUsuario(ByVal id_usuario As Integer) As List(Of CTareasPendientes) Implements IService1.Obtener_tareasPendientesUsuario
        Dim Resultado As New List(Of CTareasPendientes)
        Dim cmd As New SqlCommand("Obtener_tareasPendientesUsuario", Conexion)
        cmd.Parameters.AddWithValue("@id_usuario", id_usuario)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CTareasPendientes
        While reader.Read
            Aux = New CTareasPendientes
            Aux.id_tarea = DirectCast(reader.Item("id_tarea"), Integer)
            Aux.descripcion = DirectCast(reader.Item("descripcion"), String)
            Aux.Prioridad = DirectCast(reader.Item("Prioridad"), String)
            Aux.Avisado = DirectCast(reader.Item("Avisado"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.fechaProgramada = DirectCast(reader.Item("fechaProgramada"), Date)
            Aux.HoraProgramada = reader.Item("HoraProgramada")
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_tareasTerminadasUsuario(ByVal id_usuario As Integer) As List(Of CTareasPendientes) Implements IService1.Obtener_tareasTerminadasUsuario
        Dim Resultado As New List(Of CTareasPendientes)
        Dim cmd As New SqlCommand("Obtener_tareasTerminadasUsuario", Conexion)
        cmd.Parameters.AddWithValue("@id_usuario", id_usuario)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CTareasPendientes
        While reader.Read
            Aux = New CTareasPendientes
            Aux.id_tarea = DirectCast(reader.Item("id_tarea"), Integer)
            Aux.descripcion = DirectCast(reader.Item("descripcion"), String)
            Aux.Prioridad = DirectCast(reader.Item("Prioridad"), String)
            Aux.Avisado = DirectCast(reader.Item("Avisado"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.fechaProgramada = DirectCast(reader.Item("fechaProgramada"), Date)
            Aux.HoraProgramada = reader.Item("HoraProgramada")
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function TerminarTarea(ByVal id_tarea As Integer) As Boolean Implements IService1.TerminarTarea

        Dim cmd As New SqlCommand("TerminarTarea", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@IDTarea", id_tarea)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
#End Region
#Region "Notificaciones"
    Function Obtener_notificaciones(ByVal idUsuario As Integer) As List(Of CNotifica) Implements IService1.Obtener_notificaciones
        Dim Resultado As New List(Of CNotifica)

        Resultado.AddRange(LlamadasPendientes(idUsuario))
        Resultado.AddRange(citasPendientes(idUsuario))
        Resultado.AddRange(TareasPendientes(idUsuario))
        Return Resultado
    End Function
    Function LlamadasPendientes(ByVal idUsuario As Integer) As List(Of CNotifica)
        Dim Resultado As New List(Of CNotifica)
        Dim Llamadas = Obtener_llamdas_porAvisar(idUsuario)
        Dim Aux As CNotifica
        For I = 0 To Llamadas.Count - 1
            Aux = New CNotifica
            ' If Enviar_Correo(Llamadas(I).EmailUsuario, Crea_HTML(Llamadas(I)), "Aviso llamada") Then
            'Update a llamada
            Aux.TituloNotificacion = "Llamada"
            Aux.DescripcionNotificacion = "Recordatorio de llamada para: " + Llamadas(I).Nombre + " " + Llamadas(I).ApellidoPaterno + " Observación: " + Llamadas(I).ObservacionUsuario
            Aux.URL = "/Usuario/NuevaLlamada.aspx"
            Resultado.Add(Aux)
            Actualiza_llamada(Llamadas(I).id_llamada)
            ' End If
        Next
        Return Resultado
    End Function
    Function citasPendientes(ByVal idUsuario As Integer) As List(Of CNotifica)
        Dim Resultado As New List(Of CNotifica)
        Dim Citas = Obtener_citasPendientesAvisar(idUsuario)
        Dim Aux As CNotifica
        For I = 0 To Citas.Count - 1
            Aux = New CNotifica

            Aux.TituloNotificacion = "Citas"
            Aux.DescripcionNotificacion = "Recordatorio de Cita para: " + Citas(I).nombre + " " + Citas(I).apellidoPaterno + " Observación: " + Citas(I).ObservacionUsuario
            Aux.URL = "/Usuario/NuevaCita.aspx"
            Resultado.Add(Aux)
            ' If Enviar_Correo(Citas(I).Email, Crea_HTMLCorreo(crea_mensaje(Citas(I))), "Recordatorio de cita") Then
            Aviso_citaUsuario(Citas(I).id_cita)
            '  End If
        Next
        Return Resultado
    End Function
    Function TareasPendientes(ByVal idUsuario As Integer) As List(Of CNotifica)
        Dim Resultado As New List(Of CNotifica)
        Dim Tareas = Obtener_tareasPorAvisar(idUsuario)
        Dim Aux As CNotifica
        For I = 0 To Tareas.Count - 1
            Aux = New CNotifica
            Aux.TituloNotificacion = "Tareas"
            Aux.DescripcionNotificacion = " Observación: " + Tareas(I).descripcion
            Aux.URL = "/Usuario/tareas.aspx"
            Resultado.Add(Aux)
            ' If Enviar_Correo(Tareas(I).Email, Crea_HTMLCorreoTarea(CreaMensaje(Tareas(I))), "Aviso de tarea Pendiente") Then
            Avisado_tarea(Tareas(I).id_tarea)
            '   End If
        Next
        Return Resultado
    End Function
    Function Obtener_tareasPorAvisar(ByVal idUsuario As Integer) As List(Of CTareasPorAvisar)
        Dim Resultado As New List(Of CTareasPorAvisar)
        Dim cmd As New SqlCommand("Obtener_tareasPorAvisar_usuario", Conexion)
        cmd.Parameters.AddWithValue("@idUsuario", idUsuario)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CTareasPorAvisar
        While reader.Read
            Aux = New CTareasPorAvisar
            Aux.id_tarea = DirectCast(reader.Item("id_tarea"), Integer)
            Aux.prioridad = DirectCast(reader.Item("prioridad"), String)
            Aux.descripcion = DirectCast(reader.Item("descripcion"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.fechaProgramada = DirectCast(reader.Item("fechaProgramada"), Date)
            Aux.HoraProgramada = reader.Item("HoraProgramada")
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Avisado_tarea(ByVal id_tarea As Integer) As Boolean

        Dim cmd As New SqlCommand("Avisado_tarea", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@IDTarea", id_tarea)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Actualiza_llamada(ByVal id_llamada As Integer) As Boolean

        Dim cmd As New SqlCommand("Avisado_llamadaUsuario", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idLlamada", id_llamada)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
    Function Obtener_llamdas_porAvisar(ByVal id_usuario As Integer) As List(Of CLlamadasPAvisar)
        Dim Resultado As New List(Of CLlamadasPAvisar)
        Dim cmd As New SqlCommand("Obtener_llamdas_porAvisar_usuario", Conexion)
        cmd.Parameters.AddWithValue("@idCliente", id_usuario)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CLlamadasPAvisar
        While reader.Read
            Aux = New CLlamadasPAvisar
            Aux.Nombre = DirectCast(reader.Item("Nombre"), String)
            Aux.ApellidoPaterno = DirectCast(reader.Item("ApellidoPaterno"), String)
            Aux.ApellidoMaterno = DirectCast(reader.Item("ApellidoMaterno"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Aux.ObservacionUsuario = DirectCast(reader.Item("ObservacionUsuario"), String)
            Aux.Telefono = DirectCast(reader.Item("Telefono"), String)
            Aux.Hora = reader.Item("HoraProgramacion").ToString
            Aux.EmailUsuario = DirectCast(reader.Item("EmailUsuario"), String)
            Aux.id_llamada = DirectCast(reader.Item("id_llamada"), Integer)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_citasPendientesAvisar(ByVal idUsuario As Integer) As List(Of CCitasPendientes)
        Dim Resultado As New List(Of CCitasPendientes)

        Dim cmd As New SqlCommand("Obtener_citasPendientesAvisar_usuario", Conexion)

        cmd.Parameters.AddWithValue("@idUsuario", idUsuario)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CCitasPendientes
        While reader.Read
            Aux = New CCitasPendientes
            Aux.id_cita = DirectCast(reader.Item("id_cita"), Integer)
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.id_usuario = DirectCast(reader.Item("id_usuario"), Integer)
            Aux.Fecha = DirectCast(reader.Item("Fecha"), Date)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.HoraProgramacion = reader.Item("HoraProgramacion")
            Aux.Programada = reader.Item("Programada")
            Aux.AvisoCliente = reader.Item("AvisoCliente")
            Aux.AvisoUsuario = reader.Item("AvisoUsuario")
            Aux.realizada = reader.Item("realizada")
            Aux.ObservacionUsuario = DirectCast(reader.Item("ObservacionUsuario"), String)
            Aux.ObservacionCliente = DirectCast(reader.Item("ObservacionCliente"), String)
            Aux.HoraTermino = reader.Item("HoraTermino")
            Aux.Lugar = DirectCast(reader.Item("Lugar"), String)
            Aux.ConfimacionCliente = reader.Item("ConfimacionCliente")
            Aux.nombre = DirectCast(reader.Item("nombre"), String)
            Aux.apellidoPaterno = DirectCast(reader.Item("apellidoPaterno"), String)
            Aux.apellidoMaterno = DirectCast(reader.Item("apellidoMaterno"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)
            Aux.Nombre1 = DirectCast(reader.Item("NombreUsuario"), String)
            Aux.ApellidoPaterno1 = DirectCast(reader.Item("ApUsuario"), String)
            Aux.ApellidoMaterno1 = DirectCast(reader.Item("Ap2Usuario"), String)
            Aux.Telefono = DirectCast(reader.Item("Telefono"), String)
            Aux.Principal = reader.Item("Principal")
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Aviso_citaUsuario(ByVal id_cita As Integer) As Boolean

        Dim cmd As New SqlCommand("Aviso_citaUsuario", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@id_cita", id_cita)
        Conexion.Close()
        Try
            Conexion.Open()
            If cmd.ExecuteNonQuery() > 0 Then
                Conexion.Close()
                Return True
            End If
        Catch ex As Exception
            Conexion.Close()
            Return False
        End Try
        Conexion.Close()
        Return False
    End Function
#End Region
#Region "Emails"
    Function CreaHTMLLlamadaCalif(ByVal idLlamada As Integer) As String
        Dim HTML As String = ""
        HTML += "
        <!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">
<html xmlns = ""http://www.w3.org/1999/xhtml"" xmlns:v = ""urn:schemas-Microsoft - com: vml"" xmlns: o = ""urn:schemas-Microsoft - com: office : office"" >
        <head>
            <meta http-equiv="" Content-Type"" content="" text/html; charset=UTF-8""/>
            <meta name="" viewport"" content="" initial-scale=1.0""/>
            <meta http-equiv="" X-UA-Compatible"" content="" IE=edge""/>
            <meta name="" format-detection"" content="" telephone=no""/>
            <title>CRM Mail</title>
            <link href="" http://fonts.googleapis.com/css?family=Roboto:400,300,700&subset=latin,cyrillic,greek"" rel=""stylesheet"" type=""text/css"">
                <!--[if mso]>
<style type=""text/css"">
td,span,a{font-family: Arial, sans-serif !important;}
a{text-decoration: none !important;}
</style>
<![endif]-->
                <style type="" text/css"">

  
    /* Resets: see reset.css for details */
    .ReadMsgBody { width: 100%; background-color: #ffffff;}
    .ExternalClass {width: 100%; background-color: #ffffff;}
    .ExternalClass, .ExternalClass p, .ExternalClass span,
    .ExternalClass font, .ExternalClass td, .ExternalClass div {line-height:100%;}
    #outlook a{ padding:0;}
    html{width: 100%; }
    body {-webkit-text-size-adjust:none; -ms-text-size-adjust:none; }
    html,body {background-color: #ffffff; margin: 0; padding: 0; }
    table {border-spacing:0;}
    table td {border-collapse:collapse;}
    br, strong br, b br, em br, i br { line-height:100%; }
    h1, h2, h3, h4, h5, h6 { line-height: 100% !important; -webkit-font-smoothing: antialiased; }
    img{height: auto !important; line-height: 100%; outline: none; text-decoration: none; display:block !important; }
    span a { text-decoration: none !important;}
    a{ text-decoration: none !important; }
    table p{margin:0;}
    .yshortcuts, .yshortcuts a, .yshortcuts a:link,.yshortcuts a:visited,
    .yshortcuts a:hover, .yshortcuts a span { text-decoration: none !important; border-bottom: none !important;}
    table{ mso-table-lspace:0pt; mso-table-rspace:0pt; }
    img{ -ms-interpolation-mode:bicubic; }
    ul{list-style: initial; margin:0; padding-left:20px;}
    /*mailChimp class*/
    .default-edit-image{
    height:20px;
    }
    .tpl-repeatblock {
    padding: 0px !important;
    border: 1px dotted rgba(0,0,0,0.2);
    }
    img{height:auto !important;}
    td[class=""image-270px""] img{
    width:270px;
    height:auto !important;
    max-width:270px !important;
    }
    td[class=""image-170px""] img{
    width:170px;
    height:auto !important;
    max-width:170px !important;
    }
    td[class=""image-185px""] img{
    width:185px;
    height:auto !important;
    max-width:185px !important;
    }
    td[class=""image-124px""] img{
    width:124px;
    height:auto !important;
    max-width:124px !important;
    }
    @media only screen and (max-width: 640px){
    body{
    width:auto!important;
    }
    table[class=""container""]{
    width: 100%!important;
    padding-left: 20px!important;
    padding-right: 20px!important;
    min-width:100% !important;
    }
    td[class=""image-270px""] img{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    }
    td[class=""image-170px""] img{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    }
    td[class=""image-185px""] img{
    width:185px !important;
    height:auto !important;
    max-width:185px !important;
    }
    td[class=""image-124px""] img{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    }
    td[class=""image-100-percent""] img{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    }
    td[class=""small-image-100-percent""] img{
    width:100% !important;
    height:auto !important;
    }
    table[class=""full-width""]{
    width:100% !important;
    min-width:100% !important;
    }
    table[class=""full-width-text""]{
    width:100% !important;
    background-color:#ffffff;
    padding-left:20px !important;
    padding-right:20px !important;
    }
    table[class=""full-width-text2""]{
    width:100% !important;
    background-color:#ffffff;
    padding-left:20px !important;
    padding-right:20px !important;
    }
    table[class=""col-2-3img""]{
    width:50% !important;
    margin-right: 20px !important;
    }
    table[class=""col-2-3img-last""]{
    width:50% !important;
    }
    table[class=""col-2-footer""]{
    width:55% !important;
    margin-right:20px !important;
    }
    table[class=""col-2-footer-last""]{
    width:40% !important;
    }
    table[class=""col-2""]{
    width:47% !important;
    margin-right:20px !important;
    }
    table[class=""col-2-last""]{
    width:47% !important;
    }
    table[class=""col-3""]{
    width:29% !important;
    margin-right:20px !important;
    }
    table[class=""col-3-last""]{
    width:29% !important;
    }
    table[class=""row-2""]{
    width:50% !important;
    }
    td[class=""text-center""]{
    text-align: center !important;
    }
    /* start clear and remove*/
    table[class=""remove""]{
    display:none !important;
    }
    td[class=""remove""]{
    display:none !important;
    }
    /* end clear and remove*/
    table[class=""fix-box""]{
    padding-left:20px !important;
    padding-right:20px !important;
    }
    td[class=""fix-box""]{
    padding-left:20px !important;
    padding-right:20px !important;
    }
    td[class=""font-resize""]{
    font-size: 18px !important;
    line-height: 22px !important;
    }
    table[class=""space-scale""]{
    width:100% !important;
    float:none !important;
    }
    table[class=""clear-align-640""]{
    float:none !important;
    }
     table[class=""show-full-mobile""]{
        display:none !important;
        width:100% !important;
        min-width:100% !important;
    }
    }
    @media only screen and (max-width: 479px){
    body{
    font-size:10px !important;
    }
    table[class=""container""]{
    width: 100%!important;
    padding-left: 10px!important;
    padding-right:10px!important;
    min-width:100% !important;
    }
    table[class=""container2""]{
    width: 100%!important;
    float:none !important;
    min-width:100% !important;
    }
    td[class=""full-width""] img{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    min-width:124px !important;
    min-width:100% !important;
    }
    td[class=""image-270px""] img{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    min-width:124px !important;
    }
    td[class=""image-170px""] img{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    min-width:124px !important;
    }
    td[class=""image-185px""] img{
    width:185px !important;
    height:auto !important;
    max-width:185px !important;
    min-width:124px !important;
    }
    td[class=""image-124px""] img{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    min-width:124px !important;
    }
    td[class=""image-100-percent""] img{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    min-width:124px !important;
    }
    td[class=""small-image-100-percent""] img{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    min-width:124px !important;
    }
    table[class=""full-width""]{
    width:100% !important;
    }
    table[class=""full-width-text""]{
    width:100% !important;
    background-color:#ffffff;
    padding-left:20px !important;
    padding-right:20px !important;
    }
    table[class=""full-width-text2""]{
    width:100% !important;
    background-color:#ffffff;
    padding-left:20px !important;
    padding-right:20px !important;
    }
    table[class=""col-2-footer""]{
    width:100% !important;
    margin-right:0px !important;
    }
    table[class=""col-2-footer-last""]{
    width:100% !important;
    }
    table[class=""col-2""]{
    width:100% !important;
    margin-right:0px !important;
    }
    table[class=""col-2-last""]{
    width:100% !important;
    }
    table[class=""col-3""]{
    width:100% !important;
    margin-right:0px !important;
    }
    table[class=""col-3-last""]{
    width:100% !important;
    }
    table[class=""row-2""]{
    width:100% !important;
    }
    table[id=""col-underline""]{
    float: none !important;
    width: 100% !important;
    border-bottom: 1px solid #eee;
    }
    td[id=""col-underline""]{
    float: none !important;
    width: 100% !important;
    border-bottom: 1px solid #eee;
    }
    td[class=""col-underline""]{
    float: none !important;
    width: 100% !important;
    border-bottom: 1px solid #eee;
    }
    /*start text center*/
    td[class=""text-center""]{
    text-align: center !important;
    }
    div[class=""text-center""]{
    text-align: center !important;
    }
    /*end text center*/
    /* start  clear and remove */
    table[id=""clear-padding""]{
    padding:0 !important;
    }
    td[id=""clear-padding""]{
    padding:0 !important;
    }
    td[class=""clear-padding""]{
    padding:0 !important;
    }
    table[class=""remove-479""]{
    display:none !important;
    }
    td[class=""remove-479""]{
    display:none !important;
    }
    table[class=""clear-align""]{
    float:none !important;
    }
    /* end  clear and remove */
    table[class=""width-small""]{
    width:100% !important;
    }
    table[class=""fix-box""]{
    padding-left:15px !important;
    padding-right:15px !important;
    }
    td[class=""fix-box""]{
    padding-left:15px !important;
    padding-right:15px !important;
    }
    td[class=""font-resize""]{
    font-size: 14px !important;
    }
    td[class=""increase-Height""]{
    height:10px !important;
    }
    td[class=""increase-Height-20""]{
    height:20px !important;
    }
    table[width=""595""]{
        width:100% !important;
      }

      table[class=""show-full-mobile""]{
        display:table !important;
        width:100% !important;
        min-width:100% !important;
      }
    }
    @media only screen and (max-width: 320px){
    table[class=""width-small""]{
    width:125px !important;
    }
    img[class=""image-100-percent""]{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    min-width:124px !important;
    }
    }
    td ul{list-style: initial; margin:0; padding-left:20px;}

	@media only screen and (max-width: 640px){ .image-100-percent{ width:100%!important; height: auto !important; max-width: 100% !important; min-width: 124px !important;}}body{background-color:#efefef;} .default-edit-image{height:20px;} tr.tpl-repeatblock , tr.tpl-repeatblock > td{ display:block !important;} .tpl-repeatblock {padding: 0px !important;border: 1px dotted rgba(0,0,0,0.2);} table[width=""595""]{width:100% !important;}</style>
                <!--[if gte mso 9]><xml><o:OfficeDocumentSettings><o:AllowPNG/><o:PixelsPerInch>96</o:PixelsPerInch></o:OfficeDocumentSettings></xml><![endif]-->
        </head>
<body  style = ""font-size:12px; width:100%; height:100%;"" >
<table id = ""mainStructure"" width=""800"" Class=""full-width"" align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"" style=""background-color:#efefef; width:800px; max-width: 800px; margin: 0 auto; outline: 1px solid #efefef; box-shadow: 0px 0px 5px #E0E0E0;"">
    <!--START VIEW ONLINE And ICON SOCAIL -->
    <tbody><tr>
        <td align =""center"" valign=""top"" style=""background-color: #3d4db8;"">
            <!-- start container 600 -->
            <table width = ""600"" align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"" Class=""container"" style=""padding-left: 20px; padding-right: 20px; min-width: 600px; width: 600px; background-color: #3d4db8;mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                <tbody><tr>
                    <td valign = ""top"">
                        <table width = ""560"" align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"" Class=""full-width"" style=""width: 560px; background-color: #3d4db8;mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                            <!-- start space -->
                            <tbody><tr>
                                <td valign = ""top"" height=""10"" style=""border-collapse: collapse; height: 10px; line-height: 10px; font-size: 10px;"">
                                </td>
                                </tr>
                                <!-- end space -->
                                <tr>
                                    <td valign = ""top"">
                                        <!-- start container -->
                                        <table width = ""100%"" align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"" style=""mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                                            <tbody><tr>
                                                <td valign = ""top"">

                                                    <table align = ""left"" border=""0"" cellspacing=""0"" cellpadding=""0"" Class=""container2"" width=""auto"" style=""mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                                                        <tbody><tr>
                                                            <td>

                                                            </td>
                                                            </tr>
                                                            <!-- start space -->
                                                            <tr>
                                                                <td valign = ""top"" Class=""increase-Height"">
                                                                </td>
                                                            </tr>
                                                            <!-- end space -->
                                                        </tbody></table><!--[if (gte mso 9)|(IE)]></td><td valign=""top"" ><![endif]-->




                                                </td>
                                                </tr>
                                            </tbody></table>
                                        <!-- end container  -->
                                    </td>
                                </tr>
                                <!-- start space -->
                                <tr>
                                    <td valign = ""top"" height=""10"" style=""border-collapse: collapse; height: 10px; line-height: 10px; font-size: 10px;"">
                                    </td>
                                </tr>
                                <!-- end space -->
                                <!-- start space -->
                                <tr>
                                    <td valign = ""top"" Class=""increase-Height"">
                                    </td>
                                </tr>
                                <!-- end space -->
                            </tbody></table>
                        <!-- end container 600-->
                    </td>
                    </tr>
                </tbody></table>
        </td>
        </tr>
        <!--END VIEW ONLINE And ICON SOCAIL-->
    </tbody>
    <!-- START LAYOUT 2-->
    <tbody><tr>
        <td align =""center"" valign=""top"" style=""background-color: #ebebec;"">
            <!-- start layout-2 container width 600px -->
            <table width = ""600"" align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"" Class=""full-width"" style=""padding-left: 20px; padding-right: 20px; min-width: 600px; width: 600px; background-color: #ffffff;mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                <tbody><tr>
                    <td valign = ""top"">
                        <!-- start layout-2 container width 600px -->
                        <table width = ""560"" align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"" Class=""full-width"" style=""width: 560px;mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                            <!-- start image And content -->
                            <tbody><tr>
                                <td valign = ""top"" width=""100%"">

                                    <table width = ""270"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""left"" Class=""full-width"" style=""width: 270px;mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                                        <tbody><tr>
                                            <td valign = ""bottom"" align=""center"" width=""270"" style=""width: 270px;"">
                                                <a href = ""#"" style=""font-size: inherit; border-style: none; text-decoration: none!important;"" border=""0"">
                                                    <img src = ""http://altaircloud.mx/logo-default.png"" width=""200"" alt=""image1"" style=""max-width: 270px; display: block;"" border=""0"" hspace=""0"" vspace=""0"">
                              </a>
                                            </td>
                                            </tr>
                                        </tbody></table><!--[if (gte mso 9)|(IE)]></td><td valign=""top"" ><![endif]-->


                                    <Table Class=""remove"" width=""1"" border=""0"" cellpadding=""0"" cellspacing=""0"" align=""left"" style=""font-size: 0px; line-height: 0; border-collapse: collapse; width: 1px;mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                                        <tbody> <tr>
        <td width = ""0"" height=""2"" style=""border-collapse: collapse; width: 0px; height: 2px; line-height: 2px; font-size: 2px;"">
                                                <p style = ""padding-left: 20px;"">&nbsp;</p>
                                            </td>
                                            </tr>
                                        </tbody></table><!--[if (gte mso 9)|(IE)]></td><td valign=""top"" ><![endif]-->

                                    <Table width = ""270"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""right"" Class=""container"" style=""width: 270px;mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                                        <!--start space height -->
                                        <tbody> <tr>
        <td height = ""20"" style=""border-collapse: collapse; height: 20px; line-height: 20px; font-size: 20px;""></td>
                                            </tr>
                                            <!--end space height -->
                                            <!--start space height -->
                                            <tr>
        <td height = ""40"" Class=""remove"" style=""border-collapse: collapse; height: 40px; line-height: 40px; font-size: 40px;""></td>
                                            </tr>
                                            <!--end space height -->
                                            <!-- start text content -->
                                            <tr>
        <td valign = ""top"">
                                                    <Table width = ""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""left"" style=""mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                                                        <tbody> <tr>
        <td style = ""font-size: 22px; font-family: Roboto, Arial, Helvetica, sans-serif; color: rgb(85, 85, 85); font-weight: 300; text-align: left; word-break: break-word; line-height: 30px;""><span style=""color: rgb(85, 85, 85); font-size: inherit; font-weight: 300; line-height: 30px;""><a href=""#"" style=""color: rgb(85, 85, 85); font-size: inherit; border-style: none; text-decoration: none !important; line-height: 30px;"" data-mce-href=""#"" border=""0"">&nbsp;<span style=""color: rgb(0, 0, 255); font-size: 22px; font-weight: 300; line-height: 30px;"">HOLA!</span><br style=""line-height: 24px;"">R</a>ecientemente un asesor nuestro realizo una llamada contigo, por favor dinos como fue esa llamada<br style=""line-height: 24px;""><br style=""line-height: 24px;""></span></td>
                                                            </tr>
                                                            <!--start space height -->
                                                            <tr>
        <td height = ""15"" style=""border-collapse: collapse; height: 15px; line-height: 15px; font-size: 15px;""></td>
                                                            </tr>
                                                            <!--end space height -->
                                                            <tr>
        <td style = ""font-size: 13px; font-family: Roboto, Arial, Helvetica, sans-serif; color: rgb(163, 162, 162); font-weight: 300; text-align: left; word-break: break-word; line-height: 21px;""><span style=""font-size: inherit; line-height: 21px;""><p style=""line-height: 24px;"">Escoge una opci&oacute;n:</p><p style=""line-height: 24px;""></p><p style=""line-height: 24px;""></p><p style=""line-height: 24px;"">Buena experiencia<a href=""http://crm.altaircloud.mx/Todos/calificaLlamada.aspx?idLlam=" + idLlamada.ToString + "&Calif=1""><img src=""http://altaircloud.mx/mail/img/Like.png"" width=""200""/></a></p></span></td>
                                                                <td style = ""font-size: 13px; font-family: Roboto, Arial, Helvetica, sans-serif; color: rgb(163, 162, 162); font-weight: 300; text-align: left; word-break: break-word; line-height: 21px;""><span style=""font-size: inherit; line-height: 21px;""><p style=""line-height: 24px;""></p><p style=""line-height: 24px;""></p><p style=""line-height: 24px;""></p><p style=""line-height: 24px;"">Mala experiencia<a href=""http://crm.altaircloud.mx/Todos/calificaLlamada.aspx?idLlam=" + idLlamada.ToString + "&Calif=2""><img src=""http://altaircloud.mx/mail/img/dislike.png"" width=""200""/></a></p></span></td>
                                                            </tr>
                                                            <!--start space height -->
                                                            <tr>
        <td height = ""15"" style=""border-collapse: collapse; height: 15px; line-height: 15px; font-size: 15px;""></td>
                                                            </tr>
                                                            <!--end space height -->
                                                            <tr>
        <td valign = ""top"" width=""auto"">
                                                                    <!-- start button -->
                                                                    <Table border = ""0"" align=""left"" cellpadding=""0"" cellspacing=""0"" width=""auto"" style=""mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                                                                        <tbody> <tr>
        <td valign = ""top"">
                                                                                <Table border = ""0"" align=""left"" cellpadding=""0"" cellspacing=""0"" dup=""0"" width=""auto"" style=""mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                                                                                    <tbody> <tr>
        <td width = ""auto"" align=""center"" valign=""middle"" height=""32"" style=""border-radius: 5px; border: 1px solid rgb(236, 236, 237); font-size: 13px; font-family: Roboto, Arial, Helvetica, sans-serif; text-align: center; color: rgb(208, 93, 104); font-weight: 300; padding-left: 18px; padding-right: 18px; word-break: break-word; line-height: 21px; background-color: rgb(255, 255, 255); background-clip: padding-box;""><a href=""http://crm.altaircloud.mx/Todos/calificaLlamada.aspx?idLlam=" + idLlamada.ToString + "&Calif=4"" <span style=""font-size: 13px; font-weight: 300; line-height: 21px;""><span style=""color: rgb(0, 0, 255); font-size: 13px; font-weight: 300; line-height: 21px;"">&iexcl;No he recibido ninguna llamada!</span><br style=""line-height: 24px;""></span></a></td>
                                                                                        <!-- start space width -->
                                                                                        <td valign = ""top"">
                                                                                            <Table width = ""20"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" style=""width: 20px;mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                                                                                                <tbody> <tr>
        <td valign = ""top"">
                                                                                                    </td>
                                                                                                    </tr>
                                                                                                </tbody></table>
                                                                                        </td>
                                                                                        <!--end space width -->
                                                                                        </tr>
                                                                                        <!--start space height -->
                                                                                        <tr>
        <td height = ""10"" style=""border-collapse: collapse; height: 10px; line-height: 10px; font-size: 10px;""></td>
                                                                                        </tr>
                                                                                        <!--end space height -->
                                                                                    </tbody></table>

                                                                            </td>
                                                                            </tr>
                                                                        </tbody></table>
                                                                    <!-- end button -->
                                                                </td>
                                                            </tr>
                                                        </tbody></table>
                                                </td>
                                            </tr>
                                            <!-- end text content -->
                                            <!--start space height -->
                                            <tr>
        <td height = ""20"" style=""border-collapse: collapse; height: 20px; line-height: 20px; font-size: 20px;""></td>
                                            </tr>
                                            <!--end space height -->
                                        </tbody></table>

                                </td>
                                </tr>
                                <!-- end image And content -->
                            </tbody></table>
                        <!-- end layout-2 container width 600px -->
                    </td>
                    </tr>
                </tbody></table>
            <!-- end layout-2 container width 600px -->
        </td>
        </tr>
        <!-- END LAYOUT 2  -->
    </tbody>
    <!--  START FOOTER COPY RIGHT -->
    <tbody> <tr>
        <td align =""center"" valign=""top"" style=""background-color: #3d4db8;"">
            <Table width = ""600"" align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"" Class=""container"" style=""min-width: 600px; padding-left: 20px; padding-right: 20px; width: 600px; background-color: #3d4db8;mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                <tbody> <tr>
        <td valign = ""top"">
                        <Table width = ""560"" align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"" Class=""container"" style=""width: 560px;mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                            <!--start space height -->
                            <tbody> <tr>
        <td height = ""10"" style=""border-collapse: collapse; height: 10px; line-height: 10px; font-size: 10px;""></td>
                                </tr>
                                <!--end space height -->
                                <tr>
                                    <!-- start COPY RIGHT content -->
                                    <td valign = ""top"" style=""font-size: 13px; font-family: Roboto, Arial, Helvetica, sans-serif; color: rgb(255, 255, 255); font-weight: 300; text-align: center; word-break: break-word; line-height: 21px;""><span style=""font-size: 13px; font-weight: 300; line-height: 21px;"">CRM By: AltairSoft&nbsp;, all rights reserved 2015 © </span></td>
                                    <!-- end COPY RIGHT content -->
                                </tr>
                                <!--start space height -->
                                <tr>
                                    <td height=""10"" style=""border-collapse: collapse; height: 10px; line-height: 10px; font-size: 10px;""></td>
                                </tr>
                                <!--end space height -->
                            </tbody></table>
                    </td>
                    </tr>
                </tbody></table>
        </td>
        </tr>
        <!--  END FOOTER COPY RIGHT -->
    </tbody></table></body>
</html>"
        Return HTML
    End Function
    Function CreaHTMLCitaCalif(ByVal idCita As Integer) As String
        Dim HTML As String = ""
        HTML += "
        <!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">
<html xmlns = ""http://www.w3.org/1999/xhtml"" xmlns:v = ""urn:schemas-Microsoft - com: vml"" xmlns: o = ""urn:schemas-Microsoft - com: office : office"" >
        <head>
            <meta http-equiv="" Content-Type"" content="" text/html; charset=UTF-8""/>
            <meta name="" viewport"" content="" initial-scale=1.0""/>
            <meta http-equiv="" X-UA-Compatible"" content="" IE=edge""/>
            <meta name="" format-detection"" content="" telephone=no""/>
            <title>CRM Mail</title>
            <link href="" http://fonts.googleapis.com/css?family=Roboto:400,300,700&subset=latin,cyrillic,greek"" rel=""stylesheet"" type=""text/css"">
                <!--[if mso]>
<style type=""text/css"">
td,span,a{font-family: Arial, sans-serif !important;}
a{text-decoration: none !important;}
</style>
<![endif]-->
                <style type="" text/css"">

  
    /* Resets: see reset.css for details */
    .ReadMsgBody { width: 100%; background-color: #ffffff;}
    .ExternalClass {width: 100%; background-color: #ffffff;}
    .ExternalClass, .ExternalClass p, .ExternalClass span,
    .ExternalClass font, .ExternalClass td, .ExternalClass div {line-height:100%;}
    #outlook a{ padding:0;}
    html{width: 100%; }
    body {-webkit-text-size-adjust:none; -ms-text-size-adjust:none; }
    html,body {background-color: #ffffff; margin: 0; padding: 0; }
    table {border-spacing:0;}
    table td {border-collapse:collapse;}
    br, strong br, b br, em br, i br { line-height:100%; }
    h1, h2, h3, h4, h5, h6 { line-height: 100% !important; -webkit-font-smoothing: antialiased; }
    img{height: auto !important; line-height: 100%; outline: none; text-decoration: none; display:block !important; }
    span a { text-decoration: none !important;}
    a{ text-decoration: none !important; }
    table p{margin:0;}
    .yshortcuts, .yshortcuts a, .yshortcuts a:link,.yshortcuts a:visited,
    .yshortcuts a:hover, .yshortcuts a span { text-decoration: none !important; border-bottom: none !important;}
    table{ mso-table-lspace:0pt; mso-table-rspace:0pt; }
    img{ -ms-interpolation-mode:bicubic; }
    ul{list-style: initial; margin:0; padding-left:20px;}
    /*mailChimp class*/
    .default-edit-image{
    height:20px;
    }
    .tpl-repeatblock {
    padding: 0px !important;
    border: 1px dotted rgba(0,0,0,0.2);
    }
    img{height:auto !important;}
    td[class=""image-270px""] img{
    width:270px;
    height:auto !important;
    max-width:270px !important;
    }
    td[class=""image-170px""] img{
    width:170px;
    height:auto !important;
    max-width:170px !important;
    }
    td[class=""image-185px""] img{
    width:185px;
    height:auto !important;
    max-width:185px !important;
    }
    td[class=""image-124px""] img{
    width:124px;
    height:auto !important;
    max-width:124px !important;
    }
    @media only screen and (max-width: 640px){
    body{
    width:auto!important;
    }
    table[class=""container""]{
    width: 100%!important;
    padding-left: 20px!important;
    padding-right: 20px!important;
    min-width:100% !important;
    }
    td[class=""image-270px""] img{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    }
    td[class=""image-170px""] img{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    }
    td[class=""image-185px""] img{
    width:185px !important;
    height:auto !important;
    max-width:185px !important;
    }
    td[class=""image-124px""] img{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    }
    td[class=""image-100-percent""] img{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    }
    td[class=""small-image-100-percent""] img{
    width:100% !important;
    height:auto !important;
    }
    table[class=""full-width""]{
    width:100% !important;
    min-width:100% !important;
    }
    table[class=""full-width-text""]{
    width:100% !important;
    background-color:#ffffff;
    padding-left:20px !important;
    padding-right:20px !important;
    }
    table[class=""full-width-text2""]{
    width:100% !important;
    background-color:#ffffff;
    padding-left:20px !important;
    padding-right:20px !important;
    }
    table[class=""col-2-3img""]{
    width:50% !important;
    margin-right: 20px !important;
    }
    table[class=""col-2-3img-last""]{
    width:50% !important;
    }
    table[class=""col-2-footer""]{
    width:55% !important;
    margin-right:20px !important;
    }
    table[class=""col-2-footer-last""]{
    width:40% !important;
    }
    table[class=""col-2""]{
    width:47% !important;
    margin-right:20px !important;
    }
    table[class=""col-2-last""]{
    width:47% !important;
    }
    table[class=""col-3""]{
    width:29% !important;
    margin-right:20px !important;
    }
    table[class=""col-3-last""]{
    width:29% !important;
    }
    table[class=""row-2""]{
    width:50% !important;
    }
    td[class=""text-center""]{
    text-align: center !important;
    }
    /* start clear and remove*/
    table[class=""remove""]{
    display:none !important;
    }
    td[class=""remove""]{
    display:none !important;
    }
    /* end clear and remove*/
    table[class=""fix-box""]{
    padding-left:20px !important;
    padding-right:20px !important;
    }
    td[class=""fix-box""]{
    padding-left:20px !important;
    padding-right:20px !important;
    }
    td[class=""font-resize""]{
    font-size: 18px !important;
    line-height: 22px !important;
    }
    table[class=""space-scale""]{
    width:100% !important;
    float:none !important;
    }
    table[class=""clear-align-640""]{
    float:none !important;
    }
     table[class=""show-full-mobile""]{
        display:none !important;
        width:100% !important;
        min-width:100% !important;
    }
    }
    @media only screen and (max-width: 479px){
    body{
    font-size:10px !important;
    }
    table[class=""container""]{
    width: 100%!important;
    padding-left: 10px!important;
    padding-right:10px!important;
    min-width:100% !important;
    }
    table[class=""container2""]{
    width: 100%!important;
    float:none !important;
    min-width:100% !important;
    }
    td[class=""full-width""] img{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    min-width:124px !important;
    min-width:100% !important;
    }
    td[class=""image-270px""] img{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    min-width:124px !important;
    }
    td[class=""image-170px""] img{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    min-width:124px !important;
    }
    td[class=""image-185px""] img{
    width:185px !important;
    height:auto !important;
    max-width:185px !important;
    min-width:124px !important;
    }
    td[class=""image-124px""] img{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    min-width:124px !important;
    }
    td[class=""image-100-percent""] img{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    min-width:124px !important;
    }
    td[class=""small-image-100-percent""] img{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    min-width:124px !important;
    }
    table[class=""full-width""]{
    width:100% !important;
    }
    table[class=""full-width-text""]{
    width:100% !important;
    background-color:#ffffff;
    padding-left:20px !important;
    padding-right:20px !important;
    }
    table[class=""full-width-text2""]{
    width:100% !important;
    background-color:#ffffff;
    padding-left:20px !important;
    padding-right:20px !important;
    }
    table[class=""col-2-footer""]{
    width:100% !important;
    margin-right:0px !important;
    }
    table[class=""col-2-footer-last""]{
    width:100% !important;
    }
    table[class=""col-2""]{
    width:100% !important;
    margin-right:0px !important;
    }
    table[class=""col-2-last""]{
    width:100% !important;
    }
    table[class=""col-3""]{
    width:100% !important;
    margin-right:0px !important;
    }
    table[class=""col-3-last""]{
    width:100% !important;
    }
    table[class=""row-2""]{
    width:100% !important;
    }
    table[id=""col-underline""]{
    float: none !important;
    width: 100% !important;
    border-bottom: 1px solid #eee;
    }
    td[id=""col-underline""]{
    float: none !important;
    width: 100% !important;
    border-bottom: 1px solid #eee;
    }
    td[class=""col-underline""]{
    float: none !important;
    width: 100% !important;
    border-bottom: 1px solid #eee;
    }
    /*start text center*/
    td[class=""text-center""]{
    text-align: center !important;
    }
    div[class=""text-center""]{
    text-align: center !important;
    }
    /*end text center*/
    /* start  clear and remove */
    table[id=""clear-padding""]{
    padding:0 !important;
    }
    td[id=""clear-padding""]{
    padding:0 !important;
    }
    td[class=""clear-padding""]{
    padding:0 !important;
    }
    table[class=""remove-479""]{
    display:none !important;
    }
    td[class=""remove-479""]{
    display:none !important;
    }
    table[class=""clear-align""]{
    float:none !important;
    }
    /* end  clear and remove */
    table[class=""width-small""]{
    width:100% !important;
    }
    table[class=""fix-box""]{
    padding-left:15px !important;
    padding-right:15px !important;
    }
    td[class=""fix-box""]{
    padding-left:15px !important;
    padding-right:15px !important;
    }
    td[class=""font-resize""]{
    font-size: 14px !important;
    }
    td[class=""increase-Height""]{
    height:10px !important;
    }
    td[class=""increase-Height-20""]{
    height:20px !important;
    }
    table[width=""595""]{
        width:100% !important;
      }

      table[class=""show-full-mobile""]{
        display:table !important;
        width:100% !important;
        min-width:100% !important;
      }
    }
    @media only screen and (max-width: 320px){
    table[class=""width-small""]{
    width:125px !important;
    }
    img[class=""image-100-percent""]{
    width:100% !important;
    height:auto !important;
    max-width:100% !important;
    min-width:124px !important;
    }
    }
    td ul{list-style: initial; margin:0; padding-left:20px;}

	@media only screen and (max-width: 640px){ .image-100-percent{ width:100%!important; height: auto !important; max-width: 100% !important; min-width: 124px !important;}}body{background-color:#efefef;} .default-edit-image{height:20px;} tr.tpl-repeatblock , tr.tpl-repeatblock > td{ display:block !important;} .tpl-repeatblock {padding: 0px !important;border: 1px dotted rgba(0,0,0,0.2);} table[width=""595""]{width:100% !important;}</style>
                <!--[if gte mso 9]><xml><o:OfficeDocumentSettings><o:AllowPNG/><o:PixelsPerInch>96</o:PixelsPerInch></o:OfficeDocumentSettings></xml><![endif]-->
        </head>
<body  style = ""font-size:12px; width:100%; height:100%;"" >
<table id = ""mainStructure"" width=""800"" Class=""full-width"" align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"" style=""background-color:#efefef; width:800px; max-width: 800px; margin: 0 auto; outline: 1px solid #efefef; box-shadow: 0px 0px 5px #E0E0E0;"">
    <!--START VIEW ONLINE And ICON SOCAIL -->
    <tbody><tr>
        <td align =""center"" valign=""top"" style=""background-color: #3d4db8;"">
            <!-- start container 600 -->
            <table width = ""600"" align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"" Class=""container"" style=""padding-left: 20px; padding-right: 20px; min-width: 600px; width: 600px; background-color: #3d4db8;mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                <tbody><tr>
                    <td valign = ""top"">
                        <table width = ""560"" align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"" Class=""full-width"" style=""width: 560px; background-color: #3d4db8;mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                            <!-- start space -->
                            <tbody><tr>
                                <td valign = ""top"" height=""10"" style=""border-collapse: collapse; height: 10px; line-height: 10px; font-size: 10px;"">
                                </td>
                                </tr>
                                <!-- end space -->
                                <tr>
                                    <td valign = ""top"">
                                        <!-- start container -->
                                        <table width = ""100%"" align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"" style=""mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                                            <tbody><tr>
                                                <td valign = ""top"">

                                                    <table align = ""left"" border=""0"" cellspacing=""0"" cellpadding=""0"" Class=""container2"" width=""auto"" style=""mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                                                        <tbody><tr>
                                                            <td>

                                                            </td>
                                                            </tr>
                                                            <!-- start space -->
                                                            <tr>
                                                                <td valign = ""top"" Class=""increase-Height"">
                                                                </td>
                                                            </tr>
                                                            <!-- end space -->
                                                        </tbody></table><!--[if (gte mso 9)|(IE)]></td><td valign=""top"" ><![endif]-->




                                                </td>
                                                </tr>
                                            </tbody></table>
                                        <!-- end container  -->
                                    </td>
                                </tr>
                                <!-- start space -->
                                <tr>
                                    <td valign = ""top"" height=""10"" style=""border-collapse: collapse; height: 10px; line-height: 10px; font-size: 10px;"">
                                    </td>
                                </tr>
                                <!-- end space -->
                                <!-- start space -->
                                <tr>
                                    <td valign = ""top"" Class=""increase-Height"">
                                    </td>
                                </tr>
                                <!-- end space -->
                            </tbody></table>
                        <!-- end container 600-->
                    </td>
                    </tr>
                </tbody></table>
        </td>
        </tr>
        <!--END VIEW ONLINE And ICON SOCAIL-->
    </tbody>
    <!-- START LAYOUT 2-->
    <tbody><tr>
        <td align =""center"" valign=""top"" style=""background-color: #ebebec;"">
            <!-- start layout-2 container width 600px -->
            <table width = ""600"" align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"" Class=""full-width"" style=""padding-left: 20px; padding-right: 20px; min-width: 600px; width: 600px; background-color: #ffffff;mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                <tbody><tr>
                    <td valign = ""top"">
                        <!-- start layout-2 container width 600px -->
                        <table width = ""560"" align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"" Class=""full-width"" style=""width: 560px;mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                            <!-- start image And content -->
                            <tbody><tr>
                                <td valign = ""top"" width=""100%"">

                                    <table width = ""270"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""left"" Class=""full-width"" style=""width: 270px;mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                                        <tbody><tr>
                                            <td valign = ""bottom"" align=""center"" width=""270"" style=""width: 270px;"">
                                                <a href = ""#"" style=""font-size: inherit; border-style: none; text-decoration: none!important;"" border=""0"">
                                                    <img src = ""http://altaircloud.mx/logo-default.png"" width=""200"" alt=""image1"" style=""max-width: 270px; display: block;"" border=""0"" hspace=""0"" vspace=""0"">
                              </a>
                                            </td>
                                            </tr>
                                        </tbody></table><!--[if (gte mso 9)|(IE)]></td><td valign=""top"" ><![endif]-->


                                    <Table Class=""remove"" width=""1"" border=""0"" cellpadding=""0"" cellspacing=""0"" align=""left"" style=""font-size: 0px; line-height: 0; border-collapse: collapse; width: 1px;mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                                        <tbody> <tr>
        <td width = ""0"" height=""2"" style=""border-collapse: collapse; width: 0px; height: 2px; line-height: 2px; font-size: 2px;"">
                                                <p style = ""padding-left: 20px;"">&nbsp;</p>
                                            </td>
                                            </tr>
                                        </tbody></table><!--[if (gte mso 9)|(IE)]></td><td valign=""top"" ><![endif]-->

                                    <Table width = ""270"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""right"" Class=""container"" style=""width: 270px;mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                                        <!--start space height -->
                                        <tbody> <tr>
        <td height = ""20"" style=""border-collapse: collapse; height: 20px; line-height: 20px; font-size: 20px;""></td>
                                            </tr>
                                            <!--end space height -->
                                            <!--start space height -->
                                            <tr>
        <td height = ""40"" Class=""remove"" style=""border-collapse: collapse; height: 40px; line-height: 40px; font-size: 40px;""></td>
                                            </tr>
                                            <!--end space height -->
                                            <!-- start text content -->
                                            <tr>
        <td valign = ""top"">
                                                    <Table width = ""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""left"" style=""mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                                                        <tbody> <tr>
        <td style = ""font-size: 22px; font-family: Roboto, Arial, Helvetica, sans-serif; color: rgb(85, 85, 85); font-weight: 300; text-align: left; word-break: break-word; line-height: 30px;""><span style=""color: rgb(85, 85, 85); font-size: inherit; font-weight: 300; line-height: 30px;""><a href=""#"" style=""color: rgb(85, 85, 85); font-size: inherit; border-style: none; text-decoration: none !important; line-height: 30px;"" data-mce-href=""#"" border=""0"">&nbsp;<span style=""color: rgb(0, 0, 255); font-size: 22px; font-weight: 300; line-height: 30px;"">HOLA!</span><br style=""line-height: 24px;"">R</a>ecientemente un asesor nuestro realizo una cita contigo, por favor dinos como fue esa cita<br style=""line-height: 24px;""><br style=""line-height: 24px;""></span></td>
                                                            </tr>
                                                            <!--start space height -->
                                                            <tr>
        <td height = ""15"" style=""border-collapse: collapse; height: 15px; line-height: 15px; font-size: 15px;""></td>
                                                            </tr>
                                                            <!--end space height -->
                                                            <tr>
        <td style = ""font-size: 13px; font-family: Roboto, Arial, Helvetica, sans-serif; color: rgb(163, 162, 162); font-weight: 300; text-align: left; word-break: break-word; line-height: 21px;""><span style=""font-size: inherit; line-height: 21px;""><p style=""line-height: 24px;"">Escoge una opci&oacute;n:</p><p style=""line-height: 24px;""></p><p style=""line-height: 24px;""></p><p style=""line-height: 24px;"">Buena experiencia<a href=""http://crm.altaircloud.mx/Todos/calificaLlamada.aspx?idCita=" + idCita.ToString + "&Calif=1""><img src=""http://altaircloud.mx/mail/img/Like.png"" width=""200""/></a></p></span></td>
                                                                <td style = ""font-size: 13px; font-family: Roboto, Arial, Helvetica, sans-serif; color: rgb(163, 162, 162); font-weight: 300; text-align: left; word-break: break-word; line-height: 21px;""><span style=""font-size: inherit; line-height: 21px;""><p style=""line-height: 24px;""></p><p style=""line-height: 24px;""></p><p style=""line-height: 24px;""></p><p style=""line-height: 24px;"">Mala experiencia<a href=""http://crm.altaircloud.mx/Todos/calificaLlamada.aspx?idCita=" + idCita.ToString + "&Calif=2""><img src=""http://altaircloud.mx/mail/img/dislike.png"" width=""200""/></a></p></span></td>
                                                            </tr>
                                                            <!--start space height -->
                                                            <tr>
        <td height = ""15"" style=""border-collapse: collapse; height: 15px; line-height: 15px; font-size: 15px;""></td>
                                                            </tr>
                                                            <!--end space height -->
                                                            <tr>
        <td valign = ""top"" width=""auto"">
                                                                    <!-- start button -->
                                                                    <Table border = ""0"" align=""left"" cellpadding=""0"" cellspacing=""0"" width=""auto"" style=""mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                                                                        <tbody> <tr>
        <td valign = ""top"">
                                                                                <Table border = ""0"" align=""left"" cellpadding=""0"" cellspacing=""0"" dup=""0"" width=""auto"" style=""mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                                                                                    <tbody> <tr>
        <td width = ""auto"" align=""center"" valign=""middle"" height=""32"" style=""border-radius: 5px; border: 1px solid rgb(236, 236, 237); font-size: 13px; font-family: Roboto, Arial, Helvetica, sans-serif; text-align: center; color: rgb(208, 93, 104); font-weight: 300; padding-left: 18px; padding-right: 18px; word-break: break-word; line-height: 21px; background-color: rgb(255, 255, 255); background-clip: padding-box;""><a href=""http://crm.altaircloud.mx/Todos/calificaLlamada.aspx?idCita=" + idCita.ToString + "&Calif=4"" <span style=""font-size: 13px; font-weight: 300; line-height: 21px;""><span style=""color: rgb(0, 0, 255); font-size: 13px; font-weight: 300; line-height: 21px;"">&iexcl;No se ha realizado dicha cita!</span><br style=""line-height: 24px;""></span></a></td>
                                                                                        <!-- start space width -->
                                                                                        <td valign = ""top"">
                                                                                            <Table width = ""20"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" style=""width: 20px;mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                                                                                                <tbody> <tr>
        <td valign = ""top"">
                                                                                                    </td>
                                                                                                    </tr>
                                                                                                </tbody></table>
                                                                                        </td>
                                                                                        <!--end space width -->
                                                                                        </tr>
                                                                                        <!--start space height -->
                                                                                        <tr>
        <td height = ""10"" style=""border-collapse: collapse; height: 10px; line-height: 10px; font-size: 10px;""></td>
                                                                                        </tr>
                                                                                        <!--end space height -->
                                                                                    </tbody></table>

                                                                            </td>
                                                                            </tr>
                                                                        </tbody></table>
                                                                    <!-- end button -->
                                                                </td>
                                                            </tr>
                                                        </tbody></table>
                                                </td>
                                            </tr>
                                            <!-- end text content -->
                                            <!--start space height -->
                                            <tr>
        <td height = ""20"" style=""border-collapse: collapse; height: 20px; line-height: 20px; font-size: 20px;""></td>
                                            </tr>
                                            <!--end space height -->
                                        </tbody></table>

                                </td>
                                </tr>
                                <!-- end image And content -->
                            </tbody></table>
                        <!-- end layout-2 container width 600px -->
                    </td>
                    </tr>
                </tbody></table>
            <!-- end layout-2 container width 600px -->
        </td>
        </tr>
        <!-- END LAYOUT 2  -->
    </tbody>
    <!--  START FOOTER COPY RIGHT -->
    <tbody> <tr>
        <td align =""center"" valign=""top"" style=""background-color: #3d4db8;"">
            <Table width = ""600"" align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"" Class=""container"" style=""min-width: 600px; padding-left: 20px; padding-right: 20px; width: 600px; background-color: #3d4db8;mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                <tbody> <tr>
        <td valign = ""top"">
                        <Table width = ""560"" align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"" Class=""container"" style=""width: 560px;mso-table-lspace:0pt; mso-table-rspace:0pt;"">
                            <!--start space height -->
                            <tbody> <tr>
        <td height = ""10"" style=""border-collapse: collapse; height: 10px; line-height: 10px; font-size: 10px;""></td>
                                </tr>
                                <!--end space height -->
                                <tr>
                                    <!-- start COPY RIGHT content -->
                                    <td valign = ""top"" style=""font-size: 13px; font-family: Roboto, Arial, Helvetica, sans-serif; color: rgb(255, 255, 255); font-weight: 300; text-align: center; word-break: break-word; line-height: 21px;""><span style=""font-size: 13px; font-weight: 300; line-height: 21px;"">CRM By: AltairSoft&nbsp;, all rights reserved 2015 © </span></td>
                                    <!-- end COPY RIGHT content -->
                                </tr>
                                <!--start space height -->
                                <tr>
                                    <td height=""10"" style=""border-collapse: collapse; height: 10px; line-height: 10px; font-size: 10px;""></td>
                                </tr>
                                <!--end space height -->
                            </tbody></table>
                    </td>
                    </tr>
                </tbody></table>
        </td>
        </tr>
        <!--  END FOOTER COPY RIGHT -->
    </tbody></table></body>
</html>"
        Return HTML
    End Function
    Function CreaHTMLAvisoSupervisor(ByVal Datos As CDetallesEmail) As String
        Dim HTML As String = ""
        HTML += "    <html lang = ""es"" >
        <!--<![endif]-->
            <!-- BEGIN HEAD -->

    <head>
        <meta charset = ""utf-8"" />
        <title>&iexcl;Observación! | CRM</title>
        <meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">
        <meta content = ""width=device-width, initial-scale=1"" name=""viewport"" />
        <meta content = """" name=""description"" />
        <meta content = """" name=""author"" />
        <!-- BEGIN GLOBAL MANDATORY STYLES -->       
        <link rel = ""shortcut icon"" href=""favicon.ico"" /> </head>
    <!-- END HEAD -->

    <body Class="""">
        <div Class=""container"">
            <div Class=""row"">
                <div Class=""col-md-12 coming-soon-header"">
                    <a Class=""brand"" href=""index.html"">
                        <img src = ""/assets/imagenes/logo-default.png"" alt=""logo"" width=""200"" /> </a>
                </div>
            </div>
            <div Class=""row"">
                <div Class=""col-md-6 coming-soon-content"">
                    <h1>&iexcl;Un cliente indico que no se realiz&oacute; una llamada la cual estest&aacute; registrada como exitosa !</h1>
                    <p> Datos del cliente:<br />
" + Datos.ApellidoPaterno + " " + Datos.ApellidoMaterno + " " + Datos.Nombre + "
<br />
" + Datos.EmailCliente.ToString + "
<br />
Teléfono: " + Datos.Tel.ToString + "
<br />
Email:" + Datos.EmailCliente + "
</p>
                                                                       
                </div>              
            </div>
            <!--/end row-->
            <div Class=""row"">
                <div Class=""col-md-12 coming-soon-footer""> 2014 &copy; CRM. By AltairSoft. </div>
            </div>
        </div>
    </body>

</html>"
        Return HTML
    End Function
    Function CreaHTMLAvisoSupervisorCita(ByVal Datos As CDetallesEmailCita) As String
        Dim HTML As String = ""
        HTML += "    <html lang = ""es"" >
        <!--<![endif]-->
            <!-- BEGIN HEAD -->

    <head>
        <meta charset = ""utf-8"" />
        <title>&iexcl;Observación! | CRM</title>
        <meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">
        <meta content = ""width=device-width, initial-scale=1"" name=""viewport"" />
        <meta content = """" name=""description"" />
        <meta content = """" name=""author"" />
        <!-- BEGIN GLOBAL MANDATORY STYLES -->       
     </head>
    <!-- END HEAD -->

    <body >
        <div>
          
            <div>
                <div >
                    <h1>&iexcl;Un cliente indico que no se realiz&oacute; una cita la cual est&aacute; registrada como exitosa !</h1>
                    <p> Datos del cliente:<br />
" + Datos.ApellidoPaterno + " " + Datos.ApellidoMaterno + " " + Datos.Nombre + "
<br />
" + Datos.EmailCliente.ToString + "
<br />
Teléfono: " + Datos.Tel.ToString + "
<br />
Email:" + Datos.EmailCliente + "
</p>
                                                                       
                </div>              
            </div>
            <!--/end row-->
            <div >
                <div > 2014 &copy; CRM. By AltairSoft. </div>
            </div>
        </div>
    </body>

</html>"
        Return HTML
    End Function
    Function Enviar_CorreoLlamadaCliente(ByVal id_llamada As Integer) As Boolean Implements IService1.Enviar_CorreoLlamadaCliente
        Dim Datos = Obtener_detallesLlamadaParaEmail(id_llamada)
        Dim HTML = CreaHTMLLlamadaCalif(id_llamada)


        Return Enviar_Correo(Datos.EmailCliente, HTML, "Evaluación de calidad")
    End Function
    Function Enviar_CorreoCitaCliente(ByVal idCita As Integer) As Boolean Implements IService1.Enviar_CorreoCitaCliente
        Dim Datos = Obtener_detallesCitaParaEmail(idCita)
        Dim HTML = CreaHTMLCitaCalif(idCita)


        Return Enviar_Correo(Datos.EmailCliente, HTML, "Evaluación de calidad")
    End Function
    Function Obtener_detallesCitaParaEmail(ByVal idCita As Integer) As CDetallesEmailCita

        Dim cmd As New SqlCommand("ObtenerDetallesCitaEmail", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@id_cita", idCita)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New CDetallesEmailCita
        While reader.Read

            Aux.email = DirectCast(reader.Item("email"), String)
            Aux.Nombre = DirectCast(reader.Item("Nombre"), String)
            Aux.ApellidoPaterno = DirectCast(reader.Item("ApellidoPaterno"), String)
            Aux.ApellidoMaterno = DirectCast(reader.Item("ApellidoMaterno"), String)
            Aux.EmailCliente = DirectCast(reader.Item("EmailCliente"), String)
            Aux.Tel = DirectCast(reader.Item("Tel"), String)
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)

        End While
        Conexion.Close()



        Return Aux
    End Function
    Function Obtener_detallesLlamadaParaEmail(ByVal idLlamada As Integer) As CDetallesEmail

        Dim cmd As New SqlCommand("Obtener_detallesEmailLlamada", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@id_llamada", idLlamada)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New CDetallesEmail
        While reader.Read

            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.email = DirectCast(reader.Item("email"), String)
            Aux.Nombre = DirectCast(reader.Item("Nombre"), String)
            Aux.ApellidoPaterno = DirectCast(reader.Item("ApellidoPaterno"), String)
            Aux.ApellidoMaterno = DirectCast(reader.Item("ApellidoMaterno"), String)
            Aux.EmailCliente = DirectCast(reader.Item("EmailCliente"), String)
            Aux.Tel = DirectCast(reader.Item("Tel"), String)

        End While
        Conexion.Close()

        Return Aux
    End Function
    Function Obtener_detallesEmailLlamada(ByVal idLlamada As Integer) As CDetallesEmail Implements IService1.Obtener_detallesEmailLlamada

        Dim cmd As New SqlCommand("Obtener_detallesEmailLlamada", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@id_llamada", idLlamada)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New CDetallesEmail
        While reader.Read

            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)
            Aux.email = DirectCast(reader.Item("email"), String)
            Aux.Nombre = DirectCast(reader.Item("Nombre"), String)
            Aux.ApellidoPaterno = DirectCast(reader.Item("ApellidoPaterno"), String)
            Aux.ApellidoMaterno = DirectCast(reader.Item("ApellidoMaterno"), String)
            Aux.EmailCliente = DirectCast(reader.Item("EmailCliente"), String)
            Aux.Tel = DirectCast(reader.Item("Tel"), String)

        End While
        Conexion.Close()
        If Aux.id_cliente > 0 Then
            Try
                Enviar_Correo(Aux.email, CreaHTMLAvisoSupervisor(Aux), "Problema con LLAMADA de cliente: " + Aux.ApellidoPaterno + " " + Aux.ApellidoMaterno + " " + Aux.Nombre)
            Catch ex As Exception

            End Try

        End If

        Return Aux
    End Function
    Function ObtenerDetallesCitaEmail(ByVal idCita As Integer) As CDetallesEmailCita Implements IService1.ObtenerDetallesCitaEmail

        Dim cmd As New SqlCommand("ObtenerDetallesCitaEmail", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@id_cita", idCita)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New CDetallesEmailCita
        While reader.Read

            Aux.email = DirectCast(reader.Item("email"), String)
            Aux.Nombre = DirectCast(reader.Item("Nombre"), String)
            Aux.ApellidoPaterno = DirectCast(reader.Item("ApellidoPaterno"), String)
            Aux.ApellidoMaterno = DirectCast(reader.Item("ApellidoMaterno"), String)
            Aux.EmailCliente = DirectCast(reader.Item("EmailCliente"), String)
            Aux.Tel = DirectCast(reader.Item("Tel"), String)
            Aux.id_cliente = DirectCast(reader.Item("id_cliente"), Integer)

        End While
        Conexion.Close()

        If Aux.id_cliente > 0 Then
            Try
                Enviar_Correo(Aux.email, CreaHTMLAvisoSupervisorCita(Aux), "Problema con CITA de cliente: " + Aux.ApellidoPaterno + " " + Aux.ApellidoMaterno + " " + Aux.Nombre)
            Catch ex As Exception

            End Try

        End If


        Return Aux
    End Function
    Public Function Enviar_Correo(ByVal Email_Destino As String, ByVal Mensaje As String, ByVal Asunto As String) As Boolean
        System.Threading.Thread.CurrentThread.CurrentUICulture =
System.Globalization.CultureInfo.GetCultureInfo("es-MX")
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.GetCultureInfo("es-MX")
        Dim ConfigCorreo = Obtener_configCorreo()
        Dim Servidor As New SmtpClient(ConfigCorreo.smtpServer, ConfigCorreo.puertoEmail)
        Dim Correo As New MailMessage(ConfigCorreo.emailSistema, Email_Destino.ToString, Asunto, "")
        Dim Errores As String = ""
        Dim Credenciales As New NetworkCredential(ConfigCorreo.emailSistema, ConfigCorreo.contraseñaEmail)

        Correo.IsBodyHtml = True
        Correo.Body = Mensaje.ToString
        Try

            Servidor.UseDefaultCredentials = False
            Servidor.EnableSsl = If(ConfigCorreo.SSL = 1, True, False)
            Servidor.Credentials = Credenciales
            Servidor.DeliveryMethod = SmtpDeliveryMethod.Network

            Servidor.Send(Correo)

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    Function Obtener_configCorreo() As CConfigCorreo
        Dim Resultado As New List(Of CConfigCorreo)
        Dim cmd As New SqlCommand("Obtener_configuracionEmail", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New CConfigCorreo
        While reader.Read
            Aux.emailSistema = DirectCast(reader.Item("emailSistema"), String)
            Aux.contraseñaEmail = DirectCast(reader.Item("contraseñaEmail"), String)
            Aux.smtpServer = DirectCast(reader.Item("smtpServer"), String)
            Aux.puertoEmail = DirectCast(reader.Item("puertoEmail"), Integer)
            Aux.SSL = reader.Item("SSL")
            Aux.EnviarEmails = reader.Item("EnviarEmails")

            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Aux
    End Function
    Public Class CConfigCorreo
        Public emailSistema As String
        Public contraseñaEmail As String
        Public smtpServer As String
        Public puertoEmail As Integer
        Public SSL As String
        Public EnviarEmails As String
    End Class

#End Region
#Region "EmailsBandejas"
    Function Obtener_correosDelCliente(ByVal emailCliente As String, ByVal emailEmpresa As String) As List(Of CCorreosCliente) Implements IService1.Obtener_correosDelCliente
        Dim Resultado As New List(Of CCorreosCliente)
        Dim cmd As New SqlCommand("Obtener_correosDelCliente", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@EmailCliente", emailCliente)
        cmd.Parameters.AddWithValue("@EmailEmpresa", emailEmpresa)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CCorreosCliente
        While reader.Read
            Aux = New CCorreosCliente
            Aux.id_email = reader.Item("id_email")
            Aux.emailFrom = DirectCast(reader.Item("emailFrom"), String)
            Aux.htmlMessage = DirectCast(reader.Item("htmlMessage"), String)
            Aux.emailTo = DirectCast(reader.Item("emailTo"), String)
            Aux.subjet = DirectCast(reader.Item("subjet"), String)
            Aux.fechaRecepcion = DirectCast(reader.Item("fechaRecepcion"), Date)
            Aux.Desc_estatus = DirectCast(reader.Item("Desc_estatus"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_mensajeEmailID(ByVal id_email As Integer) As String Implements IService1.Obtener_mensajeEmailID
        Dim cmd As New SqlCommand("Obtener_mensajeEmailID", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@idEmail", id_email)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As String = ""
        While reader.Read

            Aux = DirectCast(reader.Item("htmlMessage"), String)

        End While
        Conexion.Close()
        Return Aux
    End Function
    Function EnviarEmailSendGrid(ByVal emailFrom As String, ByVal emailTo As String, ByVal subjet As String, ByVal htmlMensaje As String, ByVal NAdjuntos As Integer) As Boolean Implements IService1.EnviarEmailSendGrid
        Try
            If Enviar_correo_sendgrid(emailFrom, emailTo, subjet, htmlMensaje) Then
                Return Inserta_emails(emailFrom, htmlMensaje, emailTo, subjet, 2, Now, NAdjuntos)
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try

    End Function
    Function Enviar_correo_sendgrid(ByVal EmailOrigen As String, ByVal EmailDestino As String, ByVal Asunto As String, ByVal MensajeHTML As String) As Boolean
        Dim Mensaje As New SendGridMessage
        Dim Credenciales As New NetworkCredential("roest", "roest.2016A")
        Mensaje.From = New MailAddress(EmailOrigen)



        Mensaje.AddTo(EmailDestino)
        Mensaje.Subject = Asunto
        Mensaje.Html = MensajeHTML


        Dim TransporteWEB As New Web(Credenciales)

        Try
            TransporteWEB.DeliverAsync(Mensaje)
            Return True
        Catch ex As Exception
            Return False
        End Try

        Return False
    End Function
    Function Obtener_emailsUsuario(ByVal Email As String) As List(Of CEmailsUsuario) Implements IService1.Obtener_emailsUsuario
        Dim Resultado As New List(Of CEmailsUsuario)
        Dim cmd As New SqlCommand("Obtener_emailsUsuario", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Email", Email)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CEmailsUsuario
        While reader.Read
            Aux = New CEmailsUsuario
            Aux.id_email = DirectCast(reader.Item("id_email"), Integer)
            Aux.emailFrom = DirectCast(reader.Item("emailFrom"), String)
            Aux.htmlMessage = DirectCast(reader.Item("htmlMessage"), String)
            Aux.emailTo = DirectCast(reader.Item("emailTo"), String)
            Aux.subjet = DirectCast(reader.Item("subjet"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.fechaRecepcion = DirectCast(reader.Item("fechaRecepcion"), Date)
            Aux.Nadjuntos = DirectCast(reader.Item("Nadjuntos"), Integer)
            Aux.Desc_estatus = DirectCast(reader.Item("Desc_estatus"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_enviadosUsuario(ByVal Email As String) As List(Of CEmailsUsuario) Implements IService1.Obtener_enviadosUsuario
        Dim Resultado As New List(Of CEmailsUsuario)
        Dim cmd As New SqlCommand("Obtener_enviadosUsuario", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Email", Email)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As CEmailsUsuario
        While reader.Read
            Aux = New CEmailsUsuario
            Aux.id_email = DirectCast(reader.Item("id_email"), Integer)
            Aux.emailFrom = DirectCast(reader.Item("emailFrom"), String)
            Aux.htmlMessage = DirectCast(reader.Item("htmlMessage"), String)
            Aux.emailTo = DirectCast(reader.Item("emailTo"), String)
            Aux.subjet = DirectCast(reader.Item("subjet"), String)
            Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            Aux.fechaRecepcion = DirectCast(reader.Item("fechaRecepcion"), Date)
            Aux.Nadjuntos = DirectCast(reader.Item("Nadjuntos"), Integer)
            Aux.Desc_estatus = DirectCast(reader.Item("Desc_estatus"), String)
            Resultado.Add(Aux)
        End While
        Conexion.Close()
        Return Resultado
    End Function
    Function Inserta_emails(ByVal emailFrom As String, ByVal htmlMessage As String, ByVal emailTo As String, ByVal subjet As String, ByVal id_estatus As Integer, ByVal FechaRecepcion As Date, ByVal NAdjuntos As Integer) As Integer Implements IService1.Inserta_emails

        Dim Resultado As Integer = 0
        Dim cmd As New SqlCommand("Inserta_Email", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@PemailFrom", emailFrom)
        cmd.Parameters.AddWithValue("@PhtmlMessage", htmlMessage)
        cmd.Parameters.AddWithValue("@PemailTo", emailTo)
        cmd.Parameters.AddWithValue("@Psubjet", subjet)
        cmd.Parameters.AddWithValue("@Pid_estatus", id_estatus)
        cmd.Parameters.AddWithValue("@PfechaRecepcion", FechaRecepcion)
        cmd.Parameters.AddWithValue("@PNadjuntos", NAdjuntos)
        Conexion.Close()
        Try
            Conexion.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader

            While reader.Read
                Resultado = reader.Item(0)
            End While

        Catch ex As Exception
            Conexion.Close()
            Return 0
        End Try
        Conexion.Close()
        Return Resultado
    End Function
    Function Obtener_configuracionEmailUsuario(ByVal id_usuario As Integer) As CConfiguracionEmail Implements IService1.Obtener_configuracionEmailUsuario

        Dim cmd As New SqlCommand("Obtener_configuracionEmailUsuario", Conexion)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@id_usuario", id_usuario)
        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As New CConfiguracionEmail
        While reader.Read

            Aux.popServer = DirectCast(reader.Item("popServer"), String)
            Aux.emailPort = DirectCast(reader.Item("emailPort"), Integer)
            Aux.useSSL = reader.Item("useSSL")
            Aux.emailPassword = DirectCast(reader.Item("emailPassword"), String)
            Aux.Email = DirectCast(reader.Item("Email"), String)

        End While
        Conexion.Close()
        Return Aux
    End Function
#End Region
#Region "Reportes"
    Public Function Obtener_DatosConsulta(ByVal Query As String) As DataSet Implements IService1.Obtener_DatosConsulta
        Return GE_SQL.SQlGetDataset(Query)
    End Function

    Public Function Obtener_MediosCampanas() As List(Of MediosCampana) Implements IService1.Obtener_MediosCampanas
        Dim Aux As MediosCampana
        Dim Lst As New List(Of MediosCampana)

        Dim Cmd As New SqlCommand("EXEC spRepObtener_MediosCampanas", Conexion)

        Conexion.Close()
        Conexion.Open()
        Dim Reader As SqlDataReader = Cmd.ExecuteReader
        While Reader.Read
            Aux = New MediosCampana
            With Aux
                .MedioId = Reader.Item("Id_Medio")
                .NombreMedio = Reader.Item("NombreMedio")
            End With

            Lst.Add(Aux)
        End While
        Conexion.Close()

        Return Lst
    End Function

    Public Function Obtener_VisitasByProyectoModeloSemana(ByVal FechaInicio As Date, ByVal FechaFin As Date) As String Implements IService1.Obtener_VisitasByProyectoModeloSemana
        Try
            Dim Aux As VisitasProyModSem
            Dim Lst As New List(Of VisitasProyModSem)

            Dim Cmd As New SqlCommand("spRepObtener_VisitasByProyectoModeloSemana", Conexion)
            With Cmd
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@Fecha_Inicio", FechaInicio)
                .Parameters.AddWithValue("@Fecha_Fin", FechaFin)
            End With

            Conexion.Close()
            Conexion.Open()
            Dim Reader As SqlDataReader = Cmd.ExecuteReader
            While Reader.Read
                Aux = New VisitasProyModSem
                With Aux
                    .Proyecto = Reader.Item("Proyecto")
                    .Modelo = Reader.Item("Modelo")
                    .Visitas = Reader.Item("Visitas")
                    .Semana = Reader.Item("Semana")
                End With

                Lst.Add(Aux)
            End While
            Conexion.Close()

            Return JsonConvert.SerializeObject(Lst)
        Catch ex As Exception
            Throw New FaultException(ex.Message)
        End Try
    End Function

    Public Function Obtener_VisitasByProyectoMedioSemana(ByVal FechaInicio As Date, ByVal FechaFin As Date) As String Implements IService1.Obtener_VisitasByProyectoMedioSemana
        Try
            Dim Aux As VisitasProyMedSem
            Dim Lst As New List(Of VisitasProyMedSem)

            Dim Cmd As New SqlCommand("spRepObtener_VisitasByProyectoMedioSemana", Conexion)
            With Cmd
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@Fecha_Inicio", FechaInicio)
                .Parameters.AddWithValue("@Fecha_Fin", FechaFin)
            End With

            Conexion.Close()
            Conexion.Open()
            Dim Reader As SqlDataReader = Cmd.ExecuteReader
            While Reader.Read
                Aux = New VisitasProyMedSem
                With Aux
                    .Proyecto = Reader.Item("Proyecto")
                    .Medio = Reader.Item("Medio")
                    .Visitas = Reader.Item("Visitas")
                    .Semana = Reader.Item("Semana")
                End With

                Lst.Add(Aux)
            End While
            Conexion.Close()

            Return JsonConvert.SerializeObject(Lst)
        Catch ex As Exception
            Throw New FaultException(ex.Message)
        End Try
    End Function

    Public Function Obtener_VisitasAyBMedio(ByVal FechaInicio As Date, ByVal FechaFin As Date, ByVal Medio As Integer) As List(Of VisitasMedio) Implements IService1.Obtener_VisitasAyBMedio
        Dim Aux As VisitasMedio
        Dim Lst As New List(Of VisitasMedio)

        Dim Cmd As New SqlCommand("spRepObtener_TotalVisitasAyBMedio", Conexion)
        With Cmd
            .CommandType = CommandType.StoredProcedure
            .Parameters.AddWithValue("@Fecha_Inicio", FechaInicio)
            .Parameters.AddWithValue("@Fecha_Fin", FechaFin)
            .Parameters.AddWithValue("@Medio", Medio)
        End With

        Conexion.Close()
        Conexion.Open()
        Dim Reader As SqlDataReader = Cmd.ExecuteReader
        While Reader.Read
            Aux = New VisitasMedio
            With Aux
                .Id_Visita = Reader.Item("Id_Visita")
                .Id_Medio = Reader.Item("Id_Medio")
                .Numcte = Reader.Item("Numcte")
            End With

            Lst.Add(Aux)
        End While
        Conexion.Close()

        Return Lst
    End Function

    Public Function Obtener_VisitasAyBXFraccionamiento(ByVal Fecha_Inicio As Date, ByVal Fecha_Final As Date, ByVal Proyecto As String) As Integer Implements IService1.Obtener_VisitasAyBXFraccionamiento
        Dim Cmd As New SqlCommand("spRepObtener_TotalVisitasAyB", Conexion)
        With Cmd
            .CommandType = CommandType.StoredProcedure
            .Parameters.AddWithValue("@Fecha_Inicio", Fecha_Inicio)
            .Parameters.AddWithValue("@Fecha_Fin", Fecha_Final)
            .Parameters.AddWithValue("@Proyecto", Proyecto)
        End With

        Conexion.Close()
        Conexion.Open()

        Dim TotalVisita As Integer = 0
        Dim Reader As SqlDataReader = Cmd.ExecuteReader
        While Reader.Read
            TotalVisita = Reader.Item("Total")
        End While
        Conexion.Close()

        Return TotalVisita
    End Function

    Public Function Obtener_CancelacionesEnkontrol(ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of CancelacionesEnkontrol) Implements IService1.Obtener_CancelacionesEnkontrol
        Dim Query As String = "SELECT CL.numcte, CL.nom_cte + ' ' + CL.ap_paterno_cte + ' ' + CL.ap_materno_cte cliente, 
                                      AG.empleado, AG.nom_empleado + ' ' + AG.ap_paterno_empleado + ' ' + AG.ap_materno_empleado agente,
                                      LI.empleado empleadoLider, LI.nom_empleado + ' ' + LI.ap_paterno_empleado + ' ' + LI.ap_materno_empleado lider,
                                      CN.fec_cancela cancelacion, FR.id_cve_fracc cc, FR.cve_conjunto
                               FROM sm_cliente_cancela CN 
                               INNER JOIN sm_cliente CL ON CL.numcte = CN.numcte
                               INNER JOIN sm_fraccionamiento FR ON FR.id_cve_fracc = CL.id_cve_fracc
                               INNER JOIN sm_agente AG ON AG.empleado = CL.empleado
                               INNER JOIN sm_agente_subagente AGS ON AGS.subagente = AG.empleado
                               INNER JOIN sm_agente LI ON LI.empleado = AGS.agente
                               WHERE CN.FEC_CANCELA BETWEEN '" & FechaInicio.ToString("yyyy-MM-dd") & "' AND '" & FechaFin.ToString("yyyy-MM-dd") & "' 
                               ORDER BY FR.id_cve_fracc"

        Dim Resultado As New List(Of CancelacionesEnkontrol)
        Dim Aux As CancelacionesEnkontrol

        Dim DS As New DataSet
        Dim DT_11 As New DataTable
        Dim DT_18 As New DataTable

        DS = ODBC_OBJ.ODBCGetDataset(Query, 11)
        If DS.Tables.Count > 0 Then
            DT_11 = DS.Tables(0).Copy()
        End If

        DS.Tables.Clear() : DS.Dispose()

        DS = ODBC_OBJ.ODBCGetDataset(Query, 18)
        If DS.Tables.Count > 0 Then
            DT_18 = DS.Tables(0).Copy()
        End If

        DS.Tables.Clear() : DS.Dispose()

        For Each row As DataRow In DT_11.Rows
            Aux = New CancelacionesEnkontrol
            Aux.numeroCliente = row("numcte")
            Aux.nombreCliente = "" & row("cliente")
            Aux.numeroAgente = row("empleado")
            Aux.nombreAgente = "" & row("agente")
            Aux.numeroLider = row("empleadoLider")
            Aux.nombreLider = "" & row("lider")
            Aux.fechaCancelacion = "" & row("cancelacion")
            Aux.cc = "" & row("cc")
            Aux.zona = "" & row("cve_conjunto")
            Aux.empresa = 11

            Resultado.Add(Aux)
        Next

        For Each row As DataRow In DT_18.Rows
            Aux = New CancelacionesEnkontrol
            Aux.numeroCliente = row("numcte")
            Aux.nombreCliente = "" & row("cliente")
            Aux.numeroAgente = row("empleado")
            Aux.nombreAgente = "" & row("agente")
            Aux.numeroLider = row("empleadoLider")
            Aux.nombreLider = "" & row("lider")
            Aux.fechaCancelacion = "" & row("cancelacion")
            Aux.cc = "" & row("cc")
            Aux.zona = "" & row("cve_conjunto")
            Aux.empresa = 18

            Resultado.Add(Aux)
        Next

        Return Resultado
    End Function

    Public Function Obtener_SeparacionesEnkontrol(ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of SeparacionesEnkontrol) Implements IService1.Obtener_SeparacionesEnkontrol
        Dim Query As String = "SELECT CL.numcte, CL.nom_cte + ' ' + CL.ap_paterno_cte + ' ' + CL.ap_materno_cte cliente, 
                                      AG.empleado, AG.nom_empleado + ' ' + AG.ap_paterno_empleado + ' ' + AG.ap_materno_empleado agente, 
                                      LI.empleado empleadoLider, LI.nom_empleado + ' ' + LI.ap_paterno_empleado + ' ' + LI.ap_materno_empleado lider,
                                      FR.id_cve_fracc cc, FR.cve_conjunto
                               FROM sm_cliente CL
                               INNER JOIN sm_fraccionamiento FR ON FR.id_cve_fracc = CL.id_cve_fracc
                               INNER JOIN sm_agente AG ON AG.empleado = CL.empleado
                               INNER JOIN sm_agente_subagente AGS ON AGS.subagente = AG.empleado
                               INNER JOIN sm_agente LI ON LI.empleado = AGS.agente
                               WHERE CL.ID_NUM_ETAPA = 7 AND CL.FEC_REGISTO BETWEEN '" & FechaInicio.ToString("yyyy-MM-dd") & "' AND '" & FechaFin.ToString("yyyy-MM-dd") & "'
                               ORDER BY FR.id_cve_fracc"

        Dim Resultado As New List(Of SeparacionesEnkontrol)
        Dim Aux As SeparacionesEnkontrol

        Dim DS As New DataSet
        Dim DT_11 As New DataTable
        Dim DT_18 As New DataTable

        DS = ODBC_OBJ.ODBCGetDataset(Query, 11)
        If DS.Tables.Count > 0 Then
            DT_11 = DS.Tables(0).Copy
        End If

        DS.Tables.Clear() : DS.Dispose()

        DS = ODBC_OBJ.ODBCGetDataset(Query, 18)
        If DS.Tables.Count > 0 Then
            DT_18 = DS.Tables(0).Copy
        End If

        DS.Tables.Clear() : DS.Dispose()

        For Each row As DataRow In DT_11.Rows
            Aux = New SeparacionesEnkontrol
            Aux.numeroCliente = row("numcte")
            Aux.cliente = "" & row("cliente")
            Aux.numeroAgente = row("empleado")
            Aux.agente = "" & row("agente")
            Aux.numeroLider = row("empleadoLider")
            Aux.nombreLider = "" & row("lider")
            Aux.cc = "" & row("cc")
            Aux.zona = "" & row("cve_conjunto")
            Aux.empresa = 11

            Resultado.Add(Aux)
        Next

        For Each row As DataRow In DT_18.Rows
            Aux = New SeparacionesEnkontrol
            Aux.numeroCliente = row("numcte")
            Aux.cliente = "" & row("cliente")
            Aux.numeroAgente = row("empleado")
            Aux.agente = "" & row("agente")
            Aux.numeroLider = row("empleadoLider")
            Aux.nombreLider = "" & row("lider")
            Aux.cc = "" & row("cc")
            Aux.zona = "" & row("cve_conjunto")
            Aux.empresa = 18

            Resultado.Add(Aux)
        Next

        Return Resultado
    End Function

    Public Function Obtener_ProspeccionesEnkontrol(ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of ProspeccionesEnkontrol) Implements IService1.Obtener_ProspeccionesEnkontrol
        Dim Query As String = "SELECT CL.numcte, CL.nom_cte + ' ' + CL.ap_paterno_cte + ' ' + CL.ap_materno_cte cliente, 
                                      AG.empleado, AG.nom_empleado + ' ' + AG.ap_paterno_empleado + ' ' + AG.ap_materno_empleado agente, 
                                      LI.empleado empleadoLider, LI.nom_empleado + ' ' + LI.ap_paterno_empleado + ' ' + LI.ap_materno_empleado lider,
                                      FR.id_cve_fracc cc, FR.cve_conjunto
                               FROM sm_cliente CL
                               INNER JOIN sm_fraccionamiento FR ON FR.id_cve_fracc = CL.id_cve_fracc
                               INNER JOIN sm_agente AG ON AG.empleado = CL.empleado
                               INNER JOIN sm_agente_subagente AGS ON AGS.subagente = AG.empleado
                               INNER JOIN sm_agente LI ON LI.empleado = AGS.agente
                               WHERE CL.ID_NUM_ETAPA IN (5,6) AND CL.FEC_REGISTO BETWEEN '" & FechaInicio.ToString("yyyy-MM-dd") & "' AND '" & FechaFin.ToString("yyyy-MM-dd") & "'
                               ORDER BY FR.id_cve_fracc"

        Dim Resultado As New List(Of ProspeccionesEnkontrol)
        Dim Aux As ProspeccionesEnkontrol

        Dim DS As New DataSet
        Dim DT_11 As New DataTable
        Dim DT_18 As New DataTable

        DS = ODBC_OBJ.ODBCGetDataset(Query, 11)
        If DS.Tables.Count > 0 Then
            DT_11 = DS.Tables(0).Copy
        End If

        DS.Tables.Clear() : DS.Dispose()

        DS = ODBC_OBJ.ODBCGetDataset(Query, 18)
        If DS.Tables.Count > 0 Then
            DT_18 = DS.Tables(0).Copy
        End If

        DS.Tables.Clear() : DS.Dispose()

        For Each row As DataRow In DT_11.Rows
            Aux = New ProspeccionesEnkontrol
            Aux.numeroCliente = row("numcte")
            Aux.cliente = "" & row("cliente")
            Aux.numeroAgente = row("empleado")
            Aux.agente = "" & row("agente")
            Aux.numeroLider = row("empleadoLider")
            Aux.nombreLider = "" & row("lider")
            Aux.cc = "" & row("cc")
            Aux.zona = "" & row("cve_conjunto")
            Aux.empresa = 11

            Resultado.Add(Aux)
        Next

        For Each row As DataRow In DT_18.Rows
            Aux = New ProspeccionesEnkontrol
            Aux.numeroCliente = row("numcte")
            Aux.cliente = "" & row("cliente")
            Aux.numeroAgente = row("empleado")
            Aux.agente = "" & row("agente")
            Aux.numeroLider = row("empleadoLider")
            Aux.nombreLider = "" & row("lider")
            Aux.cc = "" & row("cc")
            Aux.zona = "" & row("cve_conjunto")
            Aux.empresa = 18

            Resultado.Add(Aux)
        Next

        Return Resultado
    End Function

    Public Function Obtener_ProspeccionesCRM(ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of ProspeccionesCRM) Implements IService1.Obtener_ProspeccionesCRM
        Dim Resultado As New List(Of ProspeccionesCRM)
        Dim cmd As New SqlCommand("Reporte_ObtenerProspecciones", Conexion)

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@FechaInicio", FechaInicio)
        cmd.Parameters.AddWithValue("@FechaFin", FechaFin)

        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As ProspeccionesCRM
        While reader.Read
            Aux = New ProspeccionesCRM
            If IsDBNull(reader.Item("idCliente")) Then Aux.idCliente = 0 Else Aux.idCliente = DirectCast(reader.Item("idCliente"), Integer)
            If IsDBNull(reader.Item("cliente")) Then Aux.nombreCliente = "" Else Aux.nombreCliente = DirectCast(reader.Item("cliente"), String)
            If IsDBNull(reader.Item("idUsuario")) Then Aux.idUsuario = 0 Else Aux.idUsuario = DirectCast(reader.Item("idUsuario"), Integer)
            If IsDBNull(reader.Item("asesor")) Then Aux.nombreUsuario = "" Else Aux.nombreUsuario = DirectCast(reader.Item("asesor"), String)
            If IsDBNull(reader.Item("idSupervisor")) Then Aux.idSupervisor = 0 Else Aux.idSupervisor = DirectCast(reader.Item("idSupervisor"), Integer)
            If IsDBNull(reader.Item("supervisor")) Then Aux.nombreSupervisor = "" Else Aux.nombreSupervisor = DirectCast(reader.Item("supervisor"), String)
            If IsDBNull(reader.Item("fechaCreacion")) Then Aux.fechaCreacion = "1900-01-01" Else Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)

            Resultado.Add(Aux)
        End While
        Conexion.Close()

        Return Resultado
    End Function

    Public Function Obtener_VisitasCRM(ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of VisitasCRM) Implements IService1.Obtener_VisitasCRM
        Dim Resultado As New List(Of VisitasCRM)
        Dim cmd As New SqlCommand("Reporte_ObtenerVisitas", Conexion)

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@FechaInicio", FechaInicio)
        cmd.Parameters.AddWithValue("@FechaFin", FechaFin)

        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As VisitasCRM
        While reader.Read
            Aux = New VisitasCRM
            If IsDBNull(reader.Item("idCliente")) Then Aux.idCliente = 0 Else Aux.idCliente = DirectCast(reader.Item("idCliente"), Integer)
            If IsDBNull(reader.Item("cliente")) Then Aux.nombreCliente = "" Else Aux.nombreCliente = DirectCast(reader.Item("cliente"), String)
            If IsDBNull(reader.Item("idUsuario")) Then Aux.idUsuario = 0 Else Aux.idUsuario = DirectCast(reader.Item("idUsuario"), Integer)
            If IsDBNull(reader.Item("asesor")) Then Aux.nombreUsuario = "" Else Aux.nombreUsuario = DirectCast(reader.Item("asesor"), String)
            If IsDBNull(reader.Item("idSupervisor")) Then Aux.idSupervisor = 0 Else Aux.idSupervisor = DirectCast(reader.Item("idSupervisor"), Integer)
            If IsDBNull(reader.Item("supervisor")) Then Aux.nombreSupervisor = "" Else Aux.nombreSupervisor = DirectCast(reader.Item("supervisor"), String)
            If IsDBNull(reader.Item("fechaCreacion")) Then Aux.fechaCreacion = "1900-01-01" Else Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            If IsDBNull(reader.Item("fechaUltimaVisita")) Then Aux.fechaCreacion = "1900-01-01" Else Aux.fechaCreacion = DirectCast(reader.Item("fechaUltimaVisita"), Date)

            Resultado.Add(Aux)
        End While
        Conexion.Close()

        Return Resultado
    End Function

    Public Function Obtener_SeparacionesCRM(ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of SeparacionesCRM) Implements IService1.Obtener_SeparacionesCRM
        Dim Resultado As New List(Of SeparacionesCRM)
        Dim cmd As New SqlCommand("Reporte_ObtenerSeparaciones", Conexion)

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@FechaInicio", FechaInicio)
        cmd.Parameters.AddWithValue("@FechaFin", FechaFin)

        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As SeparacionesCRM
        While reader.Read
            Aux = New SeparacionesCRM
            If IsDBNull(reader.Item("idCliente")) Then Aux.idCliente = 0 Else Aux.idCliente = DirectCast(reader.Item("idCliente"), Integer)
            If IsDBNull(reader.Item("cliente")) Then Aux.nombreCliente = "" Else Aux.nombreCliente = DirectCast(reader.Item("cliente"), String)
            If IsDBNull(reader.Item("idUsuario")) Then Aux.idUsuario = 0 Else Aux.idUsuario = DirectCast(reader.Item("idUsuario"), Integer)
            If IsDBNull(reader.Item("asesor")) Then Aux.nombreUsuario = "" Else Aux.nombreUsuario = DirectCast(reader.Item("asesor"), String)
            If IsDBNull(reader.Item("idSupervisor")) Then Aux.idSupervisor = 0 Else Aux.idSupervisor = DirectCast(reader.Item("idSupervisor"), Integer)
            If IsDBNull(reader.Item("supervisor")) Then Aux.nombreSupervisor = "" Else Aux.nombreSupervisor = DirectCast(reader.Item("supervisor"), String)
            If IsDBNull(reader.Item("fechaCreacion")) Then Aux.fechaCreacion = "1900-01-01" Else Aux.fechaCreacion = DirectCast(reader.Item("fechaCreacion"), Date)
            If IsDBNull(reader.Item("fechaUltimaVisita")) Then Aux.fechaCreacion = "1900-01-01" Else Aux.fechaCreacion = DirectCast(reader.Item("fechaUltimaVisita"), Date)
            If IsDBNull(reader.Item("fecha_cierre")) Then Aux.fechaCreacion = "1900-01-01" Else Aux.fechaCreacion = DirectCast(reader.Item("fecha_cierre"), Date)

            Resultado.Add(Aux)
        End While
        Conexion.Close()

        Return Resultado
    End Function

    Public Function Obtener_DatosARPA(ByVal ID_Usuario As Integer, ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of ReporteArpa) Implements IService1.Obtener_DatosARPA
        Dim Resultado As New List(Of ReporteArpa)
        Dim cmd As New SqlCommand("Reportes_ARPA", Conexion)

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@ID_Usuario", ID_Usuario)
        cmd.Parameters.AddWithValue("@FechaInicio", FechaInicio)
        cmd.Parameters.AddWithValue("@FechaFin", FechaFin)

        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As ReporteArpa
        While reader.Read
            Aux = New ReporteArpa
            Aux.Prospectos = DirectCast(reader.Item("Prospectos"), Integer)
            Aux.Visitas = DirectCast(reader.Item("Visitas"), Integer)
            Aux.Separaciones = DirectCast(reader.Item("Separaciones"), Integer)

            Resultado.Add(Aux)
        End While
        Conexion.Close()

        Return Resultado
    End Function

    Public Function Obtener_DatosCumplimientoProyecto(ByVal Proyecto As String, ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of DatosCumplimiento) Implements IService1.Obtener_DatosCumplimientoProyecto
        Dim Resultado As New List(Of DatosCumplimiento)
        Dim cmd As New SqlCommand("Reporte_CumplimientoProyectoDetalle", Conexion)

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Proyecto", Proyecto)
        cmd.Parameters.AddWithValue("@FechaInicio", FechaInicio)
        cmd.Parameters.AddWithValue("@FechaFin", FechaFin)

        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As DatosCumplimiento
        While reader.Read
            Aux = New DatosCumplimiento
            Aux.semana = DirectCast(reader.Item("Semana"), Integer)
            Aux.asesor = DirectCast(reader.Item("Asesor"), String)
            Aux.cliente = DirectCast(reader.Item("Cliente"), String)
            Aux.prototipo = DirectCast(reader.Item("Prototipo"), String)
            Aux.fraccionamiento = DirectCast(reader.Item("Fraccionamiento"), String)
            Aux.fechaInicio = DirectCast(reader.Item("FechaInicio"), Date)
            Aux.etapa = DirectCast(reader.Item("Etapa"), String)
            Aux.nombreCampaña = DirectCast(reader.Item("campañaNombre"), String)

            Resultado.Add(Aux)
        End While

        Conexion.Close()

        Return Resultado
    End Function

    Public Function Obtener_DetallesCumplimientoProyecto(ByVal Proyecto As String, ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of DetallesCumplimiento) Implements IService1.Obtener_DetallesCumplimientoProyecto
        Dim Resultado As New List(Of DetallesCumplimiento)
        Dim cmd As New SqlCommand("Reporte_CumplimientoProyectoDetalles", Conexion)

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Proyecto", Proyecto)
        cmd.Parameters.AddWithValue("@FechaInicio", FechaInicio)
        cmd.Parameters.AddWithValue("@FechaFin", FechaFin)

        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As DetallesCumplimiento
        While reader.Read
            Aux = New DetallesCumplimiento
            Aux.semana = DirectCast(reader.Item("Semana"), Integer)
            Aux.asesor = DirectCast(reader.Item("Asesor"), String)
            Aux.cliente = DirectCast(reader.Item("Cliente"), String)
            Aux.Ranking = DirectCast(reader.Item("Ranking"), String)
            Aux.prototipo = DirectCast(reader.Item("Prototipo"), String)
            Aux.fraccionamiento = DirectCast(reader.Item("Fraccionamiento"), String)
            Aux.fechaInicio = DirectCast(reader.Item("FechaInicio"), Date)
            Aux.etapa = DirectCast(reader.Item("Etapa"), String)
            Aux.nombreCampaña = DirectCast(reader.Item("campañaNombre"), String)

            Resultado.Add(Aux)
        End While

        Conexion.Close()

        Return Resultado
    End Function

    Public Function Obtener_DatosCalidad(ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of ReporteCalidad) Implements IService1.Obtener_DatosCalidad
        Dim Resultado As New List(Of ReporteCalidad)
        Dim cmd As New SqlCommand("Reporte_Calidad", Conexion)

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@FechaInicio", FechaInicio)
        cmd.Parameters.AddWithValue("@FechaFin", FechaFin)

        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As ReporteCalidad
        While reader.Read
            Aux = New ReporteCalidad
            Aux.ClienteID = reader.Item("ClienteId")
            Aux.ClienteIDCRM = reader.Item("ID_CRM")
            Aux.Numcte = reader.Item("Numcte")
            Aux.Plaza = reader.Item("Plaza")
            Aux.Fraccionamiento = reader.Item("Fraccionamiento")
            Aux.Proyecto = reader.Item("Proyecto")
            Aux.Mza = reader.Item("Mza")
            Aux.Lote = reader.Item("Lote")
            Aux.TipoFraccionamiento = reader.Item("Tipo_Fracc")
            Aux.EsqFraccionamiento = reader.Item("EsqFracc")
            Aux.Cliente = reader.Item("Cliente")
            Aux.Contrato = reader.Item("Contrato")
            Aux.FechaEntrega = reader.Item("FechaEntrega")
            Aux.Telefono1 = "" & reader.Item("Telefono_1")
            Aux.Telefono2 = "" & reader.Item("Telefono_2")
            Aux.Empresa = reader.Item("Empresa")
            Aux.Departamento = reader.Item("Departamento")
            Aux.Conyuge = reader.Item("Conyuge")
            Aux.Asesor = reader.Item("Asesor")
            Aux.Integrador = reader.Item("Integrador")
            Aux.Titular = reader.Item("Titular")
            Aux.Responsable = reader.Item("Responsable")
            Aux.Ranqueo = reader.Item("Ranqueo")
            Aux.Visita = reader.Item("Visita")

            Resultado.Add(Aux)
        End While

        Conexion.Close()

        Return Resultado
    End Function

    Public Function Obtener_TotalesCumplimientoProyecto(ByVal Proyecto As String, ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of TotalesCumplimiento) Implements IService1.Obtener_TotalesCumplimientoProyecto
        Dim Resultado As New List(Of TotalesCumplimiento)
        Dim cmd As New SqlCommand("Reporte_CumplimientoProyectoCampaña", Conexion)

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@Proyecto", Proyecto)
        cmd.Parameters.AddWithValue("@FechaInicio", FechaInicio)
        cmd.Parameters.AddWithValue("@FechaFin", FechaFin)

        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As TotalesCumplimiento
        While reader.Read
            Aux = New TotalesCumplimiento
            Aux.Cantidad = DirectCast(reader.Item("Cantidad"), Integer)
            Aux.Campana = DirectCast(reader.Item("Campana"), String)
            Aux.TipoCampana = DirectCast(reader.Item("TipoCampana"), String)

            Resultado.Add(Aux)
        End While
        Conexion.Close()

        Return Resultado
    End Function

    Public Function Obtener_Proyectos() As List(Of Proyectos) Implements IService1.Obtener_Proyectos
        Dim Resultado As New List(Of Proyectos)
        Dim cmd As New SqlCommand("Obtener_Proyectos", Conexion)

        cmd.CommandType = CommandType.StoredProcedure

        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As Proyectos
        While reader.Read
            Aux = New Proyectos
            Aux.Proyecto = DirectCast(reader.Item("abrev_fracc"), String)

            Resultado.Add(Aux)
        End While
        Conexion.Close()

        Return Resultado
    End Function

    Public Function Obtener_VisitasAyBXModelo(ByVal Fecha_Final As Date) As String Implements IService1.Obtener_VisitasAyBXModelo
        Dim Cmd As New SqlCommand("spRepObtener_TotalVisitasAyB_Modelo", Conexion)
        Dim Lst As New List(Of VisitasProyModSem)
        With Cmd
            .CommandType = CommandType.StoredProcedure
            .Parameters.AddWithValue("@Fecha_Fin", Fecha_Final)
        End With

        Conexion.Close()
        Conexion.Open()

        Dim TotalVisita As Integer = 0
        Dim Reader As SqlDataReader = Cmd.ExecuteReader
        Dim Aux As VisitasProyModSem

        While Reader.Read
            Aux = New VisitasProyModSem
            Aux.Proyecto = DirectCast(Reader.Item("Proyecto"), String)
            Aux.Modelo = DirectCast(Reader.Item("Modelo"), String)
            Aux.Ano = DirectCast(Reader.Item("Año"), Integer)
            Aux.Semana = DirectCast(Reader.Item("Semana"), Integer)
            Aux.Visitas = DirectCast(Reader.Item("Total"), Integer)
            Lst.Add(Aux)
        End While
        Conexion.Close()
        Return JsonConvert.SerializeObject(Lst)
    End Function

    Public Function Obtener_VisitasAyBXMedio(ByVal Fecha_Final As Date) As String Implements IService1.Obtener_VisitasAyBXMedio
        Dim Cmd As New SqlCommand("spRepObtener_VisitasByProyectoMedioSemana", Conexion)
        Dim Lst As New List(Of VisitasProyMedSem)
        With Cmd
            .CommandType = CommandType.StoredProcedure
            .Parameters.AddWithValue("@Fecha_Fin", Fecha_Final)
        End With

        Conexion.Close()
        Conexion.Open()

        Dim TotalVisita As Integer = 0
        Dim Reader As SqlDataReader = Cmd.ExecuteReader
        Dim Aux As VisitasProyMedSem

        While Reader.Read
            Aux = New VisitasProyMedSem
            Aux.Proyecto = DirectCast(Reader.Item("Proyecto"), String)
            Aux.Medio = DirectCast(Reader.Item("Medio"), String)
            Aux.Semana = DirectCast(Reader.Item("Semana"), Integer)
            Aux.Visitas = DirectCast(Reader.Item("Visitas"), Integer)
            Lst.Add(Aux)
        End While
        Conexion.Close()

        Return JsonConvert.SerializeObject(Lst)
    End Function

    Public Function Obtener_VisitasAgenteXSemana(ByVal Fecha_Inicio As Date, ByVal Fecha_Fin As Date) As String Implements IService1.Obtener_VisitasAgenteXSemana
        Dim Cmd As New SqlCommand("spRepObtener_VisitasProspectador_Coordinador", Conexion)
        Dim Lst As New List(Of VisitasAgenteSem)
        With Cmd
            .CommandType = CommandType.StoredProcedure
            .Parameters.AddWithValue("@Fecha_Inicio", Fecha_Inicio)
            .Parameters.AddWithValue("@Fecha_Fin", Fecha_Fin)
        End With

        Conexion.Close()
        Conexion.Open()

        Dim TotalVisita As Integer = 0
        Dim Reader As SqlDataReader = Cmd.ExecuteReader
        Dim Aux As VisitasAgenteSem

        While Reader.Read
            Aux = New VisitasAgenteSem
            Aux.Agente = DirectCast(Reader.Item("Usuario"), String)
            Aux.Semana = DirectCast(Reader.Item("Semana"), Integer)
            Aux.Visitas = DirectCast(Reader.Item("Visitas"), Integer)
            Aux.ObjetivoSemanas_5 = DirectCast(IIf(String.IsNullOrEmpty(Reader.Item("objUltimas5").ToString()), 0, Reader.Item("objUltimas5")), Integer)
            Aux.ObjetivoSemanas_12 = DirectCast(IIf(String.IsNullOrEmpty(Reader.Item("objUltimas12").ToString()), 0, Reader.Item("objUltimas12")), Integer)
            Lst.Add(Aux)
        End While
        Conexion.Close()

        Return JsonConvert.SerializeObject(Lst)
    End Function

    Public Function Obtener_VisitasAyBXAgente(ByVal Fecha_Inicio As Date, ByVal Fecha_Fin As Date, ByVal Usuario As String) As Integer Implements IService1.Obtener_VisitasAyBXAgente
        Dim Result As Integer = 0
        Dim cmd As New SqlCommand("spObtener_VisitasByAgente", Conexion)
        With cmd
            .CommandType = CommandType.StoredProcedure
            .Parameters.AddWithValue("@Usuario", Usuario)
            .Parameters.AddWithValue("@FechaInicio", Fecha_Inicio)
            .Parameters.AddWithValue("@FechaFin", Fecha_Fin)
        End With

        Conexion.Close()
        Conexion.Open()
        Dim reader As SqlDataReader = cmd.ExecuteReader
        While reader.Read
            Result = reader.Item("Visitas")
        End While
        Conexion.Close()

        Return Result
    End Function

    Public Function Obtener_ARPAVisitas(ByVal Fecha As Date) As String Implements IService1.Obtener_ARPAVisitas
        Dim cmd As New SqlCommand("spRepObtener_VisitasARPA", Conexion)
        Dim Lst As New List(Of VisitasARPA)
        With cmd
            .CommandType = CommandType.StoredProcedure
            .Parameters.AddWithValue("@Fecha", Fecha)
        End With

        Conexion.Close()
        Conexion.Open()
        Dim Reader As SqlDataReader = cmd.ExecuteReader
        Dim Aux As VisitasARPA

        While Reader.Read
            Aux = New VisitasARPA
            Aux.Proyecto = DirectCast(Reader.Item("Proyecto"), String)
            Aux.Tipo_Casa = DirectCast(Reader.Item("TipoCasa"), Integer)
            Aux.Modelo = DirectCast(Reader.Item("Modelo"), String)
            Aux.Prog_Anual = DirectCast(IIf(String.IsNullOrEmpty(Reader.Item("Prog_Anual").ToString()), 0, Reader.Item("Prog_Anual")), Integer)
            Aux.Visitas_Acumuladas = DirectCast(IIf(String.IsNullOrEmpty(Reader.Item("A_B_Ano").ToString()), 0, Reader.Item("A_B_Ano")), Integer)
            Aux.Visitas_Semana = DirectCast(IIf(String.IsNullOrEmpty(Reader.Item("A_B_semana").ToString()), 0, Reader.Item("A_B_semana")), Integer)
            Aux.Semana_Anterior = DirectCast(IIf(String.IsNullOrEmpty(Reader.Item("Avance_semana_Anterior").ToString()), 0, Reader.Item("Avance_semana_Anterior")), Integer)
            Lst.Add(Aux)

        End While
        Conexion.Close()

        Return JsonConvert.SerializeObject(Lst)
    End Function
#End Region
#Region "Enkontrol"
    Public Function Obtener_OrigenAgente(ByVal Empleado As Integer) As List(Of OrigenAgente)
        Dim Query As String = "SELECT direccion_archivo
                               FROM sm_agente
                               WHERE empleado = " & Empleado & ""

        Dim Resultado As New List(Of OrigenAgente)
        Dim Aux As OrigenAgente

        Dim DS As New DataSet
        Dim DT As New DataTable

        DS = ODBC_OBJ.ODBCGetDataset(Query, 11)
        If DS.Tables.Count > 1 Then
            DT = DS.Tables(0).Copy()
        End If

        DS.Clear() : DS.Dispose()
    End Function
#End Region
#Region "Comisiones"
    Public Function Obtener_IndicadoresComisiones_Prospectador(ByVal ListadoUsuarios As String, ByVal FechaInicio As Date, ByVal FechaFin As Date, ByVal NumEmpresa As Integer) As String Implements IService1.Obtener_IndicadoresComisiones_Prospectador
        Try
            Dim Resultado As New List(Of Indicadores)
            Dim cmd As New SqlCommand("Obtener_IndicadoresProspeccion", Conexion)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@FechaInicio", FechaInicio)
            cmd.Parameters.AddWithValue("@FechaFin", FechaFin)
            cmd.Parameters.AddWithValue("@ListadoUsuarios", ListadoUsuarios)
            cmd.Parameters.AddWithValue("@PEmpresa", NumEmpresa)
            Conexion.Close()
            Conexion.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader
            Dim Aux As Indicadores
            Dim int = 1
            While reader.Read
                Aux = New Indicadores
                Aux.Empresa = If(String.IsNullOrEmpty(reader.Item("empresa").ToString), "", reader.Item("empresa"))
                Aux.Nombre_Completo_Empleado = reader.Item("nombre")
                Aux.Nombre_Cliente = reader.Item("nombre_cliente")
                Aux.apPaterno_Cliente = reader.Item("ap_Paterno_cliente")
                Aux.apMaterno_Cliente = reader.Item("ap_Materno_cliente")
                Aux.cliente = reader.Item("cliente")
                Aux.clienteEK = reader.Item("Numcte")
                Aux.Empleado = reader.Item("usuario")
                Aux.NumSeparaciones = If(IsDBNull(reader.Item("separaciones")), 0, Convert.ToInt32(reader.Item("separaciones")))
                Aux.NumVisitas = If(IsDBNull(reader.Item("visitas")), 0, Convert.ToInt32(reader.Item("visitas")))
                Aux.Modelo = reader.Item("Modelo")
                Aux.Proyecto = reader.Item("Proyecto")
                Aux.NombreCorto = reader.Item("NombreCorto")
                Aux.CC = If(String.IsNullOrEmpty(reader.Item("CC").ToString), "", reader.Item("CC"))
                Aux.Status_Agente = reader.Item("Status_Agente")
                Resultado.Add(Aux)
                int = int + 1
            End While
            Conexion.Close()

            Return JsonConvert.SerializeObject(Resultado)

        Catch ex As Exception
            Throw New FaultException(ex.Message)
        End Try
    End Function

    Public Function Obtener_IndicadoresComisiones_Prospectador_Visitas(ByVal FechaInicio As Date, ByVal FechaFin As Date) As String Implements IService1.Obtener_IndicadoresComisiones_Prospectador_Visitas
        Try
            Dim Resultado As New List(Of Indicadores)
            Dim cmd As New SqlCommand("Obtener_IndicadoresProspeccion_Visitas", Conexion)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@FechaInicio", FechaInicio)
            cmd.Parameters.AddWithValue("@FechaFin", FechaFin)
            Conexion.Close()
            Conexion.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader
            Dim Aux As Indicadores
            Dim int = 1
            While reader.Read
                Aux = New Indicadores
                Aux.Id_Visita = reader.Item("Id_Visita")
                Aux.Empresa = If(String.IsNullOrEmpty(reader.Item("empresa").ToString), "", reader.Item("empresa"))
                Aux.Nombre_Completo_Empleado = reader.Item("nombre")
                Aux.Nombre_Cliente = reader.Item("nombre_cliente")
                Aux.apPaterno_Cliente = reader.Item("ap_Paterno_cliente")
                Aux.apMaterno_Cliente = reader.Item("ap_Materno_cliente")
                Aux.cliente = reader.Item("cliente")
                Aux.clienteEK = reader.Item("Numcte")
                Aux.Empleado = reader.Item("usuario")
                If IsDBNull(reader.Item("FechaVisita")) Then
                    Aux.FechaVisita = "1900-01-01"
                Else
                    Aux.FechaVisita = DirectCast(reader.Item("FechaVisita"), Date)
                End If
                Aux.NumSeparaciones = If(IsDBNull(reader.Item("separaciones")), 0, Convert.ToInt32(reader.Item("separaciones")))
                Aux.NumVisitas = If(IsDBNull(reader.Item("visitas")), 0, Convert.ToInt32(reader.Item("visitas")))
                Aux.Modelo = reader.Item("Modelo")
                Aux.Proyecto = reader.Item("Proyecto")
                Aux.NombreCorto = reader.Item("NombreCorto")
                Aux.CC = If(String.IsNullOrEmpty(reader.Item("CC").ToString), "", reader.Item("CC"))
                Aux.Status_Agente = reader.Item("Status_Agente")
                Resultado.Add(Aux)
                int = int + 1
            End While
            Conexion.Close()

            Return JsonConvert.SerializeObject(Resultado)

        Catch ex As Exception
            Throw New FaultException(ex.Message)
        End Try
    End Function

    Public Function Obtener_IndicadoresComisiones_Prospectador_Separaciones(ByVal ListadoUsuarios As String, ByVal FechaInicio As Date, ByVal FechaFin As Date, ByVal NumEmpresa As Integer) As String Implements IService1.Obtener_IndicadoresComisiones_Prospectador_Separaciones
        Try
            Dim Resultado As New List(Of Indicadores)
            Dim cmd As New SqlCommand("Obtener_IndicadoresProspeccion_Separaciones", Conexion)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@FechaInicio", FechaInicio)
            cmd.Parameters.AddWithValue("@FechaFin", FechaFin)
            cmd.Parameters.AddWithValue("@ListadoUsuarios", ListadoUsuarios)
            cmd.Parameters.AddWithValue("@PEmpresa", NumEmpresa)
            Conexion.Close()
            Conexion.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader
            Dim Aux As Indicadores
            Dim int = 1
            While reader.Read
                Aux = New Indicadores
                Aux.Empresa = If(String.IsNullOrEmpty(reader.Item("empresa").ToString), "", reader.Item("empresa"))
                Aux.Nombre_Completo_Empleado = reader.Item("nombre")
                Aux.Nombre_Cliente = reader.Item("nombre_cliente")
                Aux.apPaterno_Cliente = reader.Item("ap_Paterno_cliente")
                Aux.apMaterno_Cliente = reader.Item("ap_Materno_cliente")
                Aux.cliente = reader.Item("cliente")
                Aux.clienteEK = reader.Item("Numcte")
                Aux.Empleado = reader.Item("usuario")
                If IsDBNull(reader.Item("FechaSeparacion")) Then
                    Aux.FechaSeparacion = "1900-01-01"
                Else
                    Aux.FechaSeparacion = DirectCast(reader.Item("FechaSeparacion"), Date)
                End If
                Aux.NumSeparaciones = If(IsDBNull(reader.Item("separaciones")), 0, Convert.ToInt32(reader.Item("separaciones")))
                Aux.NumVisitas = If(IsDBNull(reader.Item("visitas")), 0, Convert.ToInt32(reader.Item("visitas")))
                Aux.Modelo = reader.Item("Modelo")
                Aux.Proyecto = reader.Item("Proyecto")
                Aux.NombreCorto = reader.Item("NombreCorto")
                Aux.CC = If(String.IsNullOrEmpty(reader.Item("CC").ToString), "", reader.Item("CC"))
                Aux.Status_Agente = reader.Item("Status_Agente")
                Resultado.Add(Aux)
                int = int + 1
            End While
            Conexion.Close()

            Return JsonConvert.SerializeObject(Resultado)

        Catch ex As Exception
            Throw New FaultException(ex.Message)
        End Try
    End Function




    Public Function Obtener_IndicadoresComisiones_Prospectador_Complemento(ByVal ListadoUsuarios As String, ByVal FechaInicio As Date, ByVal FechaFin As Date, ByVal NumEmpresa As Integer) As String Implements IService1.Obtener_IndicadoresComisiones_Prospectador_Complemento
        Try
            Dim Resultado As New List(Of Indicadores)
            Dim cmd As New SqlCommand("Obtener_IndicadoresProspeccion_Escrituracion", Conexion)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@FechaInicio", FechaInicio)
            cmd.Parameters.AddWithValue("@FechaFin", FechaFin)
            cmd.Parameters.AddWithValue("@ListadoUsuarios", ListadoUsuarios)
            cmd.Parameters.AddWithValue("@PEmpresa", NumEmpresa)
            Conexion.Close()
            Conexion.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader
            Dim Aux As Indicadores
            Dim int = 1
            While reader.Read
                Aux = New Indicadores
                Aux.Empresa = reader.Item("empresa")
                Aux.Nombre_Completo_Empleado = reader.Item("nombre")
                Aux.Nombre_Cliente = reader.Item("nombre_cliente")
                Aux.apPaterno_Cliente = reader.Item("ap_Paterno_cliente")
                Aux.apMaterno_Cliente = reader.Item("ap_Materno_cliente")
                Aux.cliente = reader.Item("cliente")
                Aux.clienteEK = reader.Item("Numcte")
                Aux.Empleado = reader.Item("usuario")
                Aux.Modelo = reader.Item("Modelo")
                Aux.Proyecto = reader.Item("Proyecto")
                Aux.NombreCorto = reader.Item("NombreCorto")
                Aux.CC = If(String.IsNullOrEmpty(reader.Item("CC").ToString), "", reader.Item("CC"))
                Aux.Status_Agente = reader.Item("Status_Agente")
                If IsDBNull(reader.Item("fecha_escritura")) Then
                    Aux.fecha_escritura = "1900-01-01"
                Else
                    Aux.fecha_escritura = DirectCast(reader.Item("fecha_escritura"), Date)
                End If
                Resultado.Add(Aux)
                int = int + 1
            End While
            Conexion.Close()

            Return JsonConvert.SerializeObject(Resultado)

        Catch ex As Exception
            Throw New FaultException(ex.Message)
        End Try
    End Function

    Public Function Obtener_IndicadoresComisiones_CallCenter(ByVal ListadoUsuarios As String, ByVal FechaInicio As Date, ByVal FechaFin As Date, ByVal NumEmpresa As Integer) As String Implements IService1.Obtener_IndicadoresComisiones_CallCenter
        Try
            Dim Resultado As New List(Of Indicadores)
            Dim cmd As New SqlCommand("Obtener_IndicadoresCallCenter", Conexion)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@FechaInicio", FechaInicio)
            cmd.Parameters.AddWithValue("@FechaFin", FechaFin)
            cmd.Parameters.AddWithValue("@ListadoUsuarios", ListadoUsuarios)
            cmd.Parameters.AddWithValue("@PEmpresa", NumEmpresa)
            Conexion.Close()
            Conexion.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader
            Dim Aux As Indicadores
            While reader.Read
                Aux = New Indicadores
                Aux.Nombre_Completo_Empleado = reader.Item("nombre")
                Aux.Nombre_Cliente = reader.Item("nombre_cliente")
                Aux.apPaterno_Cliente = reader.Item("ap_Paterno_cliente")
                Aux.apMaterno_Cliente = reader.Item("ap_Materno_cliente")
                Aux.Empleado = Convert.ToInt32(reader.Item("usuario"))
                Aux.NumSeparaciones = If(IsDBNull(reader.Item("separaciones")), 0, Convert.ToInt32(reader.Item("separaciones")))
                Aux.NumVisitas = If(IsDBNull(reader.Item("visitas")), 0, Convert.ToInt32(reader.Item("visitas")))
                Aux.Proyecto = reader.Item("abrev_fracc")
                Aux.Modelo = reader.Item("id_producto")
                Aux.Status_Agente = reader.Item("Status_Agente")
                Resultado.Add(Aux)
            End While
            Conexion.Close()

            Return JsonConvert.SerializeObject(Resultado)
        Catch ex As Exception
            Throw New FaultException(ex.Message)
        End Try
    End Function

    Public Function Obtener_IndicadoresComisiones_Cerradores(ByVal ListadoUsuarios As String, ByVal FechaInicio As Date, ByVal FechaFin As Date) As String Implements IService1.Obtener_IndicadoresComisiones_Cerradores
        Try
            Dim Resultado As New List(Of Indicadores)
            Dim cmd As New SqlCommand("Obtener_IndicadoresCerradores", Conexion)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@FechaInicio", FechaInicio)
            cmd.Parameters.AddWithValue("@FechaFin", FechaFin)
            cmd.Parameters.AddWithValue("@ListadoUsuarios", ListadoUsuarios)
            Conexion.Close()
            Conexion.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader
            Dim Aux As Indicadores
            While reader.Read
                Aux = New Indicadores
                Aux.Nombre_Completo_Empleado = reader.Item("nombre")
                Aux.Empleado = Convert.ToInt32(reader.Item("usuario"))
                Aux.NumSeparaciones = If(IsDBNull(reader.Item("separaciones")), 0, Convert.ToInt32(reader.Item("separaciones")))
                Aux.Proyecto = reader.Item("abrev_fracc")
                Aux.Modelo = reader.Item("id_producto")
                Aux.Status_Agente = reader.Item("Status_Agente")
                Resultado.Add(Aux)
            End While
            Conexion.Close()

            Return JsonConvert.SerializeObject(Resultado)

        Catch ex As Exception
            Throw New FaultException(ex.Message)
        End Try
    End Function

    Public Function Obtener_IndicadoresComisiones_Moviles(ByVal ListadoUsuarios As String, ByVal FechaInicio As Date, ByVal FechaFin As Date) As String Implements IService1.Obtener_IndicadoresComisiones_Moviles
        Try
            Dim Resultado As New List(Of Indicadores)
            Dim cmd As New SqlCommand("Obtener_IndicadoresMoviles", Conexion)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@FechaInicio", FechaInicio)
            cmd.Parameters.AddWithValue("@FechaFin", FechaFin)
            cmd.Parameters.AddWithValue("@ListadoUsuarios", ListadoUsuarios)
            Conexion.Close()
            Conexion.Open()
            Dim reader As SqlDataReader = cmd.ExecuteReader
            Dim Aux As Indicadores
            While reader.Read
                Aux = New Indicadores
                Aux.Nombre_Completo_Empleado = reader.Item("nombre")
                Aux.Empleado = Convert.ToInt32(reader.Item("usuario"))
                Aux.NumSeparaciones = If(IsDBNull(reader.Item("separaciones")), 0, Convert.ToInt32(reader.Item("separaciones")))
                Aux.Proyecto = reader.Item("abrev_fracc")
                Aux.Modelo = reader.Item("id_producto")
                Aux.Status_Agente = reader.Item("Status_Agente")
                Resultado.Add(Aux)
            End While
            Conexion.Close()

            Return JsonConvert.SerializeObject(Resultado)

        Catch ex As Exception
            Throw New FaultException(ex.Message)
        End Try
    End Function
#End Region
End Class
