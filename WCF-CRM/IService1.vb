' NOTE: You can use the "Rename" command on the context menu to change the interface name "IService1" in both code and config file together.
<ServiceContract()>
Public Interface IService1
#Region "Adm"
    <OperationContract()>
    Function Obtener_nombreClientesAdm() As List(Of CCLientesSupervisor)
    <OperationContract()>
    Function Obtener_nombresUsuariosAdm() As List(Of CSupervisorUsuarios)
    <OperationContract()>
    Function Obtener_supervisor_Adm(ByVal id_supervisor As Integer) As CDetallesSupervisor
    <OperationContract()>
    Function Actualiza_contraseña_Admin(ByVal id_usuario As Integer, ByVal Contraseña As String) As Boolean
    <OperationContract()>
    Function Obtener_clientesAdm() As List(Of ClientesSupervisor)
    <OperationContract()>
    Function Obtener_AcumuladosAdm(ByVal FechaInicio As Date, ByVal FechaFinal As Date) As List(Of CAcumuladosSupervisor)
#End Region
#Region "Status"
    <OperationContract()>
    Function Obtener_totalesUsuario(ByVal idUsuario As Integer) As CTotalesUsuario
#End Region
#Region "Error Login"
    <OperationContract()>
    Function Inserta_error(ByVal Mensaje As String, ByVal Operacion As String) As Boolean
#End Region
#Region "LogIn"
    <OperationContract()>
    Function LogIN(ByVal Usuario As String, ByVal ConstraseñaMD5 As String) As CUsuarios
#End Region
#Region "Clientes"
    <OperationContract()>
    Function Inserta_ClientesCC(ByVal id_cliente As Integer, ByVal id_usuario As Integer, tipoCredito As String) As Boolean
    <OperationContract()>
    Function CompletaCliente(ByVal id_cliente As Integer,
                                ByVal NPersona As String, ByVal NContrato As String,
                                ByVal RFC As String, ByVal NHijos As Integer,
                                ByVal IngresosPersonales As Integer,
                                ByVal Edo_Civil As String) As Boolean
    <OperationContract()>
    Function Actualiza_rankingCliente(ByVal id_cliente As Integer, ByVal ranking As String) As Boolean
    <OperationContract>
    Function Actualizar_Ranking(ByVal id_cliente As Integer, ByVal id_usuario As Integer, ByVal ranking_org As String, ByVal ranking_nvo As String) As Boolean
    <OperationContract>
    Function Actualizar_Ranking_Visitas(ByVal id_cliente As Integer, ByVal id_usuario As Integer, ByVal ranking_org As String, ByVal ranking_nvo As String, ByVal id_Visita As Integer, ByVal id_Impedimento As Integer) As Boolean
    <OperationContract()>
    Function ValidaEmail(ByVal Email As String) As Boolean
    <OperationContract()>
    Function ValidaCliente(ByVal nombre As String, ByVal app1 As String, ByVal app2 As String) As Boolean
    <OperationContract()>
    Function Obtener_ids_cliente(ByVal id_cliente As Integer) As List(Of CidCliente)
    <OperationContract()>
    Function Obtener_nombresClientesidUsuario(ByVal id_usuario As Integer) As List(Of CNombresCliente)
    <OperationContract()>
    Function VerificaCliente(ByVal idcliente As Integer, ByVal idusuario As Integer) As Boolean
    <OperationContract()>
    Function Inserta_clientes(ByVal Nombre As String, ByVal ApellidoPaterno As String, ByVal ApellidoMaterno As String, ByVal Email As String, ByVal id_producto As Integer, ByVal id_nivel As Integer, ByVal id_empresa As Integer, ByVal id_etapaActual As Integer, ByVal id_campaña As Integer, ByVal id_usuarioOriginal As Integer, ByVal Observaciones As String, ByVal fotografia As String, ByVal fotoTpresentacion As String, ByVal NSS As String, ByVal CURP As String, ByVal FechaNacimiento As Date, ByVal id_cve_fracc As String) As Integer
    <OperationContract()>
    Function Elimina_clientes(ByVal id_cliente As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_clientes(ByVal id_cliente As Integer, ByVal Nombre As String, ByVal ApellidoPaterno As String, ByVal ApellidoMaterno As String, ByVal Email As String, ByVal id_producto As Integer, ByVal id_nivel As Integer, ByVal id_empresa As Integer, ByVal id_campaña As Integer, ByVal Observaciones As String, ByVal fotografia As String, ByVal fotoTpresentacion As String, ByVal Monto As Decimal) As Boolean
    <OperationContract>
    Function Actualiza_clientes_callcenter(ByVal id_cliente As Integer, ByVal Nombre As String, ByVal ApellidoPaterno As String, ByVal ApellidoMaterno As String, ByVal Email As String, ByVal id_producto As Integer, ByVal id_nivel As Integer, ByVal id_empresa As Integer, ByVal id_campaña As Integer, ByVal Observaciones As String, ByVal fotografia As String, ByVal fotoTpresentacion As String, ByVal Monto As Decimal, ByVal NSS As String, ByVal FechaNacimiento As Date) As Boolean
    <OperationContract()>
    Function Obtener_ClienteObservaciones(ByVal idCliente As Integer) As List(Of CClienteObservaciones)
    <OperationContract()>
    Function Obtener_clientes_Todos() As List(Of CClientes)
    <OperationContract()>
    Function Obtener_clientes_detalles_todos() As List(Of CClientesDetalles)
    <OperationContract()>
    Function Obtener_clientes_detalles_idUsuario(ByVal id_usuario As Integer) As List(Of CClientesDetalles)
    <OperationContract()>
    Function Obtener_Clientes_Telefonos_idCliente(ByVal id_cliente As Integer) As List(Of CClienteTelefonos)
    <OperationContract()>
    Function Obtener_Clientes_detalles_idCliente(ByVal id_cliente As Integer) As List(Of CClientesDetalles)
    <OperationContract()>
    Function Obtener_Clientes_AsesorCallCenter(ByVal id_cliente As Integer) As List(Of AsesorCallCenter)
    <OperationContract()>
    Function Obtener_Clientes_TipoCredito_idCliente(ByVal id_cliente As Integer) As List(Of CClientesTipoCredito)
    <OperationContract()>
    Function Actualiza_ultimafecha(ByVal idcliente As Integer) As Boolean
    <OperationContract()>
    Function ComprobarNSS(ByVal NSS As String) As Integer
    <OperationContract()>
    Function ComprobarCURP(ByVal CURP As String) As Integer
#End Region
#Region "Campañas"
    <OperationContract()>
    Function Obtener_combo_campañas() As List(Of CComboCampañas)
    <OperationContract()>
    Function Inserta_campañas(ByVal campañaNombre As String, ByVal id_tipoCampaña As Integer, ByVal id_MedioCampaña As Integer, ByVal fechaInicio As Date, ByVal fechaFinal As Date, ByVal Observaciones As String) As Boolean
    <OperationContract()>
    Function Elimina_campañas(ByVal id_campaña As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_campañas(ByVal id_campaña As Integer, ByVal campañaNombre As String, ByVal id_tipoCampaña As Integer, ByVal fechaCreacion As Date, ByVal fechaInicio As Date, ByVal fechaFinal As Date, ByVal Observaciones As String) As Boolean
    <OperationContract()>
    Function Obtener_campañas() As List(Of CCampaña)
    <OperationContract()>
    Function Obtener_campañas_idCampaña(ByVal id_campaña As Integer) As CCampaña
    <OperationContract>
    Function Obtener_campañasDetalles() As List(Of CCampañaDetalles)
#End Region
#Region "Categorias Productos"
    <OperationContract()>
    Function Inserta_categoriasProductos(ByVal categoria As String) As Boolean
    <OperationContract()>
    Function Elimina_categoriasProductos(ByVal id_categoria As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_categoriasProductos(ByVal id_categoria As Integer, ByVal categoria As String) As Boolean
    <OperationContract()>
    Function Obtener_categoriasProductos() As List(Of CCategoriasProducto)
#End Region
#Region "Citas"
    <OperationContract()>
    Function Insertar_Cita(ByVal IdCliente As Integer, ByVal IdUsuario As Integer, ByVal IdUsuarioAsignado As Integer, ByVal IdCampana As Integer, ByVal TipoCampana As String,
                           ByVal Origen As String, ByVal LugarContacto As String, ByVal Proyecto As String, ByVal Modelo As Integer, ByVal VigenciaInicial As Date,
                           ByVal VigenciaFinal As Date, ByVal FechaCita As Date, ByVal Ranking As String, ByVal Status As Integer) As Boolean
    <OperationContract()>
    Function Insertar_CitasCallCenter(ByVal IdCliente As Integer, ByVal IdUsuario As Integer, ByVal IdUsuarioAsignado As Integer, ByVal IdCampana As Integer, ByVal TipoCampana As String,
                           ByVal Origen As String, ByVal LugarContacto As String, ByVal Proyecto As String, ByVal Modelo As Integer, ByVal VigenciaInicial As Date,
                           ByVal VigenciaFinal As Date, ByVal FechaCita As Date, ByVal Ranking As String, ByVal Status As Integer) As Boolean

    <OperationContract()>
    Function Insertar_CitasProspectador(ByVal IdCliente As Integer, ByVal IdUsuario As Integer, ByVal IdUsuarioAsignado As Integer, ByVal IdCampana As Integer, ByVal TipoCampana As String,
                           ByVal Origen As String, ByVal LugarContacto As String, ByVal Proyecto As String, ByVal Modelo As Integer, ByVal VigenciaInicial As Date,
                           ByVal VigenciaFinal As Date, ByVal FechaCita As Date, ByVal Ranking As String, ByVal Status As Integer) As Boolean
    <OperationContract()>
    Function Insertar_CitasCaseta(ByVal IdCliente As Integer, ByVal IdUsuario As Integer, ByVal IdUsuarioAsignado As Integer, ByVal IdCampana As Integer, ByVal TipoCampana As String,
                           ByVal Origen As String, ByVal LugarContacto As String, ByVal Proyecto As String, ByVal Modelo As Integer, ByVal VigenciaInicial As Date,
                           ByVal VigenciaFinal As Date, ByVal FechaCita As Date, ByVal Ranking As String, ByVal Status As Integer) As Boolean
    <OperationContract()>
    Function Insertar_CitaCallCenter(ByVal IdCliente As Integer, ByVal IdUsuario As Integer, ByVal IdUsuarioAsignado As Integer, ByVal Origen As String,
                                     ByVal LugarContacto As String, ByVal Proyecto As String, ByVal Modelo As Integer, ByVal VigenciaInicial As Date,
                                     ByVal VigenciaFinal As Date, ByVal FechaCita As Date, ByVal Estatus As String, ByVal IdCampana As Integer,
                                     ByVal TipoCampana As String, ByVal Activa As Integer) As Boolean
    <OperationContract()>
    Function Insertar_ObservacionesCitas(ByVal IdCita As Integer, ByVal IdUsuario As Integer, ByVal Completada As Integer, ByVal Observaciones As String) As Boolean
    <OperationContract()>
    Function Inserta_CitasCall(ByVal id_cliente As Integer, ByVal id_usuarioCC As Integer, ByVal id_usuarioAsesor As Integer, ByVal Origen As String, ByVal Lugar_Contacto As String, ByVal ProyectoVisita As String, ByVal Modelo As String, ByVal VigenciaInicio As Date, ByVal VigenciaFinal As Date, ByVal FechaCita As Date, ByVal Estatus As String) As Boolean
    <OperationContract()>
    Function Verificar_VigenciaCitas(ByVal Id_Cliente As Integer) As List(Of VigenciaCitas)
    <OperationContract()>
    Function Verificar_VigenciaCita(ByVal Id_Cliente As Integer) As List(Of VigenciaCitas)
    <OperationContract()>
    Function Inserta_citas(ByVal id_cliente As Integer, ByVal id_usuario As Integer, ByVal Fecha As Date, ByVal HoraProgramacion As String, ByVal Programada As String, ByVal AvisoCliente As String, ByVal AvisoUsuario As String, ByVal realizada As String, ByVal ObservacionUsuario As String, ByVal ObservacionCliente As String, ByVal HoraTermino As String, ByVal Lugar As String, ByVal ConfimacionCliente As String) As Integer
    <OperationContract()>
    Function Elimina_citas(ByVal id_cita As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_citas(ByVal id_cita As Integer, ByVal id_cliente As Integer, ByVal id_usuario As Integer, ByVal Fecha As Date, ByVal fechaCreacion As Date, ByVal HoraProgramacion As String, ByVal Programada As String, ByVal AvisoCliente As String, ByVal AvisoUsuario As String, ByVal realizada As String, ByVal ObservacionUsuario As String, ByVal ObservacionCliente As String, ByVal HoraTermino As String, ByVal Lugar As String, ByVal ConfimacionCliente As String) As Boolean
    <OperationContract()>
    Function Obtener_citas_detalles_idusuario(ByVal id_usuario As Integer) As List(Of CDetallesCitaUsuario)
    <OperationContract()>
    Function Obtener_citas_id_usuario(ByVal id_usuario As Integer) As List(Of CCitas)
    <OperationContract()>
    Function Obtener_citas_detalles_idCliente(ByVal id_cliente As Integer) As List(Of CDetallesCitaUsuario)
    <OperationContract()>
    Function CalificaCita(ByVal id_llamada As Integer, ByVal Calificacion As Integer) As Boolean
    <OperationContract()>
    Function ObservacionesCita(ByVal idCita As Integer, ByVal Observaciones As String, ByVal Realizada As Integer) As Boolean
    <OperationContract()>
    Function Obtener_observacionesCita(ByVal idCita As Integer) As CObservacionesCita
    <OperationContract()>
    Function Obtener_citasCliente(ByVal idCLiente As Integer) As List(Of CCitasDetallesCliente)
#End Region
#Region "Visitas"
    <OperationContract()>
    Function Insertar_VisitasClientes(ByVal IdCita As Integer, ByVal IdCliente As Integer, ByVal IdUsuario As Integer, ByVal IdUsuarioAsignado As Integer, ByVal IdUsuarioVisita As Integer,
                                      ByVal IdCampana As Integer, ByVal IdImpedimento As Integer, ByVal TipoCredito As String, ByVal Monto As Double, ByVal Ranking As String,
                                      ByVal Origen As String, ByVal Proyecto As String, ByVal Modelo As Integer, ByVal TipoCampana As String, ByVal VigenciaIncial As Date,
                                      ByVal VigenciaFinal As Date, ByVal FechaVisita As Date, ByVal Status As Integer) As Boolean
#End Region
#Region "Configuraciones"
    <OperationContract()>
    Function Actualiza_configuraciones(ByVal id_configuracion As Integer, ByVal diasDeGracias As Integer, ByVal emailSistema As String, ByVal contraseñaEmail As String, ByVal smtpServer As String, ByVal puertoEmail As Integer, ByVal SSL As String, ByVal EnviarEmails As String) As Boolean
    <OperationContract()>
    Function Obtener_configuraciones() As CConfiguraciones
#End Region
#Region "Contacto Empresa"
    <OperationContract()>
    Function Inserta_ContactoEmpresa(ByVal id_empresa As Integer, ByVal Nombre As String, ByVal ApellidoPaterno As String, ByVal ApellidoMaterno As String, ByVal Email As String, ByVal Observaciones As String, ByVal fotografia As String) As Boolean
    <OperationContract()>
    Function Elimina_ContactoEmpresa(ByVal id_contactoEmpresa As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_ContactoEmpresa(ByVal id_contactoEmpresa As Integer, ByVal id_empresa As Integer, ByVal Nombre As String, ByVal ApellidoPaterno As String, ByVal ApellidoMaterno As String, ByVal Email As String, ByVal Observaciones As String, ByVal fotografia As String) As Boolean
    <OperationContract()>
    Function Obtener_ContactoEmpresa(ByVal id_empresa As Integer) As List(Of CContactoEmpresa)
    <OperationContract()>
    Function Obtener_detalles_empresa_idEmpresa(ByVal id_contactoEmpresa As Integer) As CContactoEmpresa
#End Region
#Region "Empresas"
    <OperationContract()>
    Function Obtener_empresasComboBusqueda(ByVal Query As String) As List(Of CComboEmpresas)
    <OperationContract()>
    Function Obtener_combo_empresas() As List(Of CComboEmpresas)
    <OperationContract()>
    Function Inserta_empresas(ByVal Empresa As String, ByVal Razon_Social As String, ByVal Direccion As String, ByVal PaginaWEb As String, ByVal Horario As String, ByVal id_rubro As Integer, ByVal id_ciudad As Integer, ByVal email As String, ByVal Observaciones As String, ByVal logotipo As String) As Boolean
    <OperationContract()>
    Function Elimina_empresas(ByVal id_empresa As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_empresas(ByVal id_empresa As Integer, ByVal Empresa As String, ByVal Razon_Social As String, ByVal Direccion As String, ByVal PaginaWEb As String, ByVal fechaCreacion As Date, ByVal Horario As String, ByVal id_rubro As Integer, ByVal id_ciudad As Integer, ByVal email As String, ByVal Observaciones As String, ByVal logotipo As String) As Boolean
    <OperationContract()>
    Function Obtener_empresas() As List(Of CEmpresas)
    <OperationContract()>
    Function Obtener_detalles_empresas() As List(Of CEmpresasDetalles)
    <OperationContract()>
    Function Obtener_detalles_empresas_idEmpresa(ByVal id_empresa As Integer) As CEmpresasDetalles
    <OperationContract()>
    Function Obtener_detallesEmpresa_idCliente(ByVal id_cliente As Integer) As CEmpresasDetalles
#End Region
#Region "Operaciones"
    <OperationContract()>
    Function Inserta_operaciones(ByVal id_cliente As Integer, ByVal id_usuario As Integer, ByVal id_etapa As Integer, ByVal FechaInicio As Date, ByVal FechaFinal As Date, ByVal Observaciones As String) As Boolean
    <OperationContract()>
    Function Elimina_operaciones(ByVal id_operacion As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_operaciones(ByVal id_operacion As Integer, ByVal id_cliente As Integer, ByVal id_usuario As Integer, ByVal id_etapa As Integer, ByVal FechaInicio As Date, ByVal FechaFinal As Date, ByVal Observaciones As String) As Boolean
    <OperationContract()>
    Function Obtener_operaciones(ByVal id_cliente As Integer) As List(Of COperaciones)
    <OperationContract()>
    Function Obtener_operacionesIdCliente(ByVal idCliente As Integer) As List(Of COperacionesCliente)
#End Region
#Region "Etapas Cliente"
    <OperationContract()>
    Function Inserta_etapasCliente(ByVal nEtapa As Integer, ByVal Descripcion As String) As Boolean
    <OperationContract()>
    Function Obtener_etapasCliente() As List(Of CEtapasCliente)
    <OperationContract()>
    Function Elimina_etapasCliente(ByVal id_etapa As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_etapasCliente(ByVal id_etapa As Integer, ByVal nEtapa As Integer, ByVal Descripcion As String) As Boolean
#End Region
#Region "Llamadas"
    <OperationContract()>
    Function Obtener_llamadasPendientesHoyUsuario(ByVal id_usuario As Integer) As List(Of CLlamadasPendientesHoyUsuario)
    <OperationContract()>
    Function Inserta_llamadas(ByVal id_cliente As Integer, ByVal id_usuario As Integer, ByVal Fecha As Date, ByVal HoraProgramacion As Date, ByVal Programada As String, ByVal AvisoCliente As String, ByVal AvisoUsuario As String, ByVal realizada As String, ByVal ObservacionUsuario As String, ByVal ObservacionCliente As String) As Integer
    <OperationContract()>
    Function Actualiza_llamadas(ByVal id_llamada As Integer, ByVal id_cliente As Integer, ByVal id_usuario As Integer, ByVal Fecha As Date, ByVal fechaCreacion As Date, ByVal HoraProgramacion As String, ByVal Programada As String, ByVal AvisoCliente As String, ByVal AvisoUsuario As String, ByVal realizada As String, ByVal ObservacionUsuario As String, ByVal ObservacionCliente As String) As Boolean
    <OperationContract()>
    Function Elimina_llamadas(ByVal id_llamada As Integer) As Boolean
    <OperationContract()>
    Function Obtener_llamadas_usuario(ByVal id_usuario As Integer) As List(Of CLlamadas)
    <OperationContract()>
    Function Obtener_llamadas_cliente(ByVal id_cliente As Integer) As List(Of CLlamadas)
    <OperationContract()>
    Function Obtener_llamadasCliente(ByVal idCliente As Integer) As List(Of CLlamadasCliente)
    <OperationContract()>
    Function Cambia_realizadaLlamada(ByVal id_llamada As Integer) As Boolean
    <OperationContract()>
    Function CalificaLlamada(ByVal id_llamada As Integer, ByVal Calificacion As Integer) As Boolean
    <OperationContract()>
    Function Obtener_llamadasFechaUsuario(ByVal idUsuario As Integer, ByVal FechaInicio As Date, ByVal FechaFinal As Date) As List(Of CLlamadasFechas)
#End Region
#Region "nivelInteres"
    <OperationContract()>
    Function Inserta_nivelinteres(ByVal nivelinteres As String) As Boolean
    <OperationContract()>
    Function Elimina_nivelinteres(ByVal id_nivelInteres As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_nivelinteres(ByVal id_nivelInteres As Integer, ByVal nivelinteres As String) As Boolean
    <OperationContract()>
    Function Obtener_nivelinteres() As List(Of CNivelInteres)
#End Region
#Region "Productos"
    <OperationContract()>
    Function Obtener_datos_comboProductos() As List(Of CComboProductos)
    <OperationContract()>
    Function Inserta_productos(ByVal NombreCorto As String, ByVal NombreCompleto As String, ByVal Descripcion As String, ByVal PrecioNormal As Integer, ByVal PrecioDescuento As Integer, ByVal id_categoria As Integer, ByVal fechaCreacion As Date, ByVal Observaciones As String, ByVal fotografia As String) As Boolean
    <OperationContract()>
    Function Elimina_productos(ByVal id_producto As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_productos(ByVal id_producto As Integer, ByVal NombreCorto As String, ByVal NombreCompleto As String, ByVal Descripcion As String, ByVal PrecioNormal As Integer, ByVal PrecioDescuento As Integer, ByVal id_Categoria As Integer, ByVal Observaciones As String) As Boolean
    <OperationContract()>
    Function Obtener_productos() As List(Of CProductos)
    <OperationContract()>
    Function Obtener_Productos_Detalles() As List(Of CProductosDetalles)
    <OperationContract()>
    Function Obtener_detallesProducto(ByVal idProducto As Integer) As CDetallesProducto

#End Region
#Region "referencias"
    <OperationContract()>
    Function Inserta_referencias(ByVal id_cliente As Integer, ByVal Nombre As String, ByVal ApellidoPaterno As String, ByVal ApellidoMaterno As String, ByVal email As String, ByVal fechaCreacion As Date, ByVal id_tiporeferencia As Integer, ByVal id_usuario As Integer, ByVal Observaciones As String, ByVal fotografia As String, ByVal fotoTPresentacion As String) As Boolean
    <OperationContract()>
    Function Elimina_referencias(ByVal id_referencia As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_referencias(ByVal id_referencia As Integer, ByVal id_cliente As Integer, ByVal Nombre As String, ByVal ApellidoPaterno As String, ByVal ApellidoMaterno As String, ByVal email As String, ByVal fechaCreacion As Date, ByVal id_tiporeferencia As Integer, ByVal id_usuario As Integer, ByVal Observaciones As String, ByVal fotografia As String, ByVal fotoTPresentacion As String) As Boolean
    <OperationContract()>
    Function Obtener_referencias_cliente(ByVal id_cliente As Integer) As List(Of CReferenciasCliente)
#End Region
#Region "rubros"
    <OperationContract()>
    Function Inserta_rubros(ByVal rubro As String) As Boolean
    <OperationContract()>
    Function Elimina_rubros(ByVal id_rubro As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_rubros(ByVal id_rubro As Integer, ByVal rubro As String) As Boolean
    <OperationContract()>
    Function Obtener_rubros() As List(Of CRubros)
#End Region
#Region "Oportunidades"
    <OperationContract()>
    Function Obtener_oportunidades() As List(Of COportunidades)
#End Region
#Region "Supervisores"
    <OperationContract()>
    Function Obtener_reporte_clientesFechas(ByVal FInicio As Date, ByVal FFinal As Date) As List(Of Obtener_reporte_clientesFechas)
    <OperationContract()>
    Function Obtener_nombresClientesidSupervisor(ByVal id_supervisor As Integer) As List(Of CCLientesSupervisor)
    <OperationContract()>
    Function Cambia_usuarioCliente(ByVal id_usuario As Integer, ByVal idCliente As Integer) As Boolean
    <OperationContract()>
    Function Cambia_usuarioClienteSupervisor(ByVal id_usuario As Integer, ByVal idCliente As Integer, ByVal idSupervsor As Integer) As Boolean
    <OperationContract()>
    Function Obtener_UsuarioDetalleSupervisor(ByVal id_supervisor As Integer) As List(Of CUsuariosDetalleSup)
    <OperationContract()>
    Function Obtener_DetalleSupervisor() As List(Of CDetalleSupervisor)
    <OperationContract()>
    Function Inserta_supervisores(ByVal nombre As String, ByVal apellidoPaterno As String, ByVal apellidoMaterno As String, ByVal email As String, ByVal usuario As String, ByVal contraseña As String, ByVal fechaCreacion As Date, ByVal fotografia As String) As Boolean
    <OperationContract()>
    Function Elimina_supervisores(ByVal id_supervisor As Integer) As Boolean
    <OperationContract()>
    Function Obtener_supervisores() As List(Of CSupervisores)
    <OperationContract()>
    Function Obtener_clientesSupervisor(ByVal idSupervisor As Integer) As List(Of ClientesSupervisor)
    <OperationContract()>
    Function Obtener_usuariosSupervisor(ByVal id_supervisor As Integer) As List(Of CUsuariosSupervisor)
    <OperationContract()>
    Function Obtener_AcumuladosSupervisor(ByVal id_supervisor As Integer, ByVal FechaInicio As Date, ByVal FechaFinal As Date) As List(Of CAcumuladosSupervisor)
    <OperationContract()>
    Function Obtener_nombresUsuariosSupervisor(ByVal id_supervisor As Integer) As List(Of CSupervisorUsuarios)
    <OperationContract()>
    Function DiasSinTrabajar(ByVal id_supervisor As Integer) As List(Of DiasSinTrabajar)
    <OperationContract()>
    Function DiasSinTrabajarEtapa(ByVal id_supervisor As Integer, ByVal Etapa As Integer) As List(Of DiasSinTrabajar)
    <OperationContract()>
    Function DiasSinTrabajarEtapaFiltro(ByVal id_supervisor As Integer, ByVal Etapa As Integer, ByVal Dias As Integer, ByVal FechaInicio As Date, ByVal FechaFinal As Date) As List(Of DiasSinTrabajar)
    <OperationContract()>
    Function DiasSinTrabajarFiltro(ByVal id_supervisor As Integer, ByVal Filtro As String) As List(Of DiasSinTrabajar)
    <OperationContract()>
    Function Obtener_supervisor_Detalles(ByVal id_supervisor As Integer) As CDetallesSupervisor
    <OperationContract()>
    Function Actualiza_contraseaSupervisor(ByVal id_usuario As Integer, ByVal Contraseña As String) As Boolean
    <OperationContract()>
    Function Actualiza_supervisores(ByVal id_usuario As Integer, ByVal nombre As String, ByVal apellidoPaterno As String, ByVal apellidoMaterno As String, ByVal Email As String, ByVal Activo As Integer) As Boolean
#End Region
#Region "supervisorUsuario"
    <OperationContract()>
    Function Inserta_supervisorUsuario(ByVal id_usuario As Integer, ByVal id_supervisor As Integer) As Boolean
    <OperationContract()>
    Function Elimina_supervisorUsuario(ByVal id_supervisorusuario As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_supervisorUsuario(ByVal id_supervisorusuario As Integer, ByVal id_usuario As Integer, ByVal id_supervisor As Integer) As Boolean
    <OperationContract()>
    Function Obtener_relacion_supervisorUsuario() As List(Of CRelacionSupervisorUsuario)
#End Region
#Region "TelefonoCliente"
    <OperationContract()>
    Function Obtener_telefonosModificaCliente(ByVal idCliente As Integer) As List(Of CTelefonosmodifica)
    <OperationContract()>
    Function Inserta_telefonoCliente(ByVal Principal As String, ByVal id_cliente As Integer, ByVal Telefono As String) As Boolean
    <OperationContract()>
    Function Elimina_telefonoCliente(ByVal id_telefonoCliente As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_telefonoCliente(ByVal id_telefonoCliente As Integer, ByVal Principal As String, ByVal id_cliente As Integer, ByVal Telefono As String) As Boolean
    <OperationContract()>
    Function Obtener_telefonoCliente(ByVal id_cliente As Integer) As List(Of CTelefonoCliente)
#End Region
#Region "telefonoContactoEmpresa"
    <OperationContract()>
    Function Inserta_telefonoContactoEmpresa(ByVal id_contactoEmpresa As Integer, ByVal principal As String, ByVal Telefono As String) As Boolean
    <OperationContract()>
    Function Actualiza_telefonoContactoEmpresa(ByVal id_telefonoContacto As Integer, ByVal id_contactoEmpresa As Integer, ByVal principal As String, ByVal Telefono As String) As Boolean
    <OperationContract()>
    Function Elimina_telefonoContactoEmpresa(ByVal id_telefonoContacto As Integer) As Boolean
    <OperationContract()>
    Function Obtener_telefonoContactoEmpresa(ByVal id_contactoempresa As Integer) As List(Of CTelefonoContactoEmpresa)
#End Region
#Region "TelefonoEmpresa"
    <OperationContract()>
    Function Inserta_TelefonoEmpresa(ByVal Principal As String, ByVal id_empresa As Integer, ByVal telefono As String) As Boolean
    <OperationContract()>
    Function Elimina_TelefonoEmpresa(ByVal id_telefonoEmpresa As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_TelefonoEmpresa(ByVal id_telefonoEmpresa As Integer, ByVal Principal As String, ByVal id_empresa As Integer, ByVal telefono As String) As Boolean
    <OperationContract()>
    Function Obtener_TelefonoEmpresa(ByVal id_empresa As Integer) As List(Of CTelefonoEmpresa)
#End Region
#Region "Telefono Referencia"
    <OperationContract()>
    Function Inserta_telefonoReferencia(ByVal Principal As String, ByVal id_referencia As Integer, ByVal Telefono As String) As Boolean
    <OperationContract()>
    Function Elimina_telefonoReferencia(ByVal id_telefonoReferencia As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_telefonoReferencia(ByVal id_telefonoReferencia As Integer, ByVal Principal As String, ByVal id_referencia As Integer, ByVal Telefono As String) As Boolean
    <OperationContract()>
    Function Obtener_telefonoReferencia(ByVal id_referencia As Integer) As List(Of CTelefonoReferencia)
#End Region
#Region "Telefono Supervisor"
    <OperationContract()>
    Function Inserta_telefonoSupervisor(ByVal Principal As String, ByVal id_supervisor As Integer, ByVal Telefono As String) As Boolean
    <OperationContract()>
    Function Elimina_telefonoSupervisor(ByVal id_telefonoSupervisor As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_telefonoSupervisor(ByVal id_telefonoSupervisor As Integer, ByVal Principal As String, ByVal id_supervisor As Integer, ByVal Telefono As String) As Boolean
    <OperationContract()>
    Function Obtener_telefonoSupervisor(ByVal id_supervisor As Integer) As List(Of CTelefonoSupervisor)
#End Region
#Region "Telefono Usuario"
    <OperationContract()>
    Function Inserta_telefonoUsuario(ByVal Principal As String, ByVal id_usuario As Integer, ByVal Telefono As String) As Boolean
    <OperationContract()>
    Function Elimina_telefonoUsuario(ByVal id_telefonoUsuario As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_telefonoUsuario(ByVal id_telefonoUsuario As Integer, ByVal Principal As String, ByVal id_usuario As Integer, ByVal Telefono As String) As Boolean
    <OperationContract()>
    Function Obtener_telefonoUsuario(ByVal id_usuario As Integer) As List(Of CTelefonoUsuario)
#End Region
#Region "Tipo Campaña"
    <OperationContract()>
    Function Inserta_tipocampaña(ByVal TipoCampaña As String) As Boolean
    <OperationContract()>
    Function Elimina_tipocampaña(ByVal id_tipoCampaña As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_tipocampaña(ByVal id_tipoCampaña As Integer, ByVal TipoCampaña As String) As Boolean
    <OperationContract()>
    Function Obtener_tipocampaña() As List(Of CTipoCampaña)
#End Region
#Region "TipoReferencia"
    <OperationContract()>
    Function Inserta_tiporeferencia(ByVal tiporeferencia As String) As Boolean
    <OperationContract()>
    Function Elimina_tiporeferencia(ByVal id_tiporeferencia As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_tiporeferencia(ByVal id_tiporeferencia As Integer, ByVal tiporeferencia As String) As Boolean
    <OperationContract()>
    Function Obtener_tiporeferencia() As List(Of CTipoReferencia)
#End Region
#Region "Usuarios"
    <OperationContract()>
    Function Inserta_usuarios(ByVal nombre As String, ByVal apellidoPaterno As String, ByVal apellidoMaterno As String, ByVal Email As String, ByVal usuario As String, ByVal contraseña As String, ByVal TipoUsuario As Integer, ByVal fotografia As String, ByVal Usuario_Coordinador As Integer, ByVal Coordinador As String) As Integer
    <OperationContract()>
    Function Elimina_usuarios(ByVal id_usuario As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_usuarios(ByVal id_usuario As Integer, ByVal nombre As String, ByVal apellidoPaterno As String, ByVal apellidoMaterno As String, ByVal Email As String, ByVal activo As Integer) As Boolean
    <OperationContract()>
    Function Actualiza_usuariosPass(ByVal id_usuario As Integer, ByVal nombre As String, ByVal apellidoPaterno As String, ByVal apellidoMaterno As String, ByVal Email As String, ByVal usuario As String, ByVal contraseña As String, ByVal activo As Integer, ByVal TipoUsuario As Integer, ByVal id_Supervisor As Integer) As Boolean
    <OperationContract()>
    Function Obtener_usuarios_todos() As List(Of CUsuarios)
    <OperationContract()>
    Function Obtener_usuarios_detalles(ByVal id_usuario As Integer) As CUsuarios
    <OperationContract()>
    Function Actualiza_contraseaUsuario(ByVal id_usuario As Integer, ByVal Contraseña As String) As Boolean
    <OperationContract()>
    Function VerificaUsuario(ByVal Usuario As String) As Boolean
#End Region
#Region "Estados"
    <OperationContract()>
    Function Obtener_estados() As List(Of CEstados)
    <OperationContract()>
    Function Obtener_ciudad(ByVal id_estado As Integer) As List(Of CCiudades)
#End Region
#Region "Avance Etapas"
    <OperationContract>
    Function Valida_EtapaCliente(ByVal IdCliente As Integer, ByVal IdEtapa As Integer, ByVal IdUsuario As Integer, ByVal IdProducto As Integer) As Boolean
    <OperationContract()>
    Function Avanza_EtapaCliente(ByVal id_cliente As Integer, ByVal id_usuario As Integer, ByVal id_etapa As Integer, ByVal Observaciones As String, ByVal id_productoRegistro As Integer) As Boolean
    <OperationContract()>
    Function Obtener_etapasClienteDetalles(ByVal id_cliente As Integer) As List(Of CEtapasDetalles)
#End Region
#Region "Tareas"
    <OperationContract()>
    Function Inserta_tareas(ByVal descripcion As String, ByVal id_prioridad As Integer, ByVal id_usuario As Integer, ByVal avisado As String, ByVal fechaCreacion As Date, ByVal fechaProgramada As Date, ByVal HoraProgramada As String) As Boolean
    <OperationContract()>
    Function Actualiza_tareas(ByVal id_tarea As Integer, ByVal descripcion As String, ByVal id_prioridad As Integer, ByVal id_usuario As Integer, ByVal avisado As String, ByVal fechaCreacion As Date, ByVal fechaProgramada As Date, ByVal HoraProgramada As String) As Boolean
    <OperationContract()>
    Function Elimina_tareas(ByVal id_tarea As Integer) As Boolean
    <OperationContract()>
    Function Obtener_tareas_prioridad() As List(Of CTareasPrioridad)
    <OperationContract()>
    Function Obtener_tareasPendientesUsuario(ByVal id_usuario As Integer) As List(Of CTareasPendientes)
    <OperationContract()>
    Function Obtener_tareasTerminadasUsuario(ByVal id_usuario As Integer) As List(Of CTareasPendientes)
    <OperationContract()>
    Function TerminarTarea(ByVal id_tarea As Integer) As Boolean
#End Region
#Region "Emails"
    <OperationContract()>
    Function Obtener_detallesEmailLlamada(ByVal idLlamada As Integer) As CDetallesEmail
    <OperationContract()>
    Function ObtenerDetallesCitaEmail(ByVal idCita As Integer) As CDetallesEmailCita
    <OperationContract()>
    Function Enviar_CorreoLlamadaCliente(ByVal id_llamada As Integer) As Boolean
    <OperationContract()>
    Function Enviar_CorreoCitaCliente(ByVal idCita As Integer) As Boolean
#End Region
#Region "EmailsBandejas"
    <OperationContract()>
    Function Obtener_correosDelCliente(ByVal emailCliente As String, ByVal emailEmpresa As String) As List(Of CCorreosCliente)
    <OperationContract()>
    Function Obtener_mensajeEmailID(ByVal id_email As Integer) As String
    <OperationContract()>
    Function EnviarEmailSendGrid(ByVal emailFrom As String, ByVal emailTo As String, ByVal subjet As String, ByVal htmlMensaje As String, ByVal NAdjuntos As Integer) As Boolean
    <OperationContract()>
    Function Inserta_emails(ByVal emailFrom As String, ByVal htmlMessage As String, ByVal emailTo As String, ByVal subjet As String, ByVal id_estatus As Integer, ByVal FechaRecepcion As Date, ByVal NAdjuntos As Integer) As Integer
    <OperationContract()>
    Function Inserta_email_adjuntos(ByVal filename As String, ByVal MediaType As String, ByVal id_email As Integer, ByVal Body64Str As String) As Boolean
    <OperationContract()>
    Function Obtener_emailsUsuario(ByVal Email As String) As List(Of CEmailsUsuario)
    <OperationContract()>
    Function Obtener_enviadosUsuario(ByVal Email As String) As List(Of CEmailsUsuario)
    <OperationContract()>
    Function Obtener_configuracionEmailUsuario(ByVal id_usuario As Integer) As CConfiguracionEmail
#End Region
#Region "Reportes"
    <OperationContract()>
    Function Obtener_CancelacionesEnkontrol(ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of CancelacionesEnkontrol)

    <OperationContract()>
    Function Obtener_SeparacionesEnkontrol(ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of SeparacionesEnkontrol)

    <OperationContract()>
    Function Obtener_ProspeccionesEnkontrol(ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of ProspeccionesEnkontrol)

    <OperationContract()>
    Function Obtener_ProspeccionesCRM(ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of ProspeccionesCRM)

    <OperationContract()>
    Function Obtener_VisitasCRM(ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of VisitasCRM)

    <OperationContract()>
    Function Obtener_SeparacionesCRM(ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of SeparacionesCRM)

    <OperationContract()>
    Function Obtener_DatosARPA(ByVal ID_Usuario As Integer, ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of ReporteArpa)

    <OperationContract()>
    Function Obtener_DatosCalidad(ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of ReporteCalidad)

    <OperationContract()>
    Function Obtener_DetallesCumplimientoProyecto(ByVal Proyecto As String, ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of DetallesCumplimiento)

    <OperationContract()>
    Function Obtener_DatosCumplimientoProyecto(ByVal Proyecto As String, ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of DatosCumplimiento)

    <OperationContract()>
    Function Obtener_TotalesCumplimientoProyecto(ByVal Proyecto As String, ByVal FechaInicio As Date, ByVal FechaFin As Date) As List(Of TotalesCumplimiento)

    <OperationContract()>
    Function Obtener_Proyectos() As List(Of Proyectos)
#End Region
#Region "Enkontrol"

#End Region
#Region "Usuarios"
    <OperationContract()>
    Function Obtener_TipoUsuario() As List(Of TipoUsuario)

    <OperationContract()>
    Function Actualizar_Coordinador(ByVal NumEmpleado As Integer, ByVal NumCordinador As Integer, ByVal Nombre_Cordinador As String) As Boolean


#End Region

    <OperationContract()>
    Function Obtener_notificaciones(ByVal idCliente As Integer) As List(Of CNotifica)
    ' TODO: Add your service operations here

End Interface

' Use a data contract as illustrated in the sample below to add composite types to service operations. <DataContract()>
<DataContract()>
Public Class Obtener_reporte_clientesFechas
    <DataMember(Order:=0)>
    Public id_cliente As Integer
    <DataMember(Order:=1)>
    Public Nombre As String
    <DataMember(Order:=2)>
    Public ApellidoPaterno As String
    <DataMember(Order:=3)>
    Public ApellidoMaterno As String
    <DataMember(Order:=4)>
    Public Email As String
    <DataMember(Order:=5)>
    Public Producto As String
    <DataMember(Order:=6)>
    Public nivelinteres As String
    <DataMember(Order:=7)>
    Public id_empresa As Integer
    <DataMember(Order:=8)>
    Public FechaDe As Date
    <DataMember(Order:=9)>
    Public Etapa As String
    <DataMember(Order:=10)>
    Public Campaña As String
    <DataMember(Order:=11)>
    Public usuario As String
    <DataMember(Order:=12)>
    Public monto As Decimal
    <DataMember(Order:=13)>
    Public UltimoMovimiento As String
    <DataMember(Order:=14)>
    Public UltimaObservacion As String
End Class

<DataContract()>
Public Class CLlamadasFechas
    <DataMember()>
    Public Cliente As String
    <DataMember()>
    Public id_llamada As Integer
    <DataMember()>
    Public Fecha As Date
    <DataMember()>
    Public HoraProgramacion As TimeSpan
    <DataMember()>
    Public realizada As String
    <DataMember()>
    Public ObservacionUsuario As String
End Class

<DataContract()>
Public Class CConfiguracionEmail
    <DataMember(Order:=0)>
    Public popServer As String
    <DataMember(Order:=1)>
    Public emailPort As Integer
    <DataMember(Order:=2)>
    Public useSSL As String
    <DataMember(Order:=3)>
    Public emailPassword As String
    <DataMember(Order:=4)>
    Public Email As String
End Class

<DataContract()>
Public Class CEmailsUsuario
    <DataMember(Order:=0)>
    Public id_email As Integer
    <DataMember(Order:=1)>
    Public emailFrom As String
    <DataMember(Order:=2)>
    Public htmlMessage As String
    <DataMember(Order:=3)>
    Public emailTo As String
    <DataMember(Order:=4)>
    Public subjet As String
    <DataMember(Order:=5)>
    Public fechaCreacion As Date
    <DataMember(Order:=6)>
    Public fechaRecepcion As Date
    <DataMember(Order:=7)>
    Public Nadjuntos As Integer
    <DataMember(Order:=8)>
    Public Desc_estatus As String
End Class

<DataContract()>
Public Class CCorreosCliente
    <DataMember(Order:=0)>
    Public id_email As Integer
    <DataMember(Order:=1)>
    Public emailFrom As String
    <DataMember(Order:=2)>
    Public htmlMessage As String
    <DataMember(Order:=3)>
    Public emailTo As String
    <DataMember(Order:=4)>
    Public subjet As String
    <DataMember(Order:=5)>
    Public fechaRecepcion As Date
    <DataMember(Order:=6)>
    Public Desc_estatus As String
End Class

<DataContract()>
Public Class COportunidades
    <DataMember(Order:=1)>
    Public Property id_cliente As Integer
    <DataMember(Order:=2)>
    Public Property Nombre As String
    <DataMember(Order:=3)>
    Public Property ApellidoPaterno As String
    <DataMember(Order:=4)>
    Public Property ApellidoMaterno As String
    <DataMember(Order:=5)>
    Public Property Email As String
    <DataMember(Order:=6)>
    Public Property NombreCorto As String
    <DataMember(Order:=7)>
    Public Property nivelinteres As String
    <DataMember(Order:=8)>
    Public Property Empresa As String
    <DataMember(Order:=9)>
    Public Property fechaCreacion As Date
    <DataMember(Order:=10)>
    Public Property Etapa As String
    <DataMember(Order:=11)>
    Public Property campañaNombre As String
    <DataMember(Order:=12)>
    Public Property Observaciones As String
End Class

<DataContract()>
Public Class CDetallesSupervisor
    <DataMember()>
    Public is_supervisor As Integer = 0
    <DataMember()>
    Public nombre As String = "-"
    <DataMember()>
    Public apellidoPaterno As String = "-"
    <DataMember()>
    Public apellidoMaterno As String = "-"
    <DataMember()>
    Public Email As String = "-"
    <DataMember()>
    Public usuario As String = "-"
    <DataMember()>
    Public fechaCreacion As Date = New Date

End Class

<DataContract()>
Public Class CDetallesProducto
    <DataMember()>
    Public id_producto As Integer
    <DataMember()>
    Public NombreCorto As String
    <DataMember()>
    Public NombreCompleto As String
    <DataMember()>
    Public Descripcion As String
    <DataMember()>
    Public PrecioNormal As Integer
    <DataMember()>
    Public PrecioDescuento As Integer
    <DataMember()>
    Public id_categoria As Integer
    <DataMember()>
    Public categoria As String
    <DataMember()>
    Public Observaciones As String
End Class

<DataContract()>
Public Class CCitasDetallesCliente
    <DataMember()>
    Public id_cita As Integer
    <DataMember()>
    Public FechaProgramada As Date
    <DataMember()>
    Public HoraProgramacion As TimeSpan
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public Programada As String
    <DataMember()>
    Public AvisoCliente As String
    <DataMember()>
    Public AvisoUsuario As String
    <DataMember()>
    Public realizada As String
    <DataMember()>
    Public Observaciones As String
    <DataMember()>
    Public HoraTermino As TimeSpan
    <DataMember()>
    Public Lugar As String
    <DataMember()>
    Public Calificacion As String
End Class

<DataContract()>
Public Class CObservacionesCita
    <DataMember()>
    Public Realizada As Integer
    <DataMember()>
    Public Observaciones As String
End Class

<DataContract()>
Public Class CDetallesEmailCita
    <DataMember()>
    Public email As String
    <DataMember()>
    Public Nombre As String
    <DataMember()>
    Public ApellidoPaterno As String
    <DataMember()>
    Public ApellidoMaterno As String
    <DataMember()>
    Public EmailCliente As String
    <DataMember()>
    Public Tel As String
    <DataMember()>
    Public id_cliente As Integer
End Class

<DataContract()>
Public Class CDetallesEmail
    <DataMember()>
    Public id_cliente As Integer
    <DataMember()>
    Public email As String
    <DataMember()>
    Public Nombre As String
    <DataMember()>
    Public ApellidoPaterno As String
    <DataMember()>
    Public ApellidoMaterno As String
    <DataMember()>
    Public EmailCliente As String
    <DataMember()>
    Public Tel As String
End Class

<DataContract()>
Public Class CNotifica
    <DataMember()>
    Public TituloNotificacion As String
    <DataMember()>
    Public DescripcionNotificacion As String
    <DataMember()>
    Public URL As String
End Class
Public Class CTareasPorAvisar
    Public id_tarea As Integer
    Public prioridad As String
    Public descripcion As String
    Public fechaCreacion As Date
    Public fechaProgramada As Date
    Public HoraProgramada As TimeSpan
    Public Email As String
End Class
Public Class CCitasPendientes
    Public id_cita As Integer
    Public id_cliente As Integer
    Public id_usuario As Integer
    Public Fecha As Date
    Public fechaCreacion As Date
    Public HoraProgramacion As TimeSpan
    Public Programada As Integer
    Public AvisoCliente As Integer
    Public AvisoUsuario As Integer
    Public realizada As Integer
    Public ObservacionUsuario As String
    Public ObservacionCliente As String
    Public HoraTermino As TimeSpan
    Public Lugar As String
    Public ConfimacionCliente As Integer
    Public nombre As String
    Public apellidoPaterno As String
    Public apellidoMaterno As String
    Public Email As String
    Public Nombre1 As String
    Public ApellidoPaterno1 As String
    Public ApellidoMaterno1 As String
    Public Telefono As String
    Public Principal As Integer
End Class
Public Class CLlamadasPAvisar
    Public Nombre As String
    Public ApellidoPaterno As String
    Public ApellidoMaterno As String
    Public Email As String
    Public ObservacionUsuario As String
    Public Telefono As String
    Public Hora As String
    Public EmailUsuario As String
    Public id_llamada As Integer
End Class

<DataContract()>
Public Class DiasSinTrabajar
    <DataMember()>
    Public ID As Integer
    <DataMember()>
    Public Cliente As String
    <DataMember()>
    Public Ultima As String
    <DataMember()>
    Public Dias As Integer
End Class

<DataContract()>
Public Class CSupervisorUsuarios
    <DataMember()>
    Public id_usuario As Integer
    <DataMember()>
    Public Usuario As String
End Class

<DataContract()>
Public Class CCLientesSupervisor
    <DataMember()>
    Public id_cliente As Integer
    <DataMember()>
    Public Cliente As String
End Class

<DataContract()>
Public Class CAcumuladosSupervisor
    <DataMember()>
    Public Cantidad As Integer
    <DataMember()>
    Public NombreCliente As String
    <DataMember()>
    Public Producto As String
    <DataMember()>
    Public Empresa As String
    <DataMember()>
    Public Etapa As String
    <DataMember()>
    Public Usuario As String
End Class

<DataContract()>
Public Class CUsuariosSupervisor
    <DataMember()>
    Public id_usuario As Integer
    <DataMember()>
    Public nombre As String
    <DataMember()>
    Public apellidoPaterno As String
    <DataMember()>
    Public apellidoMaterno As String
    <DataMember()>
    Public Email As String
    <DataMember()>
    Public usuario As String
    <DataMember()>
    Public fechaCreacion As Date
End Class

<DataContract()>
Public Class ClientesSupervisor
    <DataMember()>
    Public id_cliente As Integer
    <DataMember()>
    Public Nombre As String
    <DataMember()>
    Public ApellidoPaterno As String
    <DataMember()>
    Public ApellidoMaterno As String
    <DataMember()>
    Public Email As String
    <DataMember()>
    Public Producto As String
    <DataMember()>
    Public Empresa As String
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public Descripcion As String
    <DataMember()>
    Public Usuario As String
    <DataMember()>
    Public Observaciones As String
    <DataMember()>
    Public fotografia As String
    <DataMember()>
    Public fotoTpresentacion As String
End Class

<DataContract()>
Public Class Cids
    <DataMember()>
    Public id_supervisor As Integer
    <DataMember()>
    Public id_usuario As Integer
    <DataMember()>
    Public id_usuario1 As Integer
End Class

<DataContract()>
Public Class CUsuariosDetalleSup
    <DataMember()>
    Public id_usuario As Integer
    <DataMember()>
    Public nombre As String
    <DataMember()>
    Public apellidoPaterno As String
    <DataMember()>
    Public apellidoMaterno As String
    <DataMember()>
    Public Email As String
    <DataMember()>
    Public usuario As String
    <DataMember()>
    Public contraseña As String
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public activo As Integer
    <DataMember()>
    Public id_TipoUsuario As Integer
    <DataMember()>
    Public TipousuarioDes As String
    <DataMember()>
    Public id_supervisor As Integer
    <DataMember()>
    Public SupervisorDes As String
End Class

<DataContract()>
Public Class CDetalleSupervisor
    <DataMember()>
    Public id_supervisor As Integer
    <DataMember()>
    Public nombre As String
    <DataMember()>
    Public apellidoPaterno As String
    <DataMember()>
    Public apellidoMaterno As String
    <DataMember()>
    Public Email As String
    <DataMember()>
    Public usuario As String
    <DataMember()>
    Public contraseña As String
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public activo As Integer
    <DataMember()>
    Public BorraEK As Integer
End Class

<DataContract()>
Public Class CLlamadasPendientesHoyUsuario
    <DataMember()>
    Public Nombre As String
    <DataMember()>
    Public ApellidoPaterno As String
    <DataMember()>
    Public ApellidoMaterno As String
    <DataMember()>
    Public Email As String
    <DataMember()>
    Public Producto As String
    <DataMember()>
    Public Fecha As Date
    <DataMember()>
    Public HORA As TimeSpan
End Class

<DataContract()>
Public Class CTareasPendientes
    <DataMember()>
    Public id_tarea As Integer
    <DataMember()>
    Public descripcion As String
    <DataMember()>
    Public Prioridad As String
    <DataMember()>
    Public Avisado As String
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public fechaProgramada As Date
    <DataMember()>
    Public HoraProgramada As TimeSpan
End Class

<DataContract()>
Public Class CTareasPrioridad
    <DataMember()>
    Public id_prioridad As Integer
    <DataMember()>
    Public prioridad As String
End Class

<DataContract()>
Public Class CTareas
    <DataMember()>
    Public id_tarea As Integer
    <DataMember()>
    Public descripcion As String
    <DataMember()>
    Public id_prioridad As Integer
    <DataMember()>
    Public id_usuario As Integer
    <DataMember()>
    Public avisado As String
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public fechaProgramada As Date
    <DataMember()>
    Public HoraProgramada As String
End Class

<DataContract()>
Public Class CTotalesUsuario
    <DataMember()>
    Public ClientesActivos As Integer
    <DataMember()>
    Public ProspectosPorsemana As Integer
    <DataMember()>
    Public ClientesCancelados As Integer
    <DataMember()>
    Public ClientesTotal As Integer

End Class

<DataContract()>
Public Class CTelefonosmodifica
    <DataMember()>
    Public id_telefonoCliente As Integer
    <DataMember()>
    Public Principal As String
    <DataMember()>
    Public Telefono As String
End Class

<DataContract()>
Public Class CidCliente
    Public id_cliente As Integer
    <DataMember()>
    Public Nombre As String
    <DataMember()>
    Public ApellidoPaterno As String
    <DataMember()>
    Public ApellidoMaterno As String
    <DataMember()>
    Public Email As String
    <DataMember()>
    Public id_producto As Integer
    <DataMember()>
    Public id_nivel As Integer
    <DataMember()>
    Public id_empresa As Integer
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public id_etapaActual As Integer
    <DataMember()>
    Public id_campaña As Integer
    <DataMember()>
    Public id_usuarioOriginal As Integer
    <DataMember()>
    Public Observaciones As String
    <DataMember()>
    Public fotografia As String
    <DataMember()>
    Public fotoTpresentacion As String
    <DataMember()>
    Public fechaUltimaVisita As Date
    <DataMember()>
    Public NSS As String
    <DataMember()>
    Public CURP As String
    <DataMember()>
    Public RFC As String
    <DataMember()>
    Public EdoCivil As String
    <DataMember()>
    Public fechaNacimiento As Date
End Class

<DataContract()>
Public Class COperacionesCliente
    <DataMember()>
    Public id_etapa As Integer
    <DataMember()>
    Public Etapa As String
    <DataMember()>
    Public usuario As String
    <DataMember()>
    Public FechaInicio As Date
    <DataMember()>
    Public Observaciones As String
    <DataMember()>
    Public Producto As String
End Class

<DataContract()>
Public Class CEtapasDetalles
    <DataMember()>
    Public id_operacion As Integer
    <DataMember()>
    Public FechaInicio As Date
    <DataMember()>
    Public Observaciones As String
    <DataMember()>
    Public usuario As String
    <DataMember()>
    Public nombre As String
    <DataMember()>
    Public Descripcion As String
    <DataMember()>
    Public NombreCorto As String
End Class

<DataContract()>
Public Class CLlamadasCliente
    <DataMember()>
    Public id_llamada As Integer
    <DataMember()>
    Public usuario As String
    <DataMember()>
    Public Fecha As Date
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public HoraProgramacion As TimeSpan
    <DataMember()>
    Public Programada As String
    <DataMember()>
    Public AvisoCliente As String
    <DataMember()>
    Public AvisoUsuario As String
    <DataMember()>
    Public realizada As String
    <DataMember()>
    Public ObservacionUsuario As String
    <DataMember()>
    Public ObservacionCliente As String
    <DataMember()>
    Public Calificacion As String
End Class

<DataContract()>
Public Class CNombresCliente
    <DataMember()>
    Public id_cliente As Integer
    <DataMember()>
    Public Cliente As String
End Class

<DataContract()>
Public Class CCiudades
    <DataMember()>
    Public id As Integer
    <DataMember()>
    Public nombre As String
End Class

<DataContract()>
Public Class CEstados
    <DataMember()>
    Public id As Integer
    <DataMember()>
    Public nombre As String
End Class

<DataContract()>
Public Class CComboCampañas
    <DataMember()>
    Public id_campaña As Integer
    <DataMember()>
    Public Campaña As String
End Class

<DataContract()>
Public Class CComboEmpresas
    <DataMember()>
    Public id_empresa As Integer
    <DataMember()>
    Public Empresa As String
End Class

<DataContract()>
Public Class CComboProductos
    <DataMember()>
    Public id_producto As Integer
    <DataMember()>
    Public NombreCorto As String
End Class

<DataContract()>
Public Class CUsuarios
    <DataMember()>
    Public id_usuario As Integer = 0
    <DataMember()>
    Public nombre As String = "-"
    <DataMember()>
    Public apellidoPaterno As String = "-"
    <DataMember()>
    Public apellidoMaterno As String = "-"
    <DataMember()>
    Public Email As String = "-"
    <DataMember()>
    Public usuario As String = "-"
    <DataMember()>
    Public contraseña As String = "-"
    <DataMember()>
    Public fechaCreacion As Date = New Date
    <DataMember()>
    Public fotografia As String = "-"
    <DataMember()>
    Public Nivel As Integer = 0
    <DataMember()>
    Public BorraEk As Integer
End Class

<DataContract()>
Public Class CTipoReferencia
    <DataMember()>
    Public id_tiporeferencia As Integer
    <DataMember()>
    Public tiporeferencia As String
End Class

<DataContract()>
Public Class CTipoCampaña
    <DataMember()>
    Public id_tipoCampaña As Integer
    <DataMember()>
    Public TipoCampaña As String
End Class

<DataContract()>
Public Class CTelefonoUsuario
    <DataMember()>
    Public id_telefonoUsuario As Integer
    <DataMember()>
    Public Principal As String
    <DataMember()>
    Public id_usuario As Integer
    <DataMember()>
    Public Telefono As String
End Class

<DataContract()>
Public Class CTelefonoSupervisor
    <DataMember()>
    Public id_telefonoSupervisor As Integer
    <DataMember()>
    Public Principal As String
    <DataMember()>
    Public id_supervisor As Integer
    <DataMember()>
    Public Telefono As String
End Class

Public Class CTelefonoReferencia
    <DataMember()>
    Public id_telefonoReferencia As Integer
    <DataMember()>
    Public Principal As String
    <DataMember()>
    Public id_referencia As Integer
    <DataMember()>
    Public Telefono As String
End Class

<DataContract()>
Public Class CTelefonoEmpresa
    <DataMember()>
    Public id_telefonoEmpresa As Integer
    <DataMember()>
    Public Principal As String
    <DataMember()>
    Public id_empresa As Integer
    <DataMember()>
    Public telefono As String
End Class

<DataContract()>
Public Class CTelefonoContactoEmpresa
    <DataMember()>
    Public id_telefonoContacto As Integer
    <DataMember()>
    Public id_contactoEmpresa As Integer
    <DataMember()>
    Public principal As String
    <DataMember()>
    Public Telefono As String
End Class

<DataContract()>
Public Class CTelefonoCliente
    <DataMember()>
    Public id_telefonoCliente As Integer
    <DataMember()>
    Public Principal As String
    <DataMember()>
    Public id_cliente As Integer
    <DataMember()>
    Public Telefono As String
End Class

<DataContract()>
Public Class CRelacionSupervisorUsuario
    <DataMember()>
    Public id_supervisorusuario As Integer
    <DataMember()>
    Public id_supervisor As Integer
    <DataMember()>
    Public nombre As String
    <DataMember()>
    Public apellidoPaterno As String
    <DataMember()>
    Public apellidoMaterno As String
    <DataMember()>
    Public email As String
    <DataMember()>
    Public id_usuario As Integer
    <DataMember()>
    Public nombre1 As String
    <DataMember()>
    Public apellidoPaterno1 As String
    <DataMember()>
    Public apellidoMaterno1 As String
    <DataMember()>
    Public Email1 As String
End Class

<DataContract()>
Public Class CSupervisorUsuario
    <DataMember()>
    Public id_supervisorusuario As Integer
    <DataMember()>
    Public id_usuario As Integer
    <DataMember()>
    Public id_supervisor As Integer
End Class

<DataContract()>
Public Class CSupervisores
    <DataMember()>
    Public id_supervisor As Integer
    <DataMember()>
    Public nombre As String
    <DataMember()>
    Public apellidoPaterno As String
    <DataMember()>
    Public apellidoMaterno As String
    <DataMember()>
    Public email As String
    <DataMember()>
    Public usuario As String
    <DataMember()>
    Public contraseña As String
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public fotografia As String
    <DataMember()>
    Public NombreCompleto As String
End Class

<DataContract()>
Public Class CRubros
    <DataMember()>
    Public id_rubro As Integer
    <DataMember()>
    Public rubro As String
End Class

<DataContract()>
Public Class CReferenciasCliente
    <DataMember()>
    Public id_referencia As Integer
    <DataMember()>
    Public Nombre As String
    <DataMember()>
    Public ApellidoPaterno As String
    <DataMember()>
    Public ApellidoMaterno As String
    <DataMember()>
    Public email As String
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public tiporeferencia As String
    <DataMember()>
    Public Observaciones As String
    <DataMember()>
    Public fotografia As String
    <DataMember()>
    Public fotoTPresentacion As String
End Class

<DataContract()>
Public Class CReferencias
    <DataMember()>
    Public id_referencia As Integer
    <DataMember()>
    Public id_cliente As Integer
    <DataMember()>
    Public Nombre As String
    <DataMember()>
    Public ApellidoPaterno As String
    <DataMember()>
    Public ApellidoMaterno As String
    <DataMember()>
    Public email As String
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public id_tiporeferencia As Integer
    <DataMember()>
    Public id_usuario As Integer
    <DataMember()>
    Public Observaciones As String
    <DataMember()>
    Public fotografia As String
    <DataMember()>
    Public fotoTPresentacion As String
End Class

<DataContract()>
Public Class CProductosDetalles
    <DataMember()>
    Public id_producto As Integer
    <DataMember()>
    Public NombreCorto As String
    <DataMember()>
    Public Descripcion As String
    <DataMember()>
    Public PrecioNormal As Integer
    <DataMember()>
    Public PrecioDescuento As Integer
    <DataMember()>
    Public categoria As String
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public Observaciones As String
    <DataMember()>
    Public fotografia As String
End Class

<DataContract()>
Public Class CProductos
    <DataMember()>
    Public id_producto As Integer
    <DataMember()>
    Public NombreCorto As String
    <DataMember()>
    Public NombreCompleto As String
    <DataMember()>
    Public Descripcion As String
    <DataMember()>
    Public PrecioNormal As Integer
    <DataMember()>
    Public PrecioDescuento As Integer
    <DataMember()>
    Public id_categoria As Integer
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public Observaciones As String
End Class

<DataContract()>
Public Class CNivelInteres
    <DataMember()>
    Public id_nivelInteres As Integer
    <DataMember()>
    Public nivelinteres As String
End Class

<DataContract()>
Public Class CLlamadas
    <DataMember()>
    Public id_llamada As Integer
    <DataMember()>
    Public id_cliente As Integer
    <DataMember()>
    Public id_usuario As Integer
    <DataMember()>
    Public Fecha As Date
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public HoraProgramacion As String
    <DataMember()>
    Public Programada As String
    <DataMember()>
    Public AvisoCliente As String
    <DataMember()>
    Public AvisoUsuario As String
    <DataMember()>
    Public realizada As String
    <DataMember()>
    Public ObservacionUsuario As String
    <DataMember()>
    Public ObservacionCliente As String
End Class

<DataContract()>
Public Class CEtapasCliente
    <DataMember()>
    Public id_etapa As Integer
    <DataMember()>
    Public nEtapa As Integer
    <DataMember()>
    Public Descripcion As String
End Class

<DataContract()>
Public Class COperaciones
    <DataMember()>
    Public id_operacion As Integer
    <DataMember()>
    Public id_cliente As Integer
    <DataMember()>
    Public id_usuario As Integer
    <DataMember()>
    Public id_etapa As Integer
    <DataMember()>
    Public FechaInicio As Date
    <DataMember()>
    Public FechaFinal As Date
    <DataMember()>
    Public Observaciones As String
End Class

<DataContract()>
Public Class CEmpresasDetalles
    <DataMember()>
    Public id_empresa As Integer
    <DataMember()>
    Public Empresa As String
    <DataMember()>
    Public Razon_Social As String
    <DataMember()>
    Public Direccion As String
    <DataMember()>
    Public PaginaWEb As String
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public Horario As String
    <DataMember()>
    Public rubro As String
    <DataMember()>
    Public Ciudad As String
    <DataMember()>
    Public Estado As String
    <DataMember()>
    Public email As String
    <DataMember()>
    Public Observaciones As String
    <DataMember()>
    Public logotipo As String
End Class

<DataContract()>
Public Class CEmpresas
    <DataMember()>
    Public id_empresa As Integer
    <DataMember()>
    Public Empresa As String
    <DataMember()>
    Public Razon_Social As String
    <DataMember()>
    Public Direccion As String
    <DataMember()>
    Public PaginaWEb As String
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public Horario As String
    <DataMember()>
    Public id_rubro As Integer
    <DataMember()>
    Public id_ciudad As Integer
    <DataMember()>
    Public email As String
    <DataMember()>
    Public Observaciones As String
    <DataMember()>
    Public logotipo As String
End Class

<DataContract()>
Public Class CContactoEmpresa
    <DataMember()>
    Public id_contactoEmpresa As Integer
    <DataMember()>
    Public id_empresa As Integer
    <DataMember()>
    Public Nombre As String
    <DataMember()>
    Public ApellidoPaterno As String
    <DataMember()>
    Public ApellidoMaterno As String
    <DataMember()>
    Public Email As String
    <DataMember()>
    Public Observaciones As String
    <DataMember()>
    Public fotografia As String
End Class

<DataContract()>
Public Class CConfiguraciones
    <DataMember()>
    Public id_configuracion As Integer
    <DataMember()>
    Public diasDeGracias As Integer
    <DataMember()>
    Public emailSistema As String
    <DataMember()>
    Public contraseñaEmail As String
    <DataMember()>
    Public smtpServer As String
    <DataMember()>
    Public puertoEmail As Integer
    <DataMember()>
    Public SSL As String
    <DataMember()>
    Public EnviarEmails As String
End Class

<DataContract()>
Public Class CDetallesCitaUsuario
    <DataMember()>
    Public id_cita As Integer
    <DataMember()>
    Public id_cliente As Integer
    <DataMember()>
    Public Nombre As String
    <DataMember()>
    Public ApellidoPaterno As String
    <DataMember()>
    Public ApellidoMaterno As String
    <DataMember()>
    Public Fecha As Date
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public HoraProgramacion As TimeSpan
    <DataMember()>
    Public Programada As String
    <DataMember()>
    Public AvisoCliente As String
    <DataMember()>
    Public AvisoUsuario As String
    <DataMember()>
    Public realizada As String
    <DataMember()>
    Public ObservacionUsuario As String
    <DataMember()>
    Public ObservacionCliente As String
    <DataMember()>
    Public HoraTermino As TimeSpan
    <DataMember()>
    Public Lugar As String
    <DataMember()>
    Public ConfimacionCliente As String
End Class

<DataContract()>
Public Class CCitas
    <DataMember()>
    Public id_cita As Integer
    <DataMember()>
    Public id_cliente As Integer
    <DataMember()>
    Public id_usuario As Integer
    <DataMember()>
    Public Fecha As Date
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public HoraProgramacion As Date
    <DataMember()>
    Public Programada As String
    <DataMember()>
    Public AvisoCliente As String
    <DataMember()>
    Public AvisoUsuario As String
    <DataMember()>
    Public realizada As String
    <DataMember()>
    Public ObservacionUsuario As String
    <DataMember()>
    Public ObservacionCliente As String
    <DataMember()>
    Public HoraTermino As Date
    <DataMember()>
    Public Lugar As String
    <DataMember()>
    Public ConfimacionCliente As String
End Class

<DataContract()>
Public Class CCategoriasProducto
    <DataMember()>
    Public id_categoria As Integer
    <DataMember()>
    Public categoria As String
End Class

<DataContract()>
Public Class CCampaña
    <DataMember()>
    Public id_campaña As Integer
    <DataMember()>
    Public campañaNombre As String
    <DataMember()>
    Public id_tipoCampaña As Integer
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public fechaInicio As Date
    <DataMember()>
    Public fechaFinal As Date
    <DataMember()>
    Public Observaciones As String
End Class

<DataContract()>
Public Class CCampañaDetalles
    <DataMember()>
    Public id_campaña As Integer
    <DataMember()>
    Public campañaNombre As String
    <DataMember()>
    Public tipoCampaña As String
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public fechaInicio As Date
    <DataMember()>
    Public fechaFinal As Date
    <DataMember()>
    Public Observaciones As String
End Class

<DataContract()>
Public Class CClientesDetalles
    <DataMember()>
    Public id_cliente As Integer
    <DataMember()>
    Public Nombre As String
    <DataMember()>
    Public ApellidoPaterno As String
    <DataMember()>
    Public ApellidoMaterno As String
    <DataMember()>
    Public Email As String
    <DataMember()>
    Public NombreCorto As String
    <DataMember()>
    Public nivelinteres As String
    <DataMember()>
    Public Empresa As String
    <DataMember()>
    Public Id_Campaña As Integer
    <DataMember()>
    Public campañaNombre As String
    <DataMember()>
    Public tipoCampana As String
    <DataMember()>
    Public id_Usuario As Integer
    <DataMember()>
    Public NombreAsesor As String
    <DataMember()>
    Public ApellidoAsesor As String
    <DataMember()>
    Public Observaciones As String
    <DataMember()>
    Public fotografia As String
    <DataMember()>
    Public NSS As String
    <DataMember()>
    Public CURP As String
    <DataMember()>
    Public fotoTpresentacion As String
    <DataMember()>
    Public ranking As String
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public fechaNacimiento As Date
    <DataMember()>
    Public id_etapaActual As Integer
    <DataMember()>
    Public id_producto As Integer
    <DataMember()>
    Public Etapa As String
    <DataMember()>
    Public Numcte As Integer
    <DataMember()>
    Public Numcte2 As Integer
    <DataMember()>
    Public FechaCierre As Date
    <DataMember()>
    Public FechaEscritura As Date
    <DataMember()>
    Public FechaCancelacion As Date
    <DataMember()>
    Public Fecha_Recuperacion As Date
    <DataMember()>
    Public ModeloEk As String
    <DataMember()>
    Public Fecha_OperacionEK As Date
    <DataMember()>
    Public EmpresaEK As Integer
End Class

<DataContract()>
Public Class CClientesTipoCredito
    <DataMember()>
    Public TipoCredito As String
End Class

<DataContract()>
Public Class CClienteTelefonos
    <DataMember()>
    Public Telefono As String
End Class

<DataContract()>
Public Class CClienteObservaciones
    <DataMember()>
    Public Observacion As String
    <DataMember()>
    Public Fecha_Registro As DateTime

End Class

<DataContract()>
Public Class CClientes
    <DataMember()>
    Public id_cliente As Integer
    <DataMember()>
    Public Nombre As String
    <DataMember()>
    Public ApellidoPaterno As String
    <DataMember()>
    Public ApellidoMaterno As String
    <DataMember()>
    Public Email As String
    <DataMember()>
    Public id_producto As Integer
    <DataMember()>
    Public id_nivel As Integer
    <DataMember()>
    Public id_empresa As Integer
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public id_etapaActual As Integer
    <DataMember()>
    Public id_campaña As Integer
    <DataMember()>
    Public id_usuarioOriginal As Integer
    <DataMember()>
    Public Observaciones As String
    <DataMember()>
    Public fotografia As String
    <DataMember()>
    Public fotoTpresentacion As String
End Class

<DataContract()>
Public Class AsesorCallCenter
    <DataMember()>
    Public id_usuario As Integer
    <DataMember()>
    Public nombre As String
    <DataMember()>
    Public apellidoPaterno As String
    <DataMember()>
    Public apellidoMaterno As String
End Class

#Region "Usuarios"
<DataContract()>
Public Class TipoUsuario
    <DataMember()>
    Public id_tipoUsuario As Integer
    <DataMember()>
    Public Tipo As String
End Class
#End Region

#Region "Reportes"
<DataContract()>
Public Class ReporteCalidad
    <DataMember()>
    Public ClienteID As String
    <DataMember()>
    Public ClienteIDCRM As Integer
    <DataMember()>
    Public Numcte As String
    <DataMember()>
    Public Plaza As String
    <DataMember()>
    Public Fraccionamiento As String
    <DataMember()>
    Public Proyecto As String
    <DataMember()>
    Public Mza As String
    <DataMember()>
    Public Lote As String
    <DataMember()>
    Public TipoFraccionamiento As String
    <DataMember()>
    Public EsqFraccionamiento As String
    <DataMember()>
    Public Cliente As String
    <DataMember()>
    Public Contrato As String
    <DataMember()>
    Public FechaEntrega As String
    <DataMember()>
    Public Telefono1 As String
    <DataMember()>
    Public Telefono2 As String
    <DataMember()>
    Public Empresa As String
    <DataMember()>
    Public Departamento As String
    <DataMember()>
    Public Conyuge As String
    <DataMember()>
    Public Asesor As String
    <DataMember()>
    Public Integrador As String
    <DataMember()>
    Public Titular As String
    <DataMember()>
    Public Responsable As String
    <DataMember()>
    Public Ranqueo As String
    <DataMember()>
    Public Visita As Date
End Class

<DataContract()>
Public Class Proyectos
    <DataMember()>
    Public Proyecto As String
End Class

<DataContract()>
Public Class DatosCumplimiento
    <DataMember()>
    Public semana As Integer
    <DataMember()>
    Public asesor As String
    <DataMember()>
    Public cliente As String
    <DataMember()>
    Public prototipo As String
    <DataMember()>
    Public fraccionamiento As String
    <DataMember()>
    Public fechaInicio As Date
    <DataMember()>
    Public etapa As String
    <DataMember()>
    Public nombreCampaña As String
End Class

<DataContract()>
Public Class DetallesCumplimiento
    <DataMember()>
    Public semana As Integer
    <DataMember()>
    Public asesor As String
    <DataMember()>
    Public cliente As String
    <DataMember()>
    Public prototipo As String
    <DataMember()>
    Public fraccionamiento As String
    <DataMember()>
    Public fechaInicio As Date
    <DataMember()>
    Public etapa As String
    <DataMember()>
    Public nombreCampaña As String
    <DataMember()>
    Public Ranking As String
End Class

<DataContract()>
Public Class TotalesCumplimiento
    <DataMember()>
    Public Cantidad As Integer
    <DataMember()>
    Public Campana As String
    <DataMember()>
    Public TipoCampana As String
End Class

<DataContract()>
Public Class ReporteArpa
    <DataMember()>
    Public Prospectos As Integer
    <DataMember()>
    Public Visitas As Integer
    <DataMember()>
    Public Separaciones As Integer
End Class

<DataContract()>
Public Class ProspeccionesCRM
    <DataMember()>
    Public idCliente As Integer
    <DataMember()>
    Public nombreCliente As String
    <DataMember()>
    Public idUsuario As Integer
    <DataMember()>
    Public nombreUsuario As String
    <DataMember()>
    Public idSupervisor As Integer
    <DataMember()>
    Public nombreSupervisor As String
    <DataMember()>
    Public fechaCreacion As Date
End Class

<DataContract()>
Public Class VisitasCRM
    <DataMember()>
    Public idCliente As Integer
    <DataMember()>
    Public nombreCliente As String
    <DataMember()>
    Public idUsuario As Integer
    <DataMember()>
    Public nombreUsuario As String
    <DataMember()>
    Public idSupervisor As Integer
    <DataMember()>
    Public nombreSupervisor As String
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public fechaUltimaVisita As Date
End Class

<DataContract()>
Public Class SeparacionesCRM
    <DataMember()>
    Public idCliente As Integer
    <DataMember()>
    Public nombreCliente As String
    <DataMember()>
    Public idUsuario As Integer
    <DataMember()>
    Public nombreUsuario As String
    <DataMember()>
    Public idSupervisor As Integer
    <DataMember()>
    Public nombreSupervisor As String
    <DataMember()>
    Public fechaCreacion As Date
    <DataMember()>
    Public fechaUltimaVisita As Date
    <DataMember()>
    Public fechaCierre As Date
End Class

<DataContract()>
Public Class CancelacionesEnkontrol
    <DataMember()>
    Public numeroCliente As Integer
    <DataMember()>
    Public nombreCliente As String
    <DataMember()>
    Public numeroAgente As Integer
    <DataMember()>
    Public nombreAgente As String
    <DataMember()>
    Public numeroLider As Integer
    <DataMember()>
    Public nombreLider As String
    <DataMember()>
    Public fechaCancelacion As Date
    <DataMember()>
    Public cc As String
    <DataMember()>
    Public zona As String
    <DataMember()>
    Public empresa As String
End Class

<DataContract()>
Public Class SeparacionesEnkontrol
    <DataMember()>
    Public numeroCliente As Integer
    <DataMember()>
    Public cliente As String
    <DataMember()>
    Public numeroAgente As Integer
    <DataMember()>
    Public agente As String
    <DataMember()>
    Public numeroLider As Integer
    <DataMember()>
    Public nombreLider As String
    <DataMember()>
    Public cc As String
    <DataMember()>
    Public zona As String
    <DataMember()>
    Public empresa As Integer
End Class

<DataContract()>
Public Class ProspeccionesEnkontrol
    <DataMember()>
    Public numeroCliente As Integer
    <DataMember()>
    Public cliente As String
    <DataMember()>
    Public numeroAgente As Integer
    <DataMember()>
    Public agente As String
    <DataMember()>
    Public numeroLider As Integer
    <DataMember()>
    Public nombreLider As String
    <DataMember()>
    Public cc As String
    <DataMember()>
    Public zona As String
    <DataMember()>
    Public empresa As Integer
End Class
#End Region

#Region "Enkontrol"
<DataContract()>
Public Class OrigenAgente
    <DataMember()>
    Public Origen As String
End Class
#End Region

#Region "Citas"
<DataContract()>
Public Class VigenciaCitas
    <DataMember()>
    Public TotalCitas As Integer
    <DataMember()>
    Public CitasVigentes As Integer
    <DataMember()>
    Public Id_Usuario As Integer
    <DataMember()>
    Public UsuarioVigente As String
    <DataMember()>
    Public TipoUsuario As String
End Class
#End Region
