Imports System.Data.SqlClient
Public Class SQL_Functions
    Private StrCon As String = ConfigurationManager.ConnectionStrings("Conexion_SQL").ConnectionString

    Private SQLCon_A As New SqlConnection(StrCon)
    Private SQLCon_B As New SqlConnection(StrCon)

    ' Objetos de MYSQL
    Private SQL_DA As New SqlDataAdapter
    Private SQL_CMD As New SqlCommand
    Private SQL_DR As SqlDataReader
    Private SQLTran As SqlTransaction

    Private _bolInTransaction As Boolean

    Enum TipoTransaccion
        OpenCon_BeginTrans = 0 ' OpenCon And BEGIN
        ContCon_Transaction = 1 ' Transaccion acumulada, Coninua
        CloseCon_CommitTrans = 2 ' CierraCon and COMMIT
        UniqueTransaction = 3 ' Transaccion unica y completa
    End Enum

    Public Sub New()
        _bolInTransaction = False
    End Sub

    Public Function SQlGetDataset(ByVal localSQL As String) As DataSet
        Dim SQL_DS As New DataSet

        Try
            SQLCon_A.Open()
            SQL_DA = New SqlDataAdapter(localSQL, SQLCon_A)
            SQL_DA.Fill(SQL_DS)
        Catch ex As Exception

        Finally
            SQlGetDataset = SQL_DS.Copy
            SQLCon_A.Close()
            SQL_DA.Dispose()
            With SQL_DS : .Clear() : .Dispose() : End With
        End Try
    End Function

    Public Function SQLGetTable(ByVal LocalSQL As String) As DataTable
        Dim SQL_DS As New DataSet

        Try
            SQLCon_A.Open()
            SQL_DA = New SqlDataAdapter(LocalSQL, SQLCon_A)
            SQL_DA.Fill(SQL_DS)

            SQLGetTable = SQL_DS.Tables(0).Copy
        Catch ex As Exception
            SQLGetTable = Nothing
        Finally
            SQLCon_A.Close()
            SQL_DA.Dispose()
            With SQL_DS : .Clear() : .Dispose() : End With
        End Try
    End Function

    Public Function SQLGetDataDbl(ByVal LocalSQL As String) As Double
        Try
            With SQL_CMD
                .Connection = SQLCon_A
                .CommandText = LocalSQL

                SQLCon_A.Open()
                SQL_DR = .ExecuteReader
                SQL_DR.Read()
                If SQL_DR.IsDBNull(0) Then SQLGetDataDbl = 0 Else SQLGetDataDbl = SQL_DR.GetValue(0)
            End With

        Catch ex As Exception
            SQLGetDataDbl = 0
        Finally
            If Not (SQL_DR Is Nothing) Then SQL_DR.Close()
            SQLCon_A.Close()
        End Try
    End Function

    Public Function SQLGetDataStr(ByVal LocalSQL As String) As String
        Try
            With SQL_CMD
                .Connection = SQLCon_A
                .CommandText = LocalSQL

                SQLCon_A.Open()
                SQL_DR = .ExecuteReader
                SQL_DR.Read()
                If SQL_DR.IsDBNull(0) Then SQLGetDataStr = "" Else SQLGetDataStr = "" & SQL_DR.GetString(0)
            End With

        Catch ex As Exception
            SQLGetDataStr = "E$ " & ex.Message.ToString
        Finally
            If Not (SQL_DR Is Nothing) Then SQL_DR.Close()
            SQLCon_A.Close()
        End Try
    End Function

    Public Sub SQLBeginTransaction(Optional ByVal OpenConn As Boolean = True)
        If Not _bolInTransaction Then
            If SQLCon_B.State = ConnectionState.Closed And OpenConn Then SQLCon_B.Open()

            SQLTran = SQLCon_B.BeginTransaction
            _bolInTransaction = True
        End If
    End Sub

    Public Sub SQLCommitTransaction(Optional ByVal CloseConn As Boolean = False)
        With SQLTran
            If _bolInTransaction Then .Commit() : .Dispose() : _bolInTransaction = False

            If CloseConn Then SQLCon_B.Close()
        End With
    End Sub

    Public Sub SQLRollBack(Optional ByVal CloseConn As Boolean = False)
        With SQLTran
            If _bolInTransaction Then .Rollback() : .Dispose() : _bolInTransaction = False

            If CloseConn Then SQLCon_B.Close()
        End With
    End Sub

    Public Function SQLInTransaction() As Boolean
        SQLInTransaction = _bolInTransaction
    End Function

    Public Function SQLExecSQL(ByVal LocalSQL As String, ByVal TransactionStep As TipoTransaccion) As Boolean
        ' TransactionStep 
        '  0 = OpenCon And BEGIN
        '  1 = Transaccion acumulada, Coninua
        '  2 = CierraCon and COMMIT
        '  3 = Transaccion unica y completa

        Try
            Select Case TransactionStep
                Case TipoTransaccion.OpenCon_BeginTrans, TipoTransaccion.UniqueTransaction '0, 3
                    If _bolInTransaction Then
                        If TransactionStep = TipoTransaccion.UniqueTransaction Then
                            Throw New Exception("No puede activarse una transaccion unica (3) dentro de una transaccion anidada previa (BeginTrans)")
                        End If
                    End If

                    SQLBeginTransaction()
                    ' Si es una transaccion unica, entonces quitar bandera de transa global
                    If TransactionStep = TipoTransaccion.UniqueTransaction Then _bolInTransaction = False

                    'If MySQLCon_B.State = ConnectionState.Closed Then MySQLCon_B.Open()

                    '' Abrir la transaccion si es necesaria
                    'If Not _bolInTransaction Then MYSQLTran = MySQLCon_B.BeginTransaction
                Case Else ' 1 y 2 Continuacion de execute SQL's, y transaccion abierta

            End Select

            With SQL_CMD
                .Connection = SQLCon_B
                .CommandText = LocalSQL
                .Transaction = SQLTran
                .ExecuteNonQuery()
            End With

            ' Cerrar la transaccion si se pide
            Select Case TransactionStep
                Case TipoTransaccion.CloseCon_CommitTrans, TipoTransaccion.UniqueTransaction '2, 3 ' Transaccion unica terminacion obligada
                    ' Si es una transaccion Unica entonces activar bandera de transa global para terminar transaccion
                    If TransactionStep = TipoTransaccion.UniqueTransaction Then _bolInTransaction = True
                    If _bolInTransaction Then SQLCommitTransaction()

                    SQLExecSQL = True
                    If SQLCon_B.State = ConnectionState.Open Then SQLCon_B.Close()
                    SQL_CMD.Dispose()

                Case Else ' Case 1 (Transaccion incompleta y conexion abierta(mas querys))
                    SQLExecSQL = True
            End Select


        Catch ex As Exception
            SQLExecSQL = False
            ' Si esta en transaccion global, entonces no cerrar la conexion
            Select Case TransactionStep
                Case TipoTransaccion.OpenCon_BeginTrans, TipoTransaccion.ContCon_Transaction, TipoTransaccion.CloseCon_CommitTrans '0, 1, 2

                Case TipoTransaccion.UniqueTransaction '3
                    If Not _bolInTransaction Then
                        SQLTran.Rollback()
                        SQLCon_B.Close()
                    End If
            End Select
            SQL_CMD.Dispose()
        End Try
    End Function
End Class
