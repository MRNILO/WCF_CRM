#Region "Imports"
Imports System.Data
Imports System.Data.Odbc
Imports System.Configuration

#End Region

Public Class DirectConn
#Region "ConexionSQLAnywhere"
    ' Apuntar al ODBC definido en la PC
    Dim ODBCconStr As String

    Dim ODBCcon_A As IDbConnection = New OdbcConnection(ODBCconStr)
    Dim ODBCcon_B As IDbConnection = New OdbcConnection(ODBCconStr)

    Dim ODBC_CMD As IDbCommand = ODBCcon_A.CreateCommand()
    Dim ODBC_DA As IDbDataAdapter = New OdbcDataAdapter
    Dim ODBC_DR As OdbcDataReader

    Public Function ODBCGetDataset(ByVal LocalSQL As String, ByVal NumEmp As Integer) As DataSet
        Dim ODBC_DS As New DataSet

        Try
            ConsODBCConStr(NumEmp)

            ODBCcon_A.Open()
            ODBC_CMD.CommandText = LocalSQL
            ODBC_DA.SelectCommand = ODBC_CMD
            ODBC_DA.Fill(ODBC_DS)

            ODBCcon_A.Close()
        Catch ex As Exception

        Finally
            ODBCGetDataset = ODBC_DS.Copy
            With ODBC_DS : .Clear() : .Dispose() : End With
        End Try
    End Function

    Public Function ODBCGetDataDbl(ByVal LocalSQL As String, ByVal NumEmp As Integer) As Double
        Try
            ConsODBCConStr(NumEmp)

            With ODBC_CMD
                .Connection = ODBCcon_A
                .CommandText = LocalSQL

                ODBCcon_A.Open()
                ODBC_DR = .ExecuteReader
                ODBC_DR.Read()
                ODBCGetDataDbl = Val("0" & ODBC_DR.GetDouble(0))
            End With

        Catch ex As Exception
            ODBCGetDataDbl = -9999999

        Finally
            ODBC_DR.Close()
            ODBCcon_A.Close()
        End Try
    End Function

    Public Sub ConsODBCConStr(ByVal NumEmp As Integer)
        If NumEmp > 0 Then
            Select Case NumEmp
                Case 1 To 9
                    ODBCconStr = "DSN=EK_ADM" & Format(NumEmp, "00")
                Case 11 To 99
                    ODBCconStr = "DSN=EK_ADM" & Format(NumEmp, "00") & "_11"
            End Select
        Else
            ' DEFAULT
            ODBCconStr = "DSN=EK_ADM11_11"
        End If
        ODBCconStr += ";UID=user_read;PWD=lectura"

        ODBCcon_A.ConnectionString = ODBCconStr
        ODBCcon_B.ConnectionString = ODBCconStr
    End Sub

    Public Function ODBCExecSQL(ByVal LocalSQL As String, ByVal NumEmp As Integer) As Boolean
        Try
            ConsODBCConStr(NumEmp)

            ODBCcon_A.Open()
            With ODBC_CMD
                .Connection = ODBCcon_A
                .CommandText = LocalSQL
                .ExecuteNonQuery()
            End With

            ODBCExecSQL = True
        Catch ex As Exception
            ODBCExecSQL = False
        Finally
            If Not (ODBC_DR Is Nothing) Then ODBC_DR.Close()
            ODBCcon_A.Close()
        End Try
    End Function
#End Region
End Class
