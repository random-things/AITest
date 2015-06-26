VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMySQLConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strServer As String
Private strDatabase As String
Private strUsername As String
Private strPassword As String
Private bConnected As Boolean

Private strLastQuery As String

Private MySQLConnection As ADODB.Connection
Private MySQLRecordSet As ADODB.Recordset

Public Event OnMySQLError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
Public Event OnQuery(ByVal QueryString As String)
Public Event OnRecordSetComplete()
Public Event OnRecordSetColumns(ByVal spaceDelimitedColumns As String)
Public Event OnRecordSetRecord(ByRef Values() As String)

Public Property Get MySQLServer() As String
    MySQLServer = strServer
End Property

Public Property Let MySQLServer(ByVal NewServer As String)
    strServer = NewServer
End Property

Public Property Get LastQuery() As String
    LastQuery = strLastQuery
End Property

Public Property Get Database() As String
    Database = strDatabase
End Property

Public Property Let Database(ByVal newDatabase As String)
    strDatabase = newDatabase
End Property

Public Property Get Username() As String
    Username = strUsername
End Property

Public Property Let Username(ByVal newUsername As String)
    strUsername = newUsername
End Property

Public Property Get Password() As String
    Password = strPassword
End Property

Public Property Let Password(ByVal newPassword As String)
    strPassword = newPassword
End Property

Private Sub Class_Initialize()
    Set MySQLConnection = New ADODB.Connection
    Set MySQLRecordSet = New ADODB.Recordset

    bConnected = False
    strServer = "localhost"
    strDatabase = ""
    strUsername = "root"
    strPassword = ""
End Sub

Public Function Connect() As Boolean
    On Error GoTo ErrorHandler:
    Connect = False

    If MySQLConnection.State = adStateOpen Then MySQLConnection.Close
    
    Dim strConnString As String
    strConnString = "DRIVER={MySQL ODBC 3.51 Driver};"
    strConnString = strConnString & "SERVER=" & strServer
    strConnString = strConnString & ";UID=" & strUsername
    strConnString = strConnString & ";PWD=" & strPassword
    strConnString = strConnString & ";DATABASE=" & strDatabase & ""
    
    Call MySQLConnection.Open(strConnString)
    bConnected = True
    Connect = True
    Exit Function
    
ErrorHandler:
    RaiseEvent OnMySQLError(Err.Number, Err.Description)
End Function

Public Function Disconnect() As Boolean
    Disconnect = False
    
    If MySQLConnection.State <> adStateClosed Then
        bConnected = False
        Disconnect = True
        MySQLConnection.Close
        Exit Function
    End If
End Function

Public Function QueryDELETE(ByVal Table As String, ByVal columnName As String, ByVal columnValue As Variant) As Long
    On Error GoTo ErrorHandler:
    
    Dim QueryString As String
    QueryString = "DELETE FROM " & Table
    QueryString = QueryString & " WHERE " & columnName & "="
    If VarType(columnValue) = vbString Then
        QueryString = QueryString & "'" & columnValue & "'"
    Else
        QueryString = QueryString & columnValue
    End If
    
    strLastQuery = QueryString
    
    RaiseEvent OnQuery(QueryString)
    
    Call MySQLConnection.Execute(QueryString, QueryDELETE)
    
    Exit Function
ErrorHandler:
    RaiseEvent OnMySQLError(Err.Number, Err.Description)
End Function

Public Function QueryINSERT(ByVal Table As String, ParamArray columnsAndValues() As Variant) As Long
    On Error GoTo ErrorHandler:
    
    Dim QueryString As String, i As Integer
    QueryString = "INSERT INTO " & Table & " ("
    For i = LBound(columnsAndValues) To UBound(columnsAndValues) Step 2
        QueryString = QueryString & columnsAndValues(i) & ","
    Next i
    QueryString = Left$(QueryString, Len(QueryString) - 1)
    
    QueryString = QueryString & ") VALUES ("
    For i = LBound(columnsAndValues) + 1 To UBound(columnsAndValues) + 1 Step 2
        If VarType(columnsAndValues(i)) = vbString Then
            If columnsAndValues(i) = "NOW()" Then
                QueryString = QueryString & columnsAndValues(i) & ","
            Else
                QueryString = QueryString & "'" & columnsAndValues(i) & "',"
            End If
        Else
            QueryString = QueryString & columnsAndValues(i) & ","
        End If
    Next i
    QueryString = Left$(QueryString, Len(QueryString) - 1)
    QueryString = QueryString & ")"
    
    strLastQuery = QueryString
    
    RaiseEvent OnQuery(QueryString)
    
    Call MySQLConnection.Execute(QueryString, QueryINSERT)
    Exit Function
ErrorHandler:
    RaiseEvent OnMySQLError(Err.Number, Err.Description)
End Function

Public Sub Query(ByVal QueryString As String)
    On Error GoTo ErrorHandler:
    MySQLConnection.Execute QueryString
    Exit Sub
ErrorHandler:
End Sub

Public Function QuerySELECT(ByVal QueryString As String)
    On Error GoTo ErrorHandler:
    If bConnected = True Then
        strLastQuery = QueryString
        
        RaiseEvent OnQuery(QueryString)
    
        MySQLRecordSet.Open QueryString, MySQLConnection, adOpenStatic, adLockReadOnly
        
        If (MySQLRecordSet.RecordCount <> 0) And (Not MySQLRecordSet.EOF) Then
            MySQLRecordSet.MoveFirst
            
            Dim i As Integer, strColumns As String
            For i = 0 To MySQLRecordSet.Fields.Count - 1
                strColumns = strColumns & MySQLRecordSet.Fields(i).Name & " "
            Next i
            
            strColumns = Left$(strColumns, Len(strColumns) - 1)
            RaiseEvent OnRecordSetColumns(strColumns)
            
            While (Not MySQLRecordSet.EOF)
                ' Loop through the records and return them.
                Dim arrReturn() As String
                ReDim arrReturn(MySQLRecordSet.Fields.Count - 1)
                
                For i = 0 To MySQLRecordSet.Fields.Count - 1
                    arrReturn(i) = MySQLRecordSet.Fields(i).Value
                Next i
                
                RaiseEvent OnRecordSetRecord(arrReturn)
                
                MySQLRecordSet.MoveNext
            Wend
        Else
            'RaiseEvent OnRecordSetComplete
            ' Do nothing
        End If
    End If
    
    MySQLRecordSet.Close
    RaiseEvent OnRecordSetComplete
    Exit Function
    
ErrorHandler:
    RaiseEvent OnMySQLError(Err.Number, Err.Description)
End Function
