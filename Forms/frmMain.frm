VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckIRC 
      Left            =   240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents IRC As clsIRC
Attribute IRC.VB_VarHelpID = -1

Private WithEvents SQL As cMySQLConnection
Attribute SQL.VB_VarHelpID = -1

Private Biggie As clsBigramArray

Private g_Command As String
Private g_User As String
Private g_User2 As String
Private g_Dest As String
Private g_lineCount As Long
Private g_Tag As String
Private g_ID As String
Private g_IDs() As String
Private g_IDIndex As Long

Private Sub Form_Load()
    Randomize
    
    ReDim g_IDs(0)
    
    Set Biggie = New clsBigramArray

    Set SQL = New cMySQLConnection
    Set IRC = New clsIRC
    
    IRC.Socket = sckIRC
    IRC.Server = "some.irc.server.com"
    IRC.Host = "my.host"
    IRC.Name = "AI Bot"
    IRC.Nick = "AIBot"
    
    IRC.IRCConnect
    
    SQL.MySQLServer = "localhost"
    SQL.Database = "irc"
    SQL.Username = "root"
    SQL.Password = "password"
    
    SQL.Connect
End Sub

Private Sub IRC_OnConnecting(ByVal Server As String, ByVal Port As String)
    Debug.Print "Connecting to " & Server
End Sub

Private Sub IRC_OnLoggedIn()
    Debug.Print "Logged in to IRC server."
    IRC.JoinChannel "#beta"
End Sub

Private Sub IRC_OnPrivateMessage(ByVal Destination As String, ByVal Nickname As String, ByVal UserHost As String, ByVal Message As String)
    'IRC.PrivateMessage Destination, Nickname & "(" & UserHost & ") said: " & Message
    
    If (Nickname = "TehUser") And (Left$(Message, 1) = ".") Then
        Dim strMessage() As String
        strMessage = Split(Message, " ")
        
        If strMessage(0) = ".join" Then
            IRC.JoinChannel strMessage(1)
        End If
        
        If strMessage(0) = ".part" Then
            IRC.LeaveChannel strMessage(1)
        End If
        
        If strMessage(0) = ".sql" Then
            SQL.Query Right$(Message, Len(Message) - 5)
        End If
    End If
    
    If Left$(Message, 1) = "." Then
        strMessage = Split(Message, " ")
        
        If strMessage(0) = ".test" And UBound(strMessage) > 0 Then
            g_Command = "test"
            g_User = Replace(Replace(strMessage(1), "\", ""), "'", "\'")
            g_Dest = Destination
            g_lineCount = 0
            SQL.QuerySELECT "SELECT line FROM irc_lines WHERE wordcount>2 AND username='" & g_User & "'"
        End If
        
        If strMessage(0) = ".test2" And UBound(strMessage) > 1 Then
            g_Command = "test2"
            g_User = Replace(Replace(strMessage(1), "\", ""), "'", "\'")
            g_User2 = Replace(Replace(strMessage(2), "\", ""), "'", "\'")
            g_Dest = Destination
            g_lineCount = 0
            SQL.QuerySELECT "SELECT line FROM irc_lines WHERE wordcount>2 AND ((username='" & g_User & "') OR (username='" & g_User2 & "'))"
        End If
        
        If strMessage(0) = ".answer" Then
            IRC.PrivateMessage Destination, IIf(getRand(2), "Yes.", "No.")
        End If
        
        If strMessage(0) = ".time" Then
            IRC.PrivateMessage Destination, "Time: " & Format(DateAdd("h", 7, DateTime.Now), "yyyy-mm-dd hh:mm:ss")
        End If
        
        If strMessage(0) = ".5" Then
            If UBound(strMessage) > 0 Then
                g_Command = "5"
                g_Dest = Destination
                g_lineCount = 0
                SQL.QuerySELECT "SELECT id,username,channel,line FROM irc_lines WHERE username='" & strMessage(1) & "' ORDER BY tstamp DESC LIMIT 5"
            Else
                g_Command = "5"
                g_Dest = Destination
                g_lineCount = 0
                SQL.QuerySELECT "SELECT id,username,channel,line FROM irc_lines ORDER BY tstamp DESC LIMIT 5"
            End If
        End If
        
        If strMessage(0) = ".tag" Then
            If UBound(strMessage) = 2 Then
                g_Command = "tag"
                g_Dest = Destination
                g_Tag = strMessage(2)
                g_ID = strMessage(1)
                
                SQL.QuerySELECT "SELECT COUNT(*) FROM irc_lines WHERE id=" & g_ID
            End If
        End If
        
        If strMessage(0) = ".random" Then
            If UBound(strMessage) = 1 Then
                g_Command = "random"
                g_Dest = Destination
                g_Tag = strMessage(1)
                g_lineCount = 0
                g_IDIndex = 0
                
                SQL.QuerySELECT "SELECT line_id FROM tags WHERE tag='" & strMessage(1) & "'"
            End If
        End If
        If strMessage(0) = ".playback" Then
            g_User = Nickname
            g_Dest = Destination
            g_Command = "playback"
            Dim tmpTime As Date
            tmpTime = DateAdd("n", 0 - CInt(strMessage(1)), DateAdd("h", 7, DateTime.Now))
            SQL.QuerySELECT "SELECT channel,username,line,tstamp FROM irc_lines WHERE (tstamp>'" & Format(tmpTime, "yyyy-mm-dd hh:mm:ss") & "')"
        End If
    Else
        While InStr(Message, "  ") > 0
            Message = Replace(Message, "  ", " ")
        Wend
        
        Message = Replace(Message, "'", "\'")
        
        Dim wordCount As Integer
        wordCount = UBound(Split(Message, " "))
        
        If (Left$(Message, 1) <> ".") And (Left$(Message, 1) <> "[") And (Left$(Message, 1) <> "<") Then
            SQL.QueryINSERT "irc_lines", "username", Replace(Nickname, "'", "\'"), "line", Replace(Replace(Message, "\", ""), "'", "\'"), "tstamp", "NOW()", "channel", Destination, "wordcount", wordCount
        End If
    End If

End Sub

Private Sub IRC_OnUserJoined(ByVal Channel As String, ByVal Nickname As String, ByVal UserHost As String)
    If Nickname = IRC.Nick Then
        Debug.Print "-> Joined " & Channel
        IRC.PrivateMessage Channel, "Joined: " & Channel
    End If
End Sub

Private Sub SQL_OnMySQLError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    Debug.Print ErrorNumber & ": " & ErrorDescription
    Debug.Print SQL.LastQuery
End Sub

Private Sub SQL_OnQuery(ByVal QueryString As String)
    'IRC.PrivateMessage "#beta", "Querying -> " & QueryString
    Debug.Print "Querying -> " & QueryString
End Sub

Private Sub SQL_OnRecordSetColumns(ByVal spaceDelimitedColumns As String)
    Debug.Print spaceDelimitedColumns
End Sub

Private Sub SQL_OnRecordSetComplete()
    If g_Command = "test" Then
        If g_lineCount > 0 Then
            If g_User = "brew" Then
                IRC.PrivateMessage g_Dest, "[" & g_lineCount & "/" & g_lineCount & "] " & g_User & ": " & Biggie.buildSentence
            Else
                IRC.PrivateMessage g_Dest, "[" & g_lineCount & "/" & Biggie.wordCount & "] " & g_User & ": " & Biggie.buildSentence
            End If
            g_lineCount = 0
            Biggie.Reset
        Else
            IRC.PrivateMessage g_Dest, g_User & " has no lines in the database."
        End If
    End If
    
    If g_Command = "test2" Then
        If g_lineCount > 0 Then
            IRC.PrivateMessage g_Dest, "[" & g_lineCount & "/" & Biggie.wordCount & "] " & g_User & " and " & g_User2 & ": " & Biggie.buildSentence
            g_lineCount = 0
            Biggie.Reset
        Else
            IRC.PrivateMessage g_Dest, g_User & " and " & g_User2 & " have no lines in the database."
        End If
    End If
    
    If g_Command = "random" Then
        If g_IDIndex > 0 Then
            g_Command = "random2"
            SQL.QuerySELECT "SELECT tstamp,username,channel,line FROM irc_lines WHERE id=" & g_IDs(getRand(g_IDIndex))
            g_IDIndex = 0
        Else
            IRC.PrivateMessage g_Dest, "Quote with tag '" & g_Tag & "' was not found."
        End If
    End If
End Sub

Private Sub SQL_OnRecordSetRecord(Values() As String)
    If g_Command = "test" Or g_Command = "test2" Then
        Dim i As Integer
        'For i = 0 To UBound(Values)
        '    Debug.Print Values(i)
        'Next i
        
        ' Read the lines by character.
        For i = 0 To UBound(Values)
            Values(i) = Replace(Values(i), "\'", "'")
            Biggie.parseLine Values(i)
        Next i
        
        g_lineCount = g_lineCount + 1
    End If
    
    If g_Command = "playback" Then
        'IRC.PrivateMessage g_User, Values(3) & "<" & Values(1) & ":" & Values(0) & "> " & Values(2)
    End If
    
    If g_Command = "5" Then
        IRC.PrivateMessage g_Dest, "ID#" & Values(0) & " <" & Values(1) & ":" & Values(2) & "> " & Values(3)
    End If
    
    If g_Command = "tag" Then
        If Values(0) = 1 Then
            SQL.QueryINSERT "tags", "line_id", g_ID, "tag", g_Tag
        End If
    End If
    
    If g_Command = "random" Then
        If Values(0) > 0 Then
            g_lineCount = g_lineCount + 1
            g_IDs(g_IDIndex) = Values(0)
            g_IDIndex = g_IDIndex + 1
            ReDim Preserve g_IDs(g_IDIndex)
        End If
    End If
    
    If g_Command = "random2" Then
        IRC.PrivateMessage g_Dest, "[Random: " & g_Tag & "] " & Values(0) & " <" & Values(1) & ":" & Values(2) & "> " & Values(3)
        ReDim g_IDs(0)
    End If
End Sub

Private Function parseFile(ByVal Nickname As String) As String
    ' Get the lines for this user.
    SQL.QuerySELECT "SELECT line FROM irc_lines WHERE username='" & Nickname & "'"
End Function

Private Function getRand(ByVal x As Integer) As Integer
    getRand = Int(Rnd() * x)
    If getRand < 0 Then getRand = 0
    'IRC.PrivateMessage g_Dest, "Random(0," & x & ") -> " & getRand
End Function

Private Sub SQL_OnRecordSetRowCount(ByVal rowCount As Long)
    g_rowCount = rowCount
    g_currentRow = 0
End Sub
