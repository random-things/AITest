VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBigramArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_WORDS As Integer = 5000

Private Const FIRST_WORD As Integer = 0
Private Const MIDDLE_WORD As Integer = 1
Private Const LAST_WORD As Integer = 2

Private Const START_SYMBOL As Integer = 0
Private Const END_SYMBOL As Integer = 1

Private Const RAND_MAX As Long = &H7FFFFFF

Private bigramArray(MAX_WORDS, MAX_WORDS) As Long
Private wordList(MAX_WORDS) As String
Private sumVector(MAX_WORDS) As Long

Private loadedWords As Integer
Private g_lastIndex As Integer
Private g_lastWord As String

Public Sub Reset()
    Dim i As Integer, j As Integer
    
    For i = 2 To loadedWords
        sumVector(i) = 0
        wordList(i) = ""
    Next i
    
    For i = 0 To loadedWords
        For j = 0 To loadedWords
            bigramArray(i, j) = 0
        Next j
    Next i
    
    loadedWords = 2
    g_lastIndex = 0
    g_lastWord = ""
End Sub

Public Function wordCount() As Long
    wordCount = loadedWords
End Function

Private Sub Class_Initialize()
    wordList(0) = "<START>"
    wordList(1) = "<END>"
    
    loadedWords = 2
    
    g_lastWord = ""
End Sub

Public Sub parseLine(ByVal line As String)
    line = LCase$(line)
    
    ' Remove excess spacing.
    While InStr(line, "  ") > 0
        line = Replace(line, "  ", " ")
    Wend

    Dim words() As String
    words = Split(line, " ")
    
    Dim windex As Integer
    windex = 0
    
    Call loadWord(words(0), FIRST_WORD)
    While windex < UBound(words)
        windex = windex + 1
        If (Right$(words(windex), 1) = ".") Or (Right$(words(windex), 1) = "!") Or (Right$(words(windex), 1) = "?") Then
            Call loadWord(words(windex), MIDDLE_WORD)
            Call loadWord(words(windex), LAST_WORD)
            windex = windex + 1
            If windex <= UBound(words) Then
                Call loadWord(words(windex), FIRST_WORD)
            End If
        Else
            Call loadWord(words(windex), MIDDLE_WORD)
        End If
    Wend
End Sub

Private Sub loadWord(ByVal word As String, ByVal order As Byte)
    
    If (Right$(word, 1) = "?") Or (Right$(word, 1) = ".") Then
        word = Left$(word, Len(word) - 1)
    End If

    ' See if the word already exists.
    Dim i As Integer
    For i = 2 To loadedWords - 1
        If wordList(i) = word Then
            Exit For
        End If
    Next i
    
    'Debug.Print "i -> " & i
    
    If i = loadedWords Then
        If loadedWords = MAX_WORDS Then
            Exit Sub
        End If
        
        'Debug.Print "loadedWords -> " & loadedWords
        wordList(loadedWords) = word
        loadedWords = loadedWords + 1
    End If
    
    If order = FIRST_WORD Then
        bigramArray(START_SYMBOL, i) = bigramArray(START_SYMBOL, i) + 1
        sumVector(START_SYMBOL) = sumVector(START_SYMBOL) + 1
    ElseIf order = LAST_WORD Then
        bigramArray(i, END_SYMBOL) = bigramArray(i, END_SYMBOL) + 1
        bigramArray(END_SYMBOL, i) = bigramArray(END_SYMBOL, i) + 1
        sumVector(END_SYMBOL) = sumVector(END_SYMBOL) + 1
    Else
        bigramArray(g_lastIndex, i) = bigramArray(g_lastIndex, i) + 1
        sumVector(g_lastIndex) = sumVector(g_lastIndex) + 1
    End If
    
    g_lastIndex = i
End Sub

Public Function buildSentence() As String
    Dim word As Integer
    Dim max As Integer
    
    word = nextWord(word)
    
    While word <> END_SYMBOL
        'Debug.Print "-> " & wordList(word)
        If g_lastWord <> wordList(word) Then
            buildSentence = buildSentence & wordList(word) & " "
            g_lastWord = wordList(word)
        End If
        
        word = nextWord(word)
        'Debug.Print "word is now -> " & word
        
        max = max + getRand(12) + 1
        If max >= 100 Then
            buildSentence = Left$(buildSentence, Len(buildSentence) - 1)
            Exit Function
        End If
    Wend
End Function

Private Function nextWord(ByVal word As Integer) As Integer
    nextWord = (word + 1)
    Dim max As Long
    max = sumVector(word)
    
    Dim lim As Long, sum As Long
    lim = getRand(max) + 1
    
    'Dim tmp As String
    'Dim i As Integer
    'For i = 1 To 10
    '    tmp = tmp & CStr(getRand(max) + 1)
    'Next i
    
    'Debug.Print tmp
    'Debug.Print "max -> " & max & " / lim -> " & lim
    
    While nextWord <> word
        nextWord = nextWord Mod loadedWords
        
        sum = sum + bigramArray(word, nextWord)
        
        If sum >= lim Then
            Exit Function
        End If
        
        nextWord = nextWord + 1
    Wend
End Function

Private Function getRand(ByVal x As Long) As Long
    'getRand = CInt(CSng(x * Rnd()) / (RAND_MAX + 1#))
    getRand = CLng(Rnd() * x)
End Function

Public Sub emitMatrix()
    Dim x As Integer
    Dim y As Integer
    Dim outLine As String
    
    For x = 0 To loadedWords - 1
        outLine = Right(Space(20) & wordList(x), 20) & ":"
        For y = 0 To loadedWords - 1
            outLine = outLine & bigramArray(x, y)
        Next y
        outLine = outLine & ":" & sumVector(x)
        Debug.Print outLine
    Next x
End Sub

Public Sub emitWords()
    Dim i As Integer
    For i = LBound(wordList) To UBound(wordList)
        Debug.Print i & " -> " & wordList(i)
    Next i
End Sub
