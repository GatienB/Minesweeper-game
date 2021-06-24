Attribute VB_Name = "Demineur"

Global Board As GameBoard
Global OffsetRows As Integer
Global OffsetColumns As Integer

Public Sub ConfigureWorksheet()
Attribute ConfigureWorksheet.VB_ProcData.VB_Invoke_Func = "d\n14"
    Dim sht As Worksheet, found As Boolean
    Dim t As Range, btn As Button
    Const BUTTONS_ROW As Integer = 1
    
    For Each sht In Sheets
        If sht.Name = "Démineur" Then
            sht.Activate
            ActiveSheet.UsedRange.Delete
            found = True
            Exit For
        End If
    Next sht
    If Not found Then
        Set sht = Sheets.Add(after:=Sheets(1))
        sht.Name = "Démineur"
    End If
    ActiveSheet.Buttons.Delete
    ' button Facile
    Set t = ActiveSheet.Range(Cells(BUTTONS_ROW, 3), Cells(BUTTONS_ROW, 5))
    t.Select
    t.Clear
    Set btn = ActiveSheet.Buttons.Add(t.Left, t.Top, t.Width, t.Height)
    With btn
      .OnAction = "StartNewGame"
      .Caption = "Facile"
      .Name = "BtnFacile"
    End With
    ' button Intermédiaire
    Set t = ActiveSheet.Range(Cells(BUTTONS_ROW, 6), Cells(BUTTONS_ROW, 8))
    t.Select
    t.Clear
    Set btn = ActiveSheet.Buttons.Add(t.Left, t.Top, t.Width, t.Height)
    With btn
      .OnAction = "StartNewGame"
      .Caption = "Intermédiaire"
      .Name = "BtnIntermediaire"
    End With
    ' button Difficile
    Set t = ActiveSheet.Range(Cells(BUTTONS_ROW, 9), Cells(BUTTONS_ROW, 11))
    t.Select
    t.Clear
    Set btn = ActiveSheet.Buttons.Add(t.Left, t.Top, t.Width, t.Height)
    With btn
      .OnAction = "StartNewGame"
      .Caption = "Difficile"
      .Name = "BtnDifficile"
    End With
End Sub

Private Sub StartNewGame()
    Dim r As Integer, c As Integer, bombs As Integer
    If TypeName(Application.Caller) = "String" Then
        Select Case (Application.Caller)
            Case "BtnFacile"
                r = 10
                c = 10
                bombs = 10
                Debug.Print "facile"
            Case "BtnIntermediaire"
                r = 16
                c = 16
                bombs = 40
                Debug.Print "Intermediaire"
            Case "BtnDifficile"
                r = 16
                c = 30
                bombs = 100
                Debug.Print "Difficile"
        End Select
    Else
        Err.Raise 1, "", "StartNewGame - Unknown caller: " & TypeName(Application.Caller)
    End If
    
    InitGameVariables r, c, bombs
    ConfigureWorksheet
    
    Application.ScreenUpdating = False
    StyleBoard
    
    Board.GenerateBoardGame
    Board.DisplayBoard
    
    Application.ScreenUpdating = True
    
    Board.StartTime = Timer
End Sub

Private Sub InitGameVariables(nbrows As Integer, nbCols As Integer, nbBombs As Integer)
    OffsetRows = 2
    OffsetColumns = 2
    Set Board = New GameBoard
    Board.BoardRows = nbrows
    Board.BoardColumns = nbCols
    Board.BombsInitial = nbBombs
    Board.BombsRemaining = Board.BombsInitial
    Board.InitBoardGame
    Board.IsInGame = True
End Sub

Private Sub StyleBoard()
    SetStyleBoardGame
    ZoomFit
End Sub

Private Sub ZoomFit()
    Range(Cells(1, 1 + OffsetColumns - 1), _
            Cells(Board.BoardRows + OffsetRows + 1, Board.BoardColumns + OffsetColumns + 1)).Select
    ActiveWindow.Zoom = True
    Cells(1 + OffsetRows, 1 + OffsetColumns).Select
End Sub

Private Sub SetStyleBoardGame()
    ' Borders game board
    Dim r As Range
    Set r = Range(Cells(1 + OffsetRows, 1 + OffsetColumns), _
            Cells(Board.BoardRows + OffsetRows, Board.BoardColumns + OffsetColumns))
    
    r.VerticalAlignment = xlVAlignCenter
    r.HorizontalAlignment = xlVAlignCenter
    r.Borders(xlEdgeTop).Weight = xlThick
    r.Borders(xlEdgeRight).Weight = xlThick
    r.Borders(xlEdgeBottom).Weight = xlThick
    r.Borders(xlEdgeLeft).Weight = xlThick
End Sub
