VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private boardTablep() As gameCell
Private bombsInitialp As Integer
Private bombsRemainingp As Integer
Private boardRowsp As Integer
Private boardColumnsp As Integer
Private isInGamep As Boolean
Private startTimep As Single

Public Property Get BoardTable() As Variant
    BoardTable = boardTablep
End Property

Public Property Get BombsInitial() As Integer
    BombsInitial = bombsInitialp
End Property

Public Property Let BombsInitial(ByVal vNewValue As Integer)
    bombsInitialp = vNewValue
End Property

Public Property Get BombsRemaining() As Integer
    BombsRemaining = bombsRemainingp
End Property

Public Property Let BombsRemaining(ByVal vNewValue As Integer)
    bombsRemainingp = vNewValue
End Property

Public Property Get BoardRows() As Integer
    BoardRows = boardRowsp
End Property

Public Property Let BoardRows(ByVal vNewValue As Integer)
    boardRowsp = vNewValue
End Property

Public Property Get BoardColumns() As Integer
    BoardColumns = boardColumnsp
End Property

Public Property Let BoardColumns(ByVal vNewValue As Integer)
    boardColumnsp = vNewValue
End Property

Public Property Get IsInGame() As Boolean
    IsInGame = isInGamep
End Property

Public Property Let IsInGame(ByVal vNewValue As Boolean)
    isInGamep = vNewValue
End Property

Public Property Get StartTime() As Single
    StartTime = startTimep
End Property

Public Property Let StartTime(ByVal vNewValue As Single)
    startTimep = vNewValue
End Property

Public Sub InitBoardGame()
    Dim r As Integer
    Dim c As Integer
    Dim rowWithOffset As Integer
    Dim colWithOffset As Integer
    
    ReDim boardTablep(boardRowsp, boardColumnsp)
    For r = 1 To boardRowsp
        rowWithOffset = r + OffsetRows
        For c = 1 To boardColumnsp
            colWithOffset = c + OffsetColumns
            Set boardTablep(r, c) = New gameCell
            With boardTablep(r, c)
                .IsUncovered = False
                .MyRow = rowWithOffset
                .MyColumn = colWithOffset
                .MyCellType = EmptyCell
            End With
        Next c
    Next r
End Sub

Public Sub GenerateBoardGame()
    Dim countBombs As Integer
    Dim r As Integer
    Dim c As Integer
    Dim cel As gameCell
    countBombs = bombsInitialp
    Debug.Print "Generate bombs"
    Do While countBombs > 0
        r = GetRandomNumber(boardRowsp)
        c = GetRandomNumber(boardColumnsp)
        Set cel = GetBoardCell(r, c)
        If Not cel.IsBomb Then
            cel.IsBomb = True
            countBombs = countBombs - 1
        End If
        Debug.Print Space(5) & cel.ReadableProperties
    Loop
    
    ' set bombsAroundp
    Debug.Print "Count Bombs around cells"
    For r = 1 To boardRowsp
        For c = 1 To boardColumnsp
            Set cel = GetBoardCell(r, c)
            If cel.MyCellType = EmptyCell Then
                cel.BombsAround = CountBombsAround(cel)
            End If
        Next c
    Next r
    
End Sub

Public Function IsCoordinatesInGameBoard(r As Integer, c As Integer) As Boolean
    IsCoordinatesInGameBoard = r > 0 And r <= boardRowsp _
        And c > 0 And c <= boardColumnsp
End Function

Public Sub CellUncovered(r As Integer, c As Integer)
    Dim gameCell As gameCell
    Set gameCell = GetBoardCell(r, c)
    If Not gameCell.IsUncovered And gameCell.MyCellType <> FlagCell Then
        If Not gameCell.IsBomb And (gameCell.MyCellType = EmptyCell Or gameCell.MyCellType = MaybeCell) Then
            gameCell.IsUncovered = True
            If gameCell.BombsAround > 0 Then
                ' uncover 1 cell
                SetCellBackgroundAndText gameCell, Cells(gameCell.MyRow, gameCell.MyColumn)
            Else
                ' uncover all cells with 0 bombsaround
                SetCellBackgroundAndText gameCell, Cells(gameCell.MyRow, gameCell.MyColumn)
                UncoverBlankCells gameCell
                Debug.Print "empty"
            End If
        ElseIf gameCell.IsBomb Then
            DisplayGameLost gameCell
            isInGamep = False
            MsgBox "Vous avez perdu" & _
                vbNewLine & vbNewLine & "Durée de la partie : " & ConvertTimerToReadableTime(Timer - startTimep)
        End If
    End If
End Sub

Public Sub SetFlag(r As Integer, c As Integer)
    Dim gameCell As gameCell
    Set gameCell = GetBoardCell(r, c)
    If Not gameCell.IsUncovered Then
        If gameCell.MyCellType = FlagCell Then
            Board.BombsRemaining = Board.BombsRemaining + 1
            gameCell.MyCellType = MaybeCell
        ElseIf gameCell.MyCellType = EmptyCell Then
            Board.BombsRemaining = Board.BombsRemaining - 1
            gameCell.MyCellType = FlagCell
        ElseIf gameCell.MyCellType = BombCell Then
            Board.BombsRemaining = Board.BombsRemaining - 1
            gameCell.MyCellType = FlagCell
        ElseIf gameCell.MyCellType = MaybeCell Then
            gameCell.MyCellType = EmptyCell
        End If
        SetCellBackgroundAndText gameCell, Cells(gameCell.MyRow, gameCell.MyColumn)
        DisplayRemainingBombs
    End If
End Sub

Public Sub CheckEndGame()
    Dim r As Integer, c As Integer
    Dim totalFlag As Integer, goodFlags As Integer
    Dim totalUncovered As Integer
    Dim gCell As gameCell
    For r = 1 To boardRowsp
        For c = 1 To boardColumnsp
            Set gCell = boardTablep(r, c)
            If gCell.MyCellType = FlagCell Then
                totalFlag = totalFlag + 1
                If gCell.IsBomb Then
                    goodFlags = goodFlags + 1
                End If
            End If
            If gCell.IsUncovered Then
                totalUncovered = totalUncovered + 1
            End If
        Next c
    Next r
    If totalUncovered = (boardRowsp * boardColumnsp - bombsInitialp) Or _
       (totalFlag = bombsInitialp And goodFlags = bombsInitialp) Then
        MsgBox "Bravo ! Vous avez gagné." & _
                vbNewLine & vbNewLine & "Durée de la partie : " & ConvertTimerToReadableTime(Timer - startTimep)
        isInGamep = False
    End If
End Sub

Public Sub DisplayBoard()
    ' ratio de 6 entre column width et rowHeight
    Dim r As Integer
    Dim c As Integer
    Dim cell As Range
    Dim gameCell As gameCell
    For r = 1 To boardRowsp
        For c = 1 To boardColumnsp
            Set gameCell = GetBoardCell(r, c)
            If c = 1 Then
                ActiveSheet.rows(gameCell.MyRow).RowHeight = 24
            End If
            If r = 1 Then
                ActiveSheet.columns(gameCell.MyColumn).ColumnWidth = 4
            End If
            Set cell = Cells(gameCell.MyRow, gameCell.MyColumn)
            cell.Borders.Color = 0
            SetCellBackgroundAndText gameCell, cell
        Next c
    Next r
    DisplayRemainingBombs True
End Sub

Public Sub DisplayGameLost(cellPlayed As gameCell)
    Dim r As Integer, c As Integer
    Dim cell As Range, gCell As gameCell
    Set cell = Cells(cellPlayed.MyRow, cellPlayed.MyColumn)
    cell.Interior.Color = RGB(255, 0, 0)
    For r = 1 To boardRowsp
        For c = 1 To boardColumnsp
            Set gCell = GetBoardCell(r, c)
            If gCell.IsBomb And gCell.MyCellType <> FlagCell Then
                Set cell = Cells(gCell.MyRow, gCell.MyColumn)
                cell.Value = ChrW(&HD83D) & ChrW(&HDCA3)
                '%uD83D%uDCA3 = &HD83D&HDCA3
            End If
            If Not gCell.IsBomb And gCell.MyCellType = FlagCell Then
                Set cell = Cells(gCell.MyRow, gCell.MyColumn)
                cell.Borders(xlDiagonalDown).Color = RGB(255, 0, 0)
                cell.Borders(xlDiagonalUp).Color = RGB(255, 0, 0)
                cell.Borders(xlDiagonalDown).LineStyle = xlContinuous
                cell.Borders(xlDiagonalDown).Weight = xlThick
            End If
        Next c
    Next r
End Sub

' Private Functions
Private Function CountBombsAround(cel As gameCell) As Integer
    Dim counter As Integer
    Dim r As Integer
    Dim c As Integer
    counter = 0
    For r = cel.MyRow - 1 To cel.MyRow + 1
        For c = cel.MyColumn - 1 To cel.MyColumn + 1
            If IsCoordinatesInGameBoard(r - OffsetRows, c - OffsetColumns) Then
                If GetBoardCell(r - OffsetRows, c - OffsetColumns).IsBomb Then
                    counter = counter + 1
                End If
            End If
        Next c
    Next r
    CountBombsAround = counter
End Function

Private Sub SetCellBackgroundAndText(gCell As gameCell, cell As Range)
    If gCell.IsUncovered Then
        If gCell.BombsAround > 0 Then
            cell.Value = gCell.BombsAround
        End If
        If gCell.MyCellType = EmptyCell Or gCell.MyCellType = MaybeCell Then
            cell.Interior.Color = RGB(255, 255, 255)
            If gCell.BombsAround = 0 Then
                cell.Value = ""
            Else
                cell.Value = gCell.BombsAround
            End If
        Else
            cell.Interior.Color = RGB(255, 0, 0)
        End If
    Else
        If gCell.MyCellType = FlagCell Then
            cell.Value = ChrW(&H26F3)
        ElseIf gCell.MyCellType = EmptyCell Then
            cell.Value = ""
        ElseIf gCell.MyCellType = MaybeCell Then
            cell.Value = "?"
        End If
        cell.Interior.Color = RGB(200, 200, 200)
    End If
End Sub

Private Sub UncoverBlankCells(gCell As gameCell)
    Dim r As Integer, c As Integer, cell As Range
    Dim i As Integer, cellToCheck As gameCell
    Dim ended As Boolean
    For r = gCell.MyRow - 1 To gCell.MyRow + 1
        For c = gCell.MyColumn - 1 To gCell.MyColumn + 1
            If IsCoordinatesInGameBoard(r - OffsetRows, c - OffsetColumns) Then
                Set cellToCheck = GetBoardCell(r - OffsetRows, c - OffsetColumns)
                'Debug.Print "(" & r & "," & c & ") " & GetBoardCell(r, c).DisplayGameCellTypeToString
                If cellToCheck.IsUncovered = False And _
                    (cellToCheck.MyCellType = EmptyCell Or cellToCheck.MyCellType = MaybeCell) And _
                    Not cellToCheck.IsBomb Then
                    cellToCheck.IsUncovered = True
                    SetCellBackgroundAndText cellToCheck, Cells(cellToCheck.MyRow, cellToCheck.MyColumn)
                    If cellToCheck.BombsAround = 0 Then
                        UncoverBlankCells cellToCheck
                    End If
                End If
            End If
        Next c
    Next r
End Sub

Private Function GetRandomNumber(max As Integer, Optional min As Integer = 1)
    GetRandomNumber = Int((max * Rnd) + min)
End Function

Private Function GetBoardCell(r As Integer, c As Integer) As gameCell
    Set GetBoardCell = boardTablep(r, c)
End Function

Private Function DisplayRemainingBombs(Optional isFirstcall As Boolean = False)
    Dim t As Range
    Set t = ActiveSheet.Range(Cells(1 + OffsetRows - 1, 1 + OffsetColumns), Cells(1 + OffsetRows - 1, 1 + OffsetColumns + 3))
    If isFirstcall Then
        t.Select
        t.Merge
        Selection.HorizontalAlignment = xlLeft
        Selection.VerticalAlignment = xlCenter
        t.Font.Size = 16
        t.Font.Bold = True
    End If
    t.Value = ChrW(&HD83D) & ChrW(&HDCA3) & Space(3) & bombsRemainingp
End Function

Private Function ConvertTimerToReadableTime(timestamp As Single) As String
    Dim minutes As Integer, seconds As Integer
    seconds = timestamp Mod 60
    minutes = Int(timestamp / 60)
    ConvertTimerToReadableTime = minutes & "m" & seconds & "s."
End Function


