VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private rowP As Integer
Private columnP As Integer
Private cellTypeP As GameCellTypesEnum
Private isUncoveredP As Boolean
Private bombsAroundp As Integer
Private isBombp As Boolean

Public Property Let MyRow(iRow As Integer)
    rowP = iRow
End Property

Public Property Get MyRow() As Integer
    MyRow = rowP
End Property

Public Property Let MyColumn(iCol As Integer)
    columnP = iCol
End Property

Public Property Get MyColumn() As Integer
    MyColumn = columnP
End Property

Public Property Let MyCellType(iCellT As GameCellTypesEnum)
    cellTypeP = iCellT
End Property

Public Property Get MyCellType() As GameCellTypesEnum
    MyCellType = cellTypeP
End Property

Public Property Let IsUncovered(isUnco As Boolean)
    isUncoveredP = isUnco
End Property

Public Property Get IsUncovered() As Boolean
    IsUncovered = isUncoveredP
End Property

Public Property Let BombsAround(bombs As Integer)
    bombsAroundp = bombs
End Property

Public Property Get BombsAround() As Integer
    BombsAround = bombsAroundp
End Property

Public Property Let IsBomb(vIsBomb As Boolean)
    isBombp = vIsBomb
End Property

Public Property Get IsBomb() As Boolean
    IsBomb = isBombp
End Property


Public Function GetCoordinates() As String
    GetCoordinates = "(" & rowP & "," & columnP & ")"
End Function

Public Function ReadableProperties() As String
    ReadableProperties = "GameCell(coordinates=" & GetCoordinates & ", " & _
                        "cellType=" & GameCellTypeToString(cellTypeP) & ", " & _
                        "isUncovered=" & isUncoveredP & ", " & _
                        "bombsAround=" & bombsAroundp & ", " & _
                        "isBomb=" & isBombp & ")"
End Function


Public Function DisplayGameCellTypeToString() As String
    Select Case cellTypeP
        Case GameCellTypesEnum.EmptyCell
            DisplayGameCellTypeToString = "EmptyCell"
        Case GameCellTypesEnum.BombCell
            DisplayGameCellTypeToString = "BombCell"
        Case GameCellTypesEnum.FlagCell
            DisplayGameCellTypeToString = "FlagCell"
        Case GameCellTypesEnum.MaybeCell
            DisplayGameCellTypeToString = "MaybeCell"
        Case Else
            Err.Raise 1, , "Unknown GameCellType : " & cellTypeP
    End Select
End Function

Public Function GameCellTypeToString(val As GameCellTypesEnum) As String
    Select Case val
        Case GameCellTypesEnum.EmptyCell
            GameCellTypeToString = "EmptyCell"
        Case GameCellTypesEnum.BombCell
            GameCellTypeToString = "BombCell"
        Case GameCellTypesEnum.FlagCell
            GameCellTypeToString = "FlagCell"
        Case GameCellTypesEnum.MaybeCell
            GameCellTypeToString = "MaybeCell"
        Case Else
            Err.Raise 1, , "Unknown GameCellType : " & val
    End Select
End Function

