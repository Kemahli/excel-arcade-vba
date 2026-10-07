Attribute VB_Name = "Module1"
Option Explicit

' ============================================================
' EXCEL CHESS - FINAL ENGINE V1.0
' Sheet: Chess Board
' Board: B2:I9
' ============================================================

Private Const BOARD_TOP As Long = 2
Private Const BOARD_BOTTOM As Long = 9
Private Const BOARD_LEFT As Long = 2
Private Const BOARD_RIGHT As Long = 9

Private Const PIECE_FONT_SIZE As Long = 34
Private Const BOARD_ROW_HEIGHT As Double = 52


' ============================================================
' UNICODE CHESS PIECES
' ============================================================

Private Const WHITE_KING As Long = 9812
Private Const WHITE_QUEEN As Long = 9813
Private Const WHITE_ROOK As Long = 9814
Private Const WHITE_BISHOP As Long = 9815
Private Const WHITE_KNIGHT As Long = 9816
Private Const WHITE_PAWN As Long = 9817

Private Const BLACK_KING As Long = 9818
Private Const BLACK_QUEEN As Long = 9819
Private Const BLACK_ROOK As Long = 9820
Private Const BLACK_BISHOP As Long = 9821
Private Const BLACK_KNIGHT As Long = 9822
Private Const BLACK_PAWN As Long = 9823


' ============================================================
' GAME STATE
' ============================================================

Public SelectedRow As Long
Public SelectedCol As Long

Public CurrentTurn As String
Public GameStatus As String
Public GameOver As Boolean

Public HalfmoveClock As Long

Public EnPassantRow As Long
Public EnPassantCol As Long

Public WhiteKingMoved As Boolean
Public BlackKingMoved As Boolean

Public WhiteRookAMoved As Boolean
Public WhiteRookHMoved As Boolean
Public BlackRookAMoved As Boolean
Public BlackRookHMoved As Boolean

Public LastFromRow As Long
Public LastFromCol As Long
Public LastToRow As Long
Public LastToCol As Long

Public MoveCount As Long
Public MoveList() As String

Public UndoCount As Long
Public UndoStack() As String

Public PositionHistoryCount As Long
Public PositionHistory() As String

Public PositionCounts As Object


' ============================================================
' START / NEW GAME
' ============================================================

Public Sub StartGame()

    Dim ws As Worksheet
    Dim c As Long

    On Error GoTo StartGameError

    Set ws = ThisWorkbook.Worksheets("Chess Board")

    Application.EnableEvents = False
    Application.ScreenUpdating = False


    ws.Range("B2:I9").ClearContents
    ws.Range("K15:M200").ClearContents


    ' ========================================================
    ' BLACK
    ' ========================================================

    ws.Cells(2, 2).Value = ChrW(BLACK_ROOK)
    ws.Cells(2, 3).Value = ChrW(BLACK_KNIGHT)
    ws.Cells(2, 4).Value = ChrW(BLACK_BISHOP)
    ws.Cells(2, 5).Value = ChrW(BLACK_QUEEN)
    ws.Cells(2, 6).Value = ChrW(BLACK_KING)
    ws.Cells(2, 7).Value = ChrW(BLACK_BISHOP)
    ws.Cells(2, 8).Value = ChrW(BLACK_KNIGHT)
    ws.Cells(2, 9).Value = ChrW(BLACK_ROOK)

    For c = 2 To 9
        ws.Cells(3, c).Value = ChrW(BLACK_PAWN)
    Next c


    ' ========================================================
    ' WHITE
    ' ========================================================

    For c = 2 To 9
        ws.Cells(8, c).Value = ChrW(WHITE_PAWN)
    Next c

    ws.Cells(9, 2).Value = ChrW(WHITE_ROOK)
    ws.Cells(9, 3).Value = ChrW(WHITE_KNIGHT)
    ws.Cells(9, 4).Value = ChrW(WHITE_BISHOP)
    ws.Cells(9, 5).Value = ChrW(WHITE_QUEEN)
    ws.Cells(9, 6).Value = ChrW(WHITE_KING)
    ws.Cells(9, 7).Value = ChrW(WHITE_BISHOP)
    ws.Cells(9, 8).Value = ChrW(WHITE_KNIGHT)
    ws.Cells(9, 9).Value = ChrW(WHITE_ROOK)


    ' ========================================================
    ' RESET STATE
    ' ========================================================

    CurrentTurn = "WHITE"

    GameStatus = "Playing"
    GameOver = False

    SelectedRow = 0
    SelectedCol = 0

    HalfmoveClock = 0

    EnPassantRow = 0
    EnPassantCol = 0

    WhiteKingMoved = False
    BlackKingMoved = False

    WhiteRookAMoved = False
    WhiteRookHMoved = False
    BlackRookAMoved = False
    BlackRookHMoved = False

    LastFromRow = 0
    LastFromCol = 0
    LastToRow = 0
    LastToCol = 0


    MoveCount = 0
    Erase MoveList


    UndoCount = 0
    Erase UndoStack


    PositionHistoryCount = 0
    Erase PositionHistory

    Set PositionCounts = CreateObject("Scripting.Dictionary")


    ApplyBoardFormatting

    ResetBoardColors

    AddCurrentPosition

    RefreshHistory
    UpdatePanel


    ws.Activate
    ws.Range("A1").Select


StartGameExit:

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Exit Sub


StartGameError:

    MsgBox _
        "StartGame error: " & Err.Description, _
        vbExclamation

    Resume StartGameExit

End Sub


' ============================================================
' BOARD FORMATTING
' ============================================================

Private Sub ApplyBoardFormatting()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Chess Board")

    With ws.Range("B2:I9")

        .Font.Name = "Segoe UI Symbol"
        .Font.Size = PIECE_FONT_SIZE

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    ws.Rows("2:9").RowHeight = BOARD_ROW_HEIGHT

End Sub


' ============================================================
' MASTER CLICK HANDLER
' ============================================================

Public Sub HandleSheetClick(ByVal Target As Range)

    Dim ws As Worksheet
    Dim clickedPiece As Long

    Set ws = ThisWorkbook.Worksheets("Chess Board")


    ' ========================================================
    ' NEW GAME CELL BUTTON
    ' ========================================================

    If Not Intersect(Target, ws.Range("K8:M9")) Is Nothing Then

        StartGame
        Exit Sub

    End If


    ' ========================================================
    ' UNDO CELL BUTTON
    ' ========================================================

    If Not Intersect(Target, ws.Range("K10:M11")) Is Nothing Then

        UndoMove
        Exit Sub

    End If


    ' ========================================================
    ' STATUS = DRAW CLAIM
    ' ========================================================

    If Not Intersect(Target, ws.Range("L5:M5")) Is Nothing Then

        If Not GameOver Then

            If CanClaimDrawNow() Then
                ClaimDraw
            End If

        End If

        Exit Sub

    End If


    ' ========================================================
    ' LOST VBA STATE
    ' ========================================================

    If CurrentTurn = "" Then

        StartGame
        Exit Sub

    End If


    If GameOver Then Exit Sub

    If Target.CountLarge > 1 Then Exit Sub


    If Intersect(Target, ws.Range("B2:I9")) Is Nothing Then
        Exit Sub
    End If


    clickedPiece = _
        PieceCodeAt(ws, Target.Row, Target.Column)


    ' ========================================================
    ' NOTHING SELECTED
    ' ========================================================

    If SelectedRow = 0 Then

        If clickedPiece = 0 Then Exit Sub

        If BelongsToSide(clickedPiece, CurrentTurn) Then

            SelectPiece _
                Target.Row, _
                Target.Column

        End If

        Exit Sub

    End If


    ' ========================================================
    ' TRY TO MOVE
    ' ========================================================

    If IsLegalMove( _
        SelectedRow, _
        SelectedCol, _
        Target.Row, _
        Target.Column _
    ) Then

        MovePiece _
            SelectedRow, _
            SelectedCol, _
            Target.Row, _
            Target.Column

        Exit Sub

    End If


    ' ========================================================
    ' SELECT DIFFERENT OWN PIECE
    ' ========================================================

    If clickedPiece <> 0 Then

        If BelongsToSide(clickedPiece, CurrentTurn) Then

            SelectPiece _
                Target.Row, _
                Target.Column

            Exit Sub

        End If

    End If


    ClearSelection

End Sub


' ============================================================
' DRAW CLAIM
' ============================================================

Public Sub ClaimDraw()

    Dim repetitionCount As Long

    If GameOver Then Exit Sub


    repetitionCount = _
        CurrentPositionOccurrenceCount()


    If repetitionCount >= 3 And _
       HalfmoveClock >= 100 Then

        GameStatus = _
            "Draw - Claimed"

        GameOver = True


    ElseIf repetitionCount >= 3 Then

        GameStatus = _
            "Draw - Threefold repetition"

        GameOver = True


    ElseIf HalfmoveClock >= 100 Then

        GameStatus = _
            "Draw - 50 move rule"

        GameOver = True


    Else

        MsgBox _
            "A draw cannot currently be claimed.", _
            vbInformation

        Exit Sub

    End If


    SelectedRow = 0
    SelectedCol = 0

    ResetBoardColors
    HighlightLastMove
    HighlightCheckedKing

    UpdatePanel

End Sub


Private Function CanClaimDrawNow() As Boolean

    CanClaimDrawNow = _
        CurrentPositionOccurrenceCount() >= 3 Or _
        HalfmoveClock >= 100

End Function


Private Function GetDrawClaimText() As String

    Dim repetitionCount As Long

    repetitionCount = _
        CurrentPositionOccurrenceCount()


    If repetitionCount >= 3 And _
       HalfmoveClock >= 100 Then

        GetDrawClaimText = _
            "Claim draw (3-fold / 50-move)"


    ElseIf repetitionCount >= 3 Then

        GetDrawClaimText = _
            "Claim draw (3-fold)"


    ElseIf HalfmoveClock >= 100 Then

        GetDrawClaimText = _
            "Claim draw (50-move)"


    Else

        GetDrawClaimText = ""

    End If

End Function


' ============================================================
' SELECT PIECE
' ============================================================

Private Sub SelectPiece( _
    ByVal r As Long, _
    ByVal c As Long _
)

    Dim ws As Worksheet

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    SelectedRow = r
    SelectedCol = c


    ResetBoardColors

    HighlightLastMove
    HighlightCheckedKing


    ' Selected piece
    ws.Cells(r, c).Interior.Color = _
        RGB(255, 235, 174)


    ShowLegalMoves r, c

    UpdatePanel

End Sub


Private Sub ClearSelection()

    SelectedRow = 0
    SelectedCol = 0


    ResetBoardColors

    HighlightLastMove
    HighlightCheckedKing

    UpdatePanel

End Sub


' ============================================================
' SHOW LEGAL MOVES
' ============================================================

Private Sub ShowLegalMoves( _
    ByVal fromRow As Long, _
    ByVal fromCol As Long _
)

    Dim ws As Worksheet

    Dim r As Long
    Dim c As Long

    Dim targetPiece As Long

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    For r = BOARD_TOP To BOARD_BOTTOM

        For c = BOARD_LEFT To BOARD_RIGHT


            If IsLegalMove( _
                fromRow, _
                fromCol, _
                r, _
                c _
            ) Then


                targetPiece = _
                    PieceCodeAt(ws, r, c)


                If targetPiece = 0 Then

                    ' Empty legal destination
                    ws.Cells(r, c).Interior.Color = _
                        RGB(211, 233, 218)

                Else

                    ' Capture destination
                    ws.Cells(r, c).Interior.Color = _
                        RGB(245, 207, 207)

                End If

            End If

        Next c

    Next r

End Sub


' ============================================================
' MOVE PIECE
' ============================================================

Private Sub MovePiece( _
    ByVal fromRow As Long, _
    ByVal fromCol As Long, _
    ByVal toRow As Long, _
    ByVal toCol As Long _
)

    Dim ws As Worksheet

    Dim movingPiece As Long
    Dim capturedPiece As Long

    Dim moverWasWhite As Boolean

    Dim isCastle As Boolean
    Dim isEnPassant As Boolean
    Dim wasCapture As Boolean

    Dim promotionChoice As String

    Dim sanText As String

    Dim opponentInCheck As Boolean
    Dim opponentHasMove As Boolean

    Dim repetitionCount As Long

    On Error GoTo MoveError

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")

    Application.ScreenUpdating = False


    movingPiece = _
        PieceCodeAt(ws, fromRow, fromCol)

    capturedPiece = _
        PieceCodeAt(ws, toRow, toCol)

    moverWasWhite = _
        IsWhitePieceCode(movingPiece)


    ' ========================================================
    ' SPECIAL MOVE DETECTION
    ' ========================================================

    isCastle = False
    isEnPassant = False


    If movingPiece = WHITE_KING Or _
       movingPiece = BLACK_KING Then

        If Abs(toCol - fromCol) = 2 Then
            isCastle = True
        End If

    End If


    If movingPiece = WHITE_PAWN Or _
       movingPiece = BLACK_PAWN Then

        If toRow = EnPassantRow And _
           toCol = EnPassantCol And _
           capturedPiece = 0 And _
           Abs(toCol - fromCol) = 1 Then

            isEnPassant = True

            capturedPiece = _
                PieceCodeAt(ws, fromRow, toCol)

        End If

    End If


    wasCapture = _
        capturedPiece <> 0 Or _
        isEnPassant


    ' ========================================================
    ' PROMOTION CHOICE
    ' ========================================================

    promotionChoice = ""


    If movingPiece = WHITE_PAWN And _
       toRow = BOARD_TOP Then

        promotionChoice = _
            GetPromotionChoice()


    ElseIf movingPiece = BLACK_PAWN And _
           toRow = BOARD_BOTTOM Then

        promotionChoice = _
            GetPromotionChoice()

    End If


    ' ========================================================
    ' BUILD SAN WHILE ORIGINAL POSITION STILL EXISTS
    ' ========================================================

    sanText = _
        BuildSANBase( _
            movingPiece, _
            fromRow, _
            fromCol, _
            toRow, _
            toCol, _
            wasCapture, _
            isCastle, _
            promotionChoice _
        )


    ' ========================================================
    ' SAVE STATE FOR UNDO
    ' ========================================================

    PushUndoSnapshot


    ' ========================================================
    ' CASTLING RIGHTS
    ' ========================================================

    UpdateCastlingRights _
        movingPiece, _
        capturedPiece, _
        fromRow, _
        fromCol, _
        toRow, _
        toCol


    ' ========================================================
    ' EN PASSANT CAPTURE
    ' ========================================================

    If isEnPassant Then

        ws.Cells(fromRow, toCol).ClearContents

    End If


    ' ========================================================
    ' NORMAL PIECE MOVE
    ' ========================================================

    ws.Cells(toRow, toCol).Value = _
        ws.Cells(fromRow, fromCol).Value

    ws.Cells(fromRow, fromCol).ClearContents


    ' ========================================================
    ' CASTLING ROOK MOVE
    ' ========================================================

    If isCastle Then

        If toCol = 8 Then

            ' h-file rook -> f-file

            ws.Cells(toRow, 7).Value = _
                ws.Cells(toRow, 9).Value

            ws.Cells(toRow, 9).ClearContents


        ElseIf toCol = 4 Then

            ' a-file rook -> d-file

            ws.Cells(toRow, 5).Value = _
                ws.Cells(toRow, 2).Value

            ws.Cells(toRow, 2).ClearContents

        End If

    End If


    ' ========================================================
    ' PROMOTION
    ' ========================================================

    If promotionChoice <> "" Then

        ws.Cells(toRow, toCol).Value = _
            ChrW( _
                PromotionPieceCode( _
                    moverWasWhite, _
                    promotionChoice _
                ) _
            )

    End If


    ' ========================================================
    ' 50 / 75 MOVE CLOCK
    ' ========================================================

    If movingPiece = WHITE_PAWN Or _
       movingPiece = BLACK_PAWN Or _
       wasCapture Then

        HalfmoveClock = 0

    Else

        HalfmoveClock = _
            HalfmoveClock + 1

    End If


    ' ========================================================
    ' NEW EN PASSANT TARGET
    ' ========================================================

    EnPassantRow = 0
    EnPassantCol = 0


    If movingPiece = WHITE_PAWN Then

        If fromRow - toRow = 2 Then

            EnPassantRow = _
                fromRow - 1

            EnPassantCol = _
                fromCol

        End If


    ElseIf movingPiece = BLACK_PAWN Then

        If toRow - fromRow = 2 Then

            EnPassantRow = _
                fromRow + 1

            EnPassantCol = _
                fromCol

        End If

    End If


    ' ========================================================
    ' LAST MOVE
    ' ========================================================

    LastFromRow = fromRow
    LastFromCol = fromCol

    LastToRow = toRow
    LastToCol = toCol


    ' ========================================================
    ' CHANGE TURN
    ' ========================================================

    If moverWasWhite Then

        CurrentTurn = "BLACK"

    Else

        CurrentTurn = "WHITE"

    End If


    SelectedRow = 0
    SelectedCol = 0


    ' ========================================================
    ' RECORD POSITION FOR REPETITION
    ' ========================================================

    repetitionCount = _
        AddCurrentPosition()


    ' ========================================================
    ' CHECK / MATE / STALEMATE
    ' ========================================================

    opponentInCheck = _
        IsKingInCheck(CurrentTurn)

    opponentHasMove = _
        HasAnyLegalMove(CurrentTurn)


    GameOver = False


    If Not opponentHasMove Then

        If opponentInCheck Then

            If moverWasWhite Then

                GameStatus = _
                    "Checkmate - White wins"

            Else

                GameStatus = _
                    "Checkmate - Black wins"

            End If

            GameOver = True


        Else

            GameStatus = _
                "Stalemate"

            GameOver = True

        End If


    Else

        If opponentInCheck Then

            GameStatus = "Check"

        Else

            GameStatus = "Playing"

        End If

    End If


    ' ========================================================
    ' AUTOMATIC 75-MOVE RULE
    '
    ' 150 halfmoves = 75 moves by each side combined
    ' Checkmate already takes priority above.
    ' ========================================================

    If Not GameOver Then

        If HalfmoveClock >= 150 Then

            GameStatus = _
                "Draw - 75 move rule"

            GameOver = True

        End If

    End If


    ' ========================================================
    ' AUTOMATIC FIVEFOLD REPETITION
    ' ========================================================

    If Not GameOver Then

        If repetitionCount >= 5 Then

            GameStatus = _
                "Draw - Fivefold repetition"

            GameOver = True

        End If

    End If


    ' ========================================================
    ' DEAD / INSUFFICIENT MATERIAL
    ' ========================================================

    If Not GameOver Then

        If IsInsufficientMaterial() Then

            GameStatus = _
                "Draw - Insufficient material"

            GameOver = True

        End If

    End If


    ' ========================================================
    ' SAN SUFFIX
    ' ========================================================

    If InStr( _
        1, _
        GameStatus, _
        "Checkmate", _
        vbTextCompare _
    ) > 0 Then

        sanText = _
            sanText & "#"


    ElseIf opponentInCheck Then

        sanText = _
            sanText & "+"

    End If


    ' ========================================================
    ' SAVE MOVE
    ' ========================================================

    RecordMove sanText


    ' ========================================================
    ' VISUAL REFRESH
    ' ========================================================

    ResetBoardColors

    HighlightLastMove
    HighlightCheckedKing

    RefreshHistory
    UpdatePanel


MoveExit:

    Application.ScreenUpdating = True
    Exit Sub


MoveError:

    MsgBox _
        "Move error: " & Err.Description, _
        vbExclamation

    Resume MoveExit

End Sub


' ============================================================
' LEGAL MOVE
' ============================================================

Private Function IsLegalMove( _
    ByVal fromRow As Long, _
    ByVal fromCol As Long, _
    ByVal toRow As Long, _
    ByVal toCol As Long _
) As Boolean

    Dim ws As Worksheet

    Dim movingPiece As Long
    Dim targetPiece As Long

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    IsLegalMove = False


    If Not IsInsideBoard(toRow, toCol) Then
        Exit Function
    End If


    If fromRow = toRow And _
       fromCol = toCol Then

        Exit Function

    End If


    movingPiece = _
        PieceCodeAt(ws, fromRow, fromCol)

    targetPiece = _
        PieceCodeAt(ws, toRow, toCol)


    If movingPiece = 0 Then
        Exit Function
    End If


    ' ========================================================
    ' KINGS ARE NEVER CAPTURED
    ' ========================================================

    If targetPiece = WHITE_KING Or _
       targetPiece = BLACK_KING Then

        Exit Function

    End If


    ' ========================================================
    ' CANNOT CAPTURE OWN PIECE
    ' ========================================================

    If targetPiece <> 0 Then


        If IsWhitePieceCode(movingPiece) And _
           IsWhitePieceCode(targetPiece) Then

            Exit Function

        End If


        If IsBlackPieceCode(movingPiece) And _
           IsBlackPieceCode(targetPiece) Then

            Exit Function

        End If

    End If


    ' ========================================================
    ' PIECE MOVEMENT RULE
    ' ========================================================

    If Not IsPseudoLegalMove( _
        fromRow, _
        fromCol, _
        toRow, _
        toCol _
    ) Then

        Exit Function

    End If


    ' ========================================================
    ' OWN KING MUST REMAIN SAFE
    ' ========================================================

    If Not SimulatedMoveLeavesKingSafe( _
        fromRow, _
        fromCol, _
        toRow, _
        toCol _
    ) Then

        Exit Function

    End If


    IsLegalMove = True

End Function


' ============================================================
' PIECE MOVEMENT
' ============================================================

Private Function IsPseudoLegalMove( _
    ByVal fromRow As Long, _
    ByVal fromCol As Long, _
    ByVal toRow As Long, _
    ByVal toCol As Long _
) As Boolean

    Dim ws As Worksheet

    Dim pieceCode As Long
    Dim targetPiece As Long

    Dim dr As Long
    Dim dc As Long

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    IsPseudoLegalMove = False


    pieceCode = _
        PieceCodeAt(ws, fromRow, fromCol)

    targetPiece = _
        PieceCodeAt(ws, toRow, toCol)


    dr = toRow - fromRow
    dc = toCol - fromCol


    Select Case pieceCode


        ' ====================================================
        ' WHITE PAWN
        ' ====================================================

        Case WHITE_PAWN


            If dc = 0 And _
               dr = -1 And _
               targetPiece = 0 Then

                IsPseudoLegalMove = True
                Exit Function

            End If


            If dc = 0 And _
               dr = -2 And _
               fromRow = 8 And _
               targetPiece = 0 Then


                If PieceCodeAt( _
                    ws, _
                    fromRow - 1, _
                    fromCol _
                ) = 0 Then

                    IsPseudoLegalMove = True
                    Exit Function

                End If

            End If


            If Abs(dc) = 1 And _
               dr = -1 Then


                If IsBlackPieceCode(targetPiece) Then

                    IsPseudoLegalMove = True
                    Exit Function

                End If


                If targetPiece = 0 And _
                   toRow = EnPassantRow And _
                   toCol = EnPassantCol Then


                    If PieceCodeAt( _
                        ws, _
                        fromRow, _
                        toCol _
                    ) = BLACK_PAWN Then

                        IsPseudoLegalMove = True
                        Exit Function

                    End If

                End If

            End If


        ' ====================================================
        ' BLACK PAWN
        ' ====================================================

        Case BLACK_PAWN


            If dc = 0 And _
               dr = 1 And _
               targetPiece = 0 Then

                IsPseudoLegalMove = True
                Exit Function

            End If


            If dc = 0 And _
               dr = 2 And _
               fromRow = 3 And _
               targetPiece = 0 Then


                If PieceCodeAt( _
                    ws, _
                    fromRow + 1, _
                    fromCol _
                ) = 0 Then

                    IsPseudoLegalMove = True
                    Exit Function

                End If

            End If


            If Abs(dc) = 1 And _
               dr = 1 Then


                If IsWhitePieceCode(targetPiece) Then

                    IsPseudoLegalMove = True
                    Exit Function

                End If


                If targetPiece = 0 And _
                   toRow = EnPassantRow And _
                   toCol = EnPassantCol Then


                    If PieceCodeAt( _
                        ws, _
                        fromRow, _
                        toCol _
                    ) = WHITE_PAWN Then

                        IsPseudoLegalMove = True
                        Exit Function

                    End If

                End If

            End If


        ' ====================================================
        ' KNIGHT
        ' ====================================================

        Case WHITE_KNIGHT, BLACK_KNIGHT


            If _
                (Abs(dr) = 2 And Abs(dc) = 1) Or _
                (Abs(dr) = 1 And Abs(dc) = 2) _
            Then

                IsPseudoLegalMove = True
                Exit Function

            End If


        ' ====================================================
        ' BISHOP
        ' ====================================================

        Case WHITE_BISHOP, BLACK_BISHOP


            If Abs(dr) = Abs(dc) Then

                If PathIsClear( _
                    fromRow, _
                    fromCol, _
                    toRow, _
                    toCol _
                ) Then

                    IsPseudoLegalMove = True
                    Exit Function

                End If

            End If


        ' ====================================================
        ' ROOK
        ' ====================================================

        Case WHITE_ROOK, BLACK_ROOK


            If dr = 0 Or dc = 0 Then

                If PathIsClear( _
                    fromRow, _
                    fromCol, _
                    toRow, _
                    toCol _
                ) Then

                    IsPseudoLegalMove = True
                    Exit Function

                End If

            End If


        ' ====================================================
        ' QUEEN
        ' ====================================================

        Case WHITE_QUEEN, BLACK_QUEEN


            If _
                dr = 0 Or _
                dc = 0 Or _
                Abs(dr) = Abs(dc) _
            Then


                If PathIsClear( _
                    fromRow, _
                    fromCol, _
                    toRow, _
                    toCol _
                ) Then

                    IsPseudoLegalMove = True
                    Exit Function

                End If

            End If


        ' ====================================================
        ' KING
        ' ====================================================

        Case WHITE_KING, BLACK_KING


            If Abs(dr) <= 1 And _
               Abs(dc) <= 1 Then

                IsPseudoLegalMove = True
                Exit Function

            End If


            If dr = 0 And _
               Abs(dc) = 2 Then


                If CanCastle( _
                    pieceCode, _
                    fromRow, _
                    fromCol, _
                    toRow, _
                    toCol _
                ) Then

                    IsPseudoLegalMove = True
                    Exit Function

                End If

            End If

    End Select

End Function


' ============================================================
' SIMULATE MOVE FOR KING SAFETY
' ============================================================

Private Function SimulatedMoveLeavesKingSafe( _
    ByVal fromRow As Long, _
    ByVal fromCol As Long, _
    ByVal toRow As Long, _
    ByVal toCol As Long _
) As Boolean

    Dim ws As Worksheet

    Dim movingPiece As Long
    Dim targetPiece As Long

    Dim originalFrom As Variant
    Dim originalTo As Variant

    Dim enPassantOriginal As Variant
    Dim enPassantUsed As Boolean

    Dim castleUsed As Boolean

    Dim rookFromCol As Long
    Dim rookToCol As Long

    Dim rookOriginalFrom As Variant
    Dim rookOriginalTo As Variant

    Dim whiteSide As Boolean

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    movingPiece = _
        PieceCodeAt(ws, fromRow, fromCol)

    targetPiece = _
        PieceCodeAt(ws, toRow, toCol)


    whiteSide = _
        IsWhitePieceCode(movingPiece)


    originalFrom = _
        ws.Cells(fromRow, fromCol).Value2

    originalTo = _
        ws.Cells(toRow, toCol).Value2


    enPassantUsed = False
    castleUsed = False


    ' ========================================================
    ' SIMULATE EN PASSANT
    ' ========================================================

    If movingPiece = WHITE_PAWN Or _
       movingPiece = BLACK_PAWN Then


        If targetPiece = 0 And _
           toRow = EnPassantRow And _
           toCol = EnPassantCol And _
           Abs(toCol - fromCol) = 1 Then


            enPassantUsed = True

            enPassantOriginal = _
                ws.Cells(fromRow, toCol).Value2

            ws.Cells( _
                fromRow, _
                toCol _
            ).ClearContents

        End If

    End If


    ' ========================================================
    ' SIMULATE MAIN MOVE
    ' ========================================================

    ws.Cells(toRow, toCol).Value = _
        originalFrom

    ws.Cells(fromRow, fromCol).ClearContents


    ' ========================================================
    ' SIMULATE CASTLING ROOK
    ' ========================================================

    If movingPiece = WHITE_KING Or _
       movingPiece = BLACK_KING Then


        If Abs(toCol - fromCol) = 2 Then

            castleUsed = True


            If toCol = 8 Then

                rookFromCol = 9
                rookToCol = 7

            Else

                rookFromCol = 2
                rookToCol = 5

            End If


            rookOriginalFrom = _
                ws.Cells( _
                    fromRow, _
                    rookFromCol _
                ).Value2

            rookOriginalTo = _
                ws.Cells( _
                    fromRow, _
                    rookToCol _
                ).Value2


            ws.Cells( _
                fromRow, _
                rookToCol _
            ).Value = _
                ws.Cells( _
                    fromRow, _
                    rookFromCol _
                ).Value


            ws.Cells( _
                fromRow, _
                rookFromCol _
            ).ClearContents

        End If

    End If


    SimulatedMoveLeavesKingSafe = _
        Not IsKingInCheckByColor(whiteSide)


    ' ========================================================
    ' RESTORE ORIGINAL POSITION
    ' ========================================================

    ws.Cells(fromRow, fromCol).Value = _
        originalFrom

    ws.Cells(toRow, toCol).Value = _
        originalTo


    If enPassantUsed Then

        ws.Cells( _
            fromRow, _
            toCol _
        ).Value = _
            enPassantOriginal

    End If


    If castleUsed Then

        ws.Cells( _
            fromRow, _
            rookFromCol _
        ).Value = _
            rookOriginalFrom

        ws.Cells( _
            fromRow, _
            rookToCol _
        ).Value = _
            rookOriginalTo

    End If

End Function


' ============================================================
' PATH CLEAR
' ============================================================

Private Function PathIsClear( _
    ByVal fromRow As Long, _
    ByVal fromCol As Long, _
    ByVal toRow As Long, _
    ByVal toCol As Long _
) As Boolean

    Dim ws As Worksheet

    Dim stepRow As Long
    Dim stepCol As Long

    Dim r As Long
    Dim c As Long

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    stepRow = _
        Sgn(toRow - fromRow)

    stepCol = _
        Sgn(toCol - fromCol)


    r = fromRow + stepRow
    c = fromCol + stepCol


    Do While _
        r <> toRow Or _
        c <> toCol


        If PieceCodeAt(ws, r, c) <> 0 Then

            PathIsClear = False
            Exit Function

        End If


        r = r + stepRow
        c = c + stepCol

    Loop


    PathIsClear = True

End Function


' ============================================================
' CHECK
' ============================================================

Private Function IsKingInCheck( _
    ByVal side As String _
) As Boolean

    If side = "WHITE" Then

        IsKingInCheck = _
            IsKingInCheckByColor(True)

    Else

        IsKingInCheck = _
            IsKingInCheckByColor(False)

    End If

End Function


Private Function IsKingInCheckByColor( _
    ByVal whiteSide As Boolean _
) As Boolean

    Dim kingRow As Long
    Dim kingCol As Long


    If Not FindKing( _
        whiteSide, _
        kingRow, _
        kingCol _
    ) Then

        IsKingInCheckByColor = True
        Exit Function

    End If


    IsKingInCheckByColor = _
        IsSquareAttacked( _
            kingRow, _
            kingCol, _
            Not whiteSide _
        )

End Function


Private Function FindKing( _
    ByVal whiteSide As Boolean, _
    ByRef kingRow As Long, _
    ByRef kingCol As Long _
) As Boolean

    Dim ws As Worksheet

    Dim wantedKing As Long

    Dim r As Long
    Dim c As Long

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    If whiteSide Then

        wantedKing = WHITE_KING

    Else

        wantedKing = BLACK_KING

    End If


    For r = BOARD_TOP To BOARD_BOTTOM

        For c = BOARD_LEFT To BOARD_RIGHT


            If PieceCodeAt(ws, r, c) = _
               wantedKing Then

                kingRow = r
                kingCol = c

                FindKing = True
                Exit Function

            End If

        Next c

    Next r


    FindKing = False

End Function


' ============================================================
' ATTACKED SQUARE
' ============================================================

Private Function IsSquareAttacked( _
    ByVal targetRow As Long, _
    ByVal targetCol As Long, _
    ByVal byWhite As Boolean _
) As Boolean

    Dim ws As Worksheet

    Dim r As Long
    Dim c As Long
    Dim i As Long

    Dim pieceCode As Long

    Dim knightRows As Variant
    Dim knightCols As Variant

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    IsSquareAttacked = False


    ' ========================================================
    ' PAWNS
    ' ========================================================

    If byWhite Then

        r = targetRow + 1


        If r <= BOARD_BOTTOM Then


            If targetCol - 1 >= BOARD_LEFT Then

                If PieceCodeAt( _
                    ws, _
                    r, _
                    targetCol - 1 _
                ) = WHITE_PAWN Then

                    IsSquareAttacked = True
                    Exit Function

                End If

            End If


            If targetCol + 1 <= BOARD_RIGHT Then

                If PieceCodeAt( _
                    ws, _
                    r, _
                    targetCol + 1 _
                ) = WHITE_PAWN Then

                    IsSquareAttacked = True
                    Exit Function

                End If

            End If

        End If


    Else

        r = targetRow - 1


        If r >= BOARD_TOP Then


            If targetCol - 1 >= BOARD_LEFT Then

                If PieceCodeAt( _
                    ws, _
                    r, _
                    targetCol - 1 _
                ) = BLACK_PAWN Then

                    IsSquareAttacked = True
                    Exit Function

                End If

            End If


            If targetCol + 1 <= BOARD_RIGHT Then

                If PieceCodeAt( _
                    ws, _
                    r, _
                    targetCol + 1 _
                ) = BLACK_PAWN Then

                    IsSquareAttacked = True
                    Exit Function

                End If

            End If

        End If

    End If


    ' ========================================================
    ' KNIGHTS
    ' ========================================================

    knightRows = _
        Array(-2, -2, -1, -1, 1, 1, 2, 2)

    knightCols = _
        Array(-1, 1, -2, 2, -2, 2, -1, 1)


    For i = 0 To 7

        r = _
            targetRow + knightRows(i)

        c = _
            targetCol + knightCols(i)


        If IsInsideBoard(r, c) Then

            pieceCode = _
                PieceCodeAt(ws, r, c)


            If byWhite Then

                If pieceCode = WHITE_KNIGHT Then

                    IsSquareAttacked = True
                    Exit Function

                End If

            Else

                If pieceCode = BLACK_KNIGHT Then

                    IsSquareAttacked = True
                    Exit Function

                End If

            End If

        End If

    Next i


    ' ========================================================
    ' KINGS
    ' ========================================================

    For r = targetRow - 1 To targetRow + 1

        For c = targetCol - 1 To targetCol + 1


            If IsInsideBoard(r, c) Then


                If Not ( _
                    r = targetRow And _
                    c = targetCol _
                ) Then


                    pieceCode = _
                        PieceCodeAt(ws, r, c)


                    If byWhite Then

                        If pieceCode = WHITE_KING Then

                            IsSquareAttacked = True
                            Exit Function

                        End If

                    Else

                        If pieceCode = BLACK_KING Then

                            IsSquareAttacked = True
                            Exit Function

                        End If

                    End If

                End If

            End If

        Next c

    Next r


    ' ========================================================
    ' ROOK / QUEEN RAYS
    ' ========================================================

    If RayHasAttacker( _
        targetRow, targetCol, _
        -1, 0, _
        byWhite, False _
    ) Then

        IsSquareAttacked = True
        Exit Function

    End If


    If RayHasAttacker( _
        targetRow, targetCol, _
        1, 0, _
        byWhite, False _
    ) Then

        IsSquareAttacked = True
        Exit Function

    End If


    If RayHasAttacker( _
        targetRow, targetCol, _
        0, -1, _
        byWhite, False _
    ) Then

        IsSquareAttacked = True
        Exit Function

    End If


    If RayHasAttacker( _
        targetRow, targetCol, _
        0, 1, _
        byWhite, False _
    ) Then

        IsSquareAttacked = True
        Exit Function

    End If


    ' ========================================================
    ' BISHOP / QUEEN RAYS
    ' ========================================================

    If RayHasAttacker( _
        targetRow, targetCol, _
        -1, -1, _
        byWhite, True _
    ) Then

        IsSquareAttacked = True
        Exit Function

    End If


    If RayHasAttacker( _
        targetRow, targetCol, _
        -1, 1, _
        byWhite, True _
    ) Then

        IsSquareAttacked = True
        Exit Function

    End If


    If RayHasAttacker( _
        targetRow, targetCol, _
        1, -1, _
        byWhite, True _
    ) Then

        IsSquareAttacked = True
        Exit Function

    End If


    If RayHasAttacker( _
        targetRow, targetCol, _
        1, 1, _
        byWhite, True _
    ) Then

        IsSquareAttacked = True
        Exit Function

    End If

End Function


Private Function RayHasAttacker( _
    ByVal startRow As Long, _
    ByVal startCol As Long, _
    ByVal stepRow As Long, _
    ByVal stepCol As Long, _
    ByVal byWhite As Boolean, _
    ByVal diagonal As Boolean _
) As Boolean

    Dim ws As Worksheet

    Dim r As Long
    Dim c As Long

    Dim pieceCode As Long

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    r = startRow + stepRow
    c = startCol + stepCol


    Do While IsInsideBoard(r, c)

        pieceCode = _
            PieceCodeAt(ws, r, c)


        If pieceCode <> 0 Then


            If byWhite Then


                If diagonal Then

                    RayHasAttacker = _
                        pieceCode = WHITE_BISHOP Or _
                        pieceCode = WHITE_QUEEN

                Else

                    RayHasAttacker = _
                        pieceCode = WHITE_ROOK Or _
                        pieceCode = WHITE_QUEEN

                End If


            Else


                If diagonal Then

                    RayHasAttacker = _
                        pieceCode = BLACK_BISHOP Or _
                        pieceCode = BLACK_QUEEN

                Else

                    RayHasAttacker = _
                        pieceCode = BLACK_ROOK Or _
                        pieceCode = BLACK_QUEEN

                End If

            End If


            Exit Function

        End If


        r = r + stepRow
        c = c + stepCol

    Loop


    RayHasAttacker = False

End Function


' ============================================================
' CASTLING
' ============================================================

Private Function CanCastle( _
    ByVal kingPiece As Long, _
    ByVal fromRow As Long, _
    ByVal fromCol As Long, _
    ByVal toRow As Long, _
    ByVal toCol As Long _
) As Boolean

    Dim ws As Worksheet

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    CanCastle = False


    ' ========================================================
    ' WHITE
    ' ========================================================

    If kingPiece = WHITE_KING Then


        If fromRow <> 9 Or _
           fromCol <> 6 Then

            Exit Function

        End If


        If WhiteKingMoved Then
            Exit Function
        End If


        ' ---------------- KINGSIDE ----------------

        If toRow = 9 And _
           toCol = 8 Then


            If WhiteRookHMoved Then
                Exit Function
            End If


            If PieceCodeAt(ws, 9, 9) <> _
               WHITE_ROOK Then

                Exit Function

            End If


            If PieceCodeAt(ws, 9, 7) <> 0 Then
                Exit Function
            End If

            If PieceCodeAt(ws, 9, 8) <> 0 Then
                Exit Function
            End If


            If IsSquareAttacked(9, 6, False) Then
                Exit Function
            End If

            If IsSquareAttacked(9, 7, False) Then
                Exit Function
            End If

            If IsSquareAttacked(9, 8, False) Then
                Exit Function
            End If


            CanCastle = True
            Exit Function

        End If


        ' ---------------- QUEENSIDE ----------------

        If toRow = 9 And _
           toCol = 4 Then


            If WhiteRookAMoved Then
                Exit Function
            End If


            If PieceCodeAt(ws, 9, 2) <> _
               WHITE_ROOK Then

                Exit Function

            End If


            If PieceCodeAt(ws, 9, 5) <> 0 Then
                Exit Function
            End If

            If PieceCodeAt(ws, 9, 4) <> 0 Then
                Exit Function
            End If

            If PieceCodeAt(ws, 9, 3) <> 0 Then
                Exit Function
            End If


            If IsSquareAttacked(9, 6, False) Then
                Exit Function
            End If

            If IsSquareAttacked(9, 5, False) Then
                Exit Function
            End If

            If IsSquareAttacked(9, 4, False) Then
                Exit Function
            End If


            CanCastle = True
            Exit Function

        End If

    End If


    ' ========================================================
    ' BLACK
    ' ========================================================

    If kingPiece = BLACK_KING Then


        If fromRow <> 2 Or _
           fromCol <> 6 Then

            Exit Function

        End If


        If BlackKingMoved Then
            Exit Function
        End If


        ' ---------------- KINGSIDE ----------------

        If toRow = 2 And _
           toCol = 8 Then


            If BlackRookHMoved Then
                Exit Function
            End If


            If PieceCodeAt(ws, 2, 9) <> _
               BLACK_ROOK Then

                Exit Function

            End If


            If PieceCodeAt(ws, 2, 7) <> 0 Then
                Exit Function
            End If

            If PieceCodeAt(ws, 2, 8) <> 0 Then
                Exit Function
            End If


            If IsSquareAttacked(2, 6, True) Then
                Exit Function
            End If

            If IsSquareAttacked(2, 7, True) Then
                Exit Function
            End If

            If IsSquareAttacked(2, 8, True) Then
                Exit Function
            End If


            CanCastle = True
            Exit Function

        End If


        ' ---------------- QUEENSIDE ----------------

        If toRow = 2 And _
           toCol = 4 Then


            If BlackRookAMoved Then
                Exit Function
            End If


            If PieceCodeAt(ws, 2, 2) <> _
               BLACK_ROOK Then

                Exit Function

            End If


            If PieceCodeAt(ws, 2, 5) <> 0 Then
                Exit Function
            End If

            If PieceCodeAt(ws, 2, 4) <> 0 Then
                Exit Function
            End If

            If PieceCodeAt(ws, 2, 3) <> 0 Then
                Exit Function
            End If


            If IsSquareAttacked(2, 6, True) Then
                Exit Function
            End If

            If IsSquareAttacked(2, 5, True) Then
                Exit Function
            End If

            If IsSquareAttacked(2, 4, True) Then
                Exit Function
            End If


            CanCastle = True
            Exit Function

        End If

    End If

End Function


Private Sub UpdateCastlingRights( _
    ByVal movingPiece As Long, _
    ByVal capturedPiece As Long, _
    ByVal fromRow As Long, _
    ByVal fromCol As Long, _
    ByVal toRow As Long, _
    ByVal toCol As Long _
)

    ' ========================================================
    ' KING MOVED
    ' ========================================================

    If movingPiece = WHITE_KING Then
        WhiteKingMoved = True
    End If

    If movingPiece = BLACK_KING Then
        BlackKingMoved = True
    End If


    ' ========================================================
    ' WHITE ROOK MOVED
    ' ========================================================

    If movingPiece = WHITE_ROOK Then


        If fromRow = 9 And _
           fromCol = 2 Then

            WhiteRookAMoved = True

        End If


        If fromRow = 9 And _
           fromCol = 9 Then

            WhiteRookHMoved = True

        End If

    End If


    ' ========================================================
    ' BLACK ROOK MOVED
    ' ========================================================

    If movingPiece = BLACK_ROOK Then


        If fromRow = 2 And _
           fromCol = 2 Then

            BlackRookAMoved = True

        End If


        If fromRow = 2 And _
           fromCol = 9 Then

            BlackRookHMoved = True

        End If

    End If


    ' ========================================================
    ' WHITE ORIGINAL ROOK CAPTURED
    ' ========================================================

    If capturedPiece = WHITE_ROOK Then


        If toRow = 9 And _
           toCol = 2 Then

            WhiteRookAMoved = True

        End If


        If toRow = 9 And _
           toCol = 9 Then

            WhiteRookHMoved = True

        End If

    End If


    ' ========================================================
    ' BLACK ORIGINAL ROOK CAPTURED
    ' ========================================================

    If capturedPiece = BLACK_ROOK Then


        If toRow = 2 And _
           toCol = 2 Then

            BlackRookAMoved = True

        End If


        If toRow = 2 And _
           toCol = 9 Then

            BlackRookHMoved = True

        End If

    End If

End Sub


' ============================================================
' ANY LEGAL MOVE?
' ============================================================

Private Function HasAnyLegalMove( _
    ByVal side As String _
) As Boolean

    Dim ws As Worksheet

    Dim fromRow As Long
    Dim fromCol As Long

    Dim toRow As Long
    Dim toCol As Long

    Dim pieceCode As Long

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    HasAnyLegalMove = False


    For fromRow = BOARD_TOP To BOARD_BOTTOM

        For fromCol = BOARD_LEFT To BOARD_RIGHT


            pieceCode = _
                PieceCodeAt( _
                    ws, _
                    fromRow, _
                    fromCol _
                )


            If BelongsToSide( _
                pieceCode, _
                side _
            ) Then


                For toRow = BOARD_TOP To BOARD_BOTTOM

                    For toCol = BOARD_LEFT To BOARD_RIGHT


                        If IsLegalMove( _
                            fromRow, _
                            fromCol, _
                            toRow, _
                            toCol _
                        ) Then

                            HasAnyLegalMove = True
                            Exit Function

                        End If

                    Next toCol

                Next toRow

            End If

        Next fromCol

    Next fromRow

End Function


' ============================================================
' PROMOTION
' ============================================================

Private Function GetPromotionChoice() As String

    Dim userChoice As String


    userChoice = _
        UCase$( _
            Trim$( _
                InputBox( _
                    "Promote pawn to Q, R, B, or N:", _
                    "Pawn Promotion", _
                    "Q" _
                ) _
            ) _
        )


    Select Case userChoice

        Case "Q", "R", "B", "N"

            GetPromotionChoice = _
                userChoice

        Case Else

            GetPromotionChoice = _
                "Q"

    End Select

End Function


Private Function PromotionPieceCode( _
    ByVal whiteSide As Boolean, _
    ByVal userChoice As String _
) As Long

    If whiteSide Then


        Select Case userChoice

            Case "R"
                PromotionPieceCode = WHITE_ROOK

            Case "B"
                PromotionPieceCode = WHITE_BISHOP

            Case "N"
                PromotionPieceCode = WHITE_KNIGHT

            Case Else
                PromotionPieceCode = WHITE_QUEEN

        End Select


    Else


        Select Case userChoice

            Case "R"
                PromotionPieceCode = BLACK_ROOK

            Case "B"
                PromotionPieceCode = BLACK_BISHOP

            Case "N"
                PromotionPieceCode = BLACK_KNIGHT

            Case Else
                PromotionPieceCode = BLACK_QUEEN

        End Select

    End If

End Function


' ============================================================
' STANDARD ALGEBRAIC NOTATION - SAN
' ============================================================

Private Function BuildSANBase( _
    ByVal pieceCode As Long, _
    ByVal fromRow As Long, _
    ByVal fromCol As Long, _
    ByVal toRow As Long, _
    ByVal toCol As Long, _
    ByVal wasCapture As Boolean, _
    ByVal wasCastle As Boolean, _
    ByVal promotionChoice As String _
) As String

    Dim resultText As String

    Dim destination As String
    Dim pieceLetter As String
    Dim disambiguation As String


    destination = _
        SquareName(toRow, toCol)


    ' ========================================================
    ' CASTLING
    ' ========================================================

    If wasCastle Then


        If toCol = 8 Then

            BuildSANBase = "O-O"

        Else

            BuildSANBase = "O-O-O"

        End If


        Exit Function

    End If


    ' ========================================================
    ' PAWN
    ' ========================================================

    If pieceCode = WHITE_PAWN Or _
       pieceCode = BLACK_PAWN Then


        resultText = ""


        If wasCapture Then

            resultText = _
                FileLetter(fromCol) & "x"

        End If


        resultText = _
            resultText & destination


        If promotionChoice <> "" Then

            resultText = _
                resultText & _
                "=" & _
                promotionChoice

        End If


        BuildSANBase = resultText

        Exit Function

    End If


    ' ========================================================
    ' NORMAL PIECE
    ' ========================================================

    pieceLetter = _
        GetPieceLetter(pieceCode)


    disambiguation = _
        GetSANDisambiguation( _
            pieceCode, _
            fromRow, _
            fromCol, _
            toRow, _
            toCol _
        )


    resultText = _
        pieceLetter & _
        disambiguation


    If wasCapture Then

        resultText = _
            resultText & "x"

    End If


    resultText = _
        resultText & destination


    BuildSANBase = resultText

End Function


Private Function GetPieceLetter( _
    ByVal pieceCode As Long _
) As String

    Select Case pieceCode


        Case WHITE_KING, BLACK_KING

            GetPieceLetter = "K"


        Case WHITE_QUEEN, BLACK_QUEEN

            GetPieceLetter = "Q"


        Case WHITE_ROOK, BLACK_ROOK

            GetPieceLetter = "R"


        Case WHITE_BISHOP, BLACK_BISHOP

            GetPieceLetter = "B"


        Case WHITE_KNIGHT, BLACK_KNIGHT

            GetPieceLetter = "N"


        Case Else

            GetPieceLetter = ""

    End Select

End Function


Private Function GetSANDisambiguation( _
    ByVal pieceCode As Long, _
    ByVal fromRow As Long, _
    ByVal fromCol As Long, _
    ByVal toRow As Long, _
    ByVal toCol As Long _
) As String

    Dim ws As Worksheet

    Dim r As Long
    Dim c As Long

    Dim foundAlternative As Boolean
    Dim sameFileExists As Boolean
    Dim sameRankExists As Boolean

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    foundAlternative = False
    sameFileExists = False
    sameRankExists = False


    For r = BOARD_TOP To BOARD_BOTTOM

        For c = BOARD_LEFT To BOARD_RIGHT


            If Not ( _
                r = fromRow And _
                c = fromCol _
            ) Then


                If PieceCodeAt(ws, r, c) = _
                   pieceCode Then


                    If IsLegalMove( _
                        r, _
                        c, _
                        toRow, _
                        toCol _
                    ) Then


                        foundAlternative = True


                        If c = fromCol Then

                            sameFileExists = True

                        End If


                        If r = fromRow Then

                            sameRankExists = True

                        End If

                    End If

                End If

            End If

        Next c

    Next r


    If Not foundAlternative Then

        GetSANDisambiguation = ""
        Exit Function

    End If


    If Not sameFileExists Then

        GetSANDisambiguation = _
            FileLetter(fromCol)


    ElseIf Not sameRankExists Then

        GetSANDisambiguation = _
            CStr(10 - fromRow)


    Else

        GetSANDisambiguation = _
            FileLetter(fromCol) & _
            CStr(10 - fromRow)

    End If

End Function


Private Function FileLetter( _
    ByVal colNumber As Long _
) As String

    FileLetter = _
        Chr$( _
            Asc("a") + _
            colNumber - _
            BOARD_LEFT _
        )

End Function


' ============================================================
' MOVE HISTORY
' ============================================================

Private Sub RecordMove( _
    ByVal notationText As String _
)

    MoveCount = MoveCount + 1


    If MoveCount = 1 Then

        ReDim MoveList(1 To 1)

    Else

        ReDim Preserve _
            MoveList(1 To MoveCount)

    End If


    MoveList(MoveCount) = _
        notationText

End Sub


Private Sub RefreshHistory()

    Dim ws As Worksheet

    Dim i As Long
    Dim historyRow As Long

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    ws.Range("K15:M200").ClearContents


    For i = 1 To MoveCount


        historyRow = _
            15 + ((i - 1) \ 2)


        ws.Cells( _
            historyRow, _
            11 _
        ).Value = _
            ((i - 1) \ 2) + 1


        If i Mod 2 = 1 Then

            ws.Cells( _
                historyRow, _
                12 _
            ).Value = _
                MoveList(i)

        Else

            ws.Cells( _
                historyRow, _
                13 _
            ).Value = _
                MoveList(i)

        End If

    Next i

End Sub


' ============================================================
' REPETITION
' ============================================================

Private Function AddCurrentPosition() As Long

    Dim positionKey As String

    positionKey = _
        CurrentPositionKey()


    PositionHistoryCount = _
        PositionHistoryCount + 1


    If PositionHistoryCount = 1 Then

        ReDim _
            PositionHistory(1 To 1)

    Else

        ReDim Preserve _
            PositionHistory( _
                1 To PositionHistoryCount _
            )

    End If


    PositionHistory( _
        PositionHistoryCount _
    ) = positionKey


    If PositionCounts Is Nothing Then

        Set PositionCounts = _
            CreateObject( _
                "Scripting.Dictionary" _
            )

    End If


    If PositionCounts.Exists(positionKey) Then

        PositionCounts(positionKey) = _
            CLng( _
                PositionCounts(positionKey) _
            ) + 1

    Else

        PositionCounts.Add _
            positionKey, _
            1

    End If


    AddCurrentPosition = _
        CLng( _
            PositionCounts(positionKey) _
        )

End Function


Private Function CurrentPositionOccurrenceCount() As Long

    Dim positionKey As String

    If PositionCounts Is Nothing Then

        CurrentPositionOccurrenceCount = 0
        Exit Function

    End If


    positionKey = _
        CurrentPositionKey()


    If PositionCounts.Exists(positionKey) Then

        CurrentPositionOccurrenceCount = _
            CLng( _
                PositionCounts(positionKey) _
            )

    Else

        CurrentPositionOccurrenceCount = 0

    End If

End Function


Private Function CurrentPositionKey() As String

    Dim ws As Worksheet

    Dim r As Long
    Dim c As Long

    Dim resultText As String

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    resultText = ""


    For r = BOARD_TOP To BOARD_BOTTOM

        For c = BOARD_LEFT To BOARD_RIGHT

            resultText = _
                resultText & _
                CStr( _
                    PieceCodeAt(ws, r, c) _
                ) & _
                ","

        Next c

    Next r


    resultText = _
        resultText & _
        "|" & _
        CurrentTurn


    resultText = _
        resultText & _
        "|" & _
        CastlingRightsKey()


    resultText = _
        resultText & _
        "|" & _
        RepetitionEnPassantKey()


    CurrentPositionKey = _
        resultText

End Function


Private Function CastlingRightsKey() As String

    Dim ws As Worksheet
    Dim rightsText As String

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    rightsText = ""


    If Not WhiteKingMoved Then


        If Not WhiteRookHMoved And _
           PieceCodeAt(ws, 9, 9) = WHITE_ROOK Then

            rightsText = _
                rightsText & "K"

        End If


        If Not WhiteRookAMoved And _
           PieceCodeAt(ws, 9, 2) = WHITE_ROOK Then

            rightsText = _
                rightsText & "Q"

        End If

    End If


    If Not BlackKingMoved Then


        If Not BlackRookHMoved And _
           PieceCodeAt(ws, 2, 9) = BLACK_ROOK Then

            rightsText = _
                rightsText & "k"

        End If


        If Not BlackRookAMoved And _
           PieceCodeAt(ws, 2, 2) = BLACK_ROOK Then

            rightsText = _
                rightsText & "q"

        End If

    End If


    If rightsText = "" Then
        rightsText = "-"
    End If


    CastlingRightsKey = _
        rightsText

End Function


Private Function RepetitionEnPassantKey() As String

    If EnPassantRow = 0 Or _
       EnPassantCol = 0 Then

        RepetitionEnPassantKey = "-"
        Exit Function

    End If


    If HasLegalEnPassantCapture() Then

        RepetitionEnPassantKey = _
            SquareName( _
                EnPassantRow, _
                EnPassantCol _
            )

    Else

        RepetitionEnPassantKey = "-"

    End If

End Function


Private Function HasLegalEnPassantCapture() As Boolean

    Dim ws As Worksheet

    Dim sourceRow As Long
    Dim leftCol As Long
    Dim rightCol As Long

    Dim wantedPawn As Long

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    HasLegalEnPassantCapture = False


    If EnPassantRow = 0 Then
        Exit Function
    End If


    If CurrentTurn = "WHITE" Then

        sourceRow = _
            EnPassantRow + 1

        wantedPawn = _
            WHITE_PAWN

    Else

        sourceRow = _
            EnPassantRow - 1

        wantedPawn = _
            BLACK_PAWN

    End If


    leftCol = _
        EnPassantCol - 1

    rightCol = _
        EnPassantCol + 1


    If IsInsideBoard( _
        sourceRow, _
        leftCol _
    ) Then


        If PieceCodeAt( _
            ws, _
            sourceRow, _
            leftCol _
        ) = wantedPawn Then


            If IsLegalMove( _
                sourceRow, _
                leftCol, _
                EnPassantRow, _
                EnPassantCol _
            ) Then

                HasLegalEnPassantCapture = True
                Exit Function

            End If

        End If

    End If


    If IsInsideBoard( _
        sourceRow, _
        rightCol _
    ) Then


        If PieceCodeAt( _
            ws, _
            sourceRow, _
            rightCol _
        ) = wantedPawn Then


            If IsLegalMove( _
                sourceRow, _
                rightCol, _
                EnPassantRow, _
                EnPassantCol _
            ) Then

                HasLegalEnPassantCapture = True
                Exit Function

            End If

        End If

    End If

End Function


Private Sub RebuildPositionCounts()

    Dim i As Long
    Dim positionKey As String


    Set PositionCounts = _
        CreateObject( _
            "Scripting.Dictionary" _
        )


    For i = 1 To PositionHistoryCount


        positionKey = _
            PositionHistory(i)


        If PositionCounts.Exists(positionKey) Then

            PositionCounts(positionKey) = _
                CLng( _
                    PositionCounts(positionKey) _
                ) + 1

        Else

            PositionCounts.Add _
                positionKey, _
                1

        End If

    Next i

End Sub


' ============================================================
' INSUFFICIENT MATERIAL / BASIC DEAD POSITIONS
' ============================================================

Private Function IsInsufficientMaterial() As Boolean

    Dim ws As Worksheet

    Dim r As Long
    Dim c As Long

    Dim pieceCode As Long

    Dim bishopCount As Long
    Dim knightCount As Long

    Dim firstBishopColor As Long
    Dim bishopColor As Long

    Dim bishopsSameColor As Boolean

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    bishopCount = 0
    knightCount = 0

    firstBishopColor = -1
    bishopsSameColor = True


    For r = BOARD_TOP To BOARD_BOTTOM

        For c = BOARD_LEFT To BOARD_RIGHT


            pieceCode = _
                PieceCodeAt(ws, r, c)


            Select Case pieceCode


                Case _
                    WHITE_PAWN, BLACK_PAWN, _
                    WHITE_ROOK, BLACK_ROOK, _
                    WHITE_QUEEN, BLACK_QUEEN


                    IsInsufficientMaterial = False
                    Exit Function


                Case _
                    WHITE_KNIGHT, BLACK_KNIGHT


                    knightCount = _
                        knightCount + 1


                Case _
                    WHITE_BISHOP, BLACK_BISHOP


                    bishopCount = _
                        bishopCount + 1


                    bishopColor = _
                        (r + c) Mod 2


                    If firstBishopColor = -1 Then

                        firstBishopColor = _
                            bishopColor

                    ElseIf bishopColor <> _
                           firstBishopColor Then

                        bishopsSameColor = False

                    End If

            End Select

        Next c

    Next r


    ' King vs King

    If bishopCount = 0 And _
       knightCount = 0 Then

        IsInsufficientMaterial = True
        Exit Function

    End If


    ' King + Bishop vs King
    ' King + Knight vs King

    If bishopCount + knightCount = 1 Then

        IsInsufficientMaterial = True
        Exit Function

    End If


    ' Only bishops and all bishops live on same-color squares

    If knightCount = 0 And _
       bishopCount > 0 And _
       bishopsSameColor Then

        IsInsufficientMaterial = True
        Exit Function

    End If


    IsInsufficientMaterial = False

End Function


' ============================================================
' UNDO
' ============================================================

Private Sub PushUndoSnapshot()

    UndoCount = _
        UndoCount + 1


    If UndoCount = 1 Then

        ReDim UndoStack(1 To 1)

    Else

        ReDim Preserve _
            UndoStack( _
                1 To UndoCount _
            )

    End If


    UndoStack(UndoCount) = _
        BuildSnapshot()

End Sub


Private Function BuildSnapshot() As String

    BuildSnapshot = _
        BoardStateString() & "|" & _
        CurrentTurn & "|" & _
        GameStatus & "|" & _
        CStr(GameOver) & "|" & _
        CStr(HalfmoveClock) & "|" & _
        CStr(EnPassantRow) & "|" & _
        CStr(EnPassantCol) & "|" & _
        CStr(WhiteKingMoved) & "|" & _
        CStr(BlackKingMoved) & "|" & _
        CStr(WhiteRookAMoved) & "|" & _
        CStr(WhiteRookHMoved) & "|" & _
        CStr(BlackRookAMoved) & "|" & _
        CStr(BlackRookHMoved) & "|" & _
        CStr(MoveCount) & "|" & _
        CStr(PositionHistoryCount) & "|" & _
        CStr(LastFromRow) & "|" & _
        CStr(LastFromCol) & "|" & _
        CStr(LastToRow) & "|" & _
        CStr(LastToCol)

End Function


Private Function BoardStateString() As String

    Dim ws As Worksheet

    Dim r As Long
    Dim c As Long

    Dim firstValue As Boolean

    Dim resultText As String

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    firstValue = True
    resultText = ""


    For r = BOARD_TOP To BOARD_BOTTOM

        For c = BOARD_LEFT To BOARD_RIGHT


            If Not firstValue Then

                resultText = _
                    resultText & ","

            End If


            resultText = _
                resultText & _
                CStr( _
                    PieceCodeAt(ws, r, c) _
                )


            firstValue = False

        Next c

    Next r


    BoardStateString = _
        resultText

End Function


Public Sub UndoMove()

    Dim snapshotText As String


    If UndoCount <= 0 Then

        MsgBox _
            "There is no move to undo.", _
            vbInformation

        Exit Sub

    End If


    snapshotText = _
        UndoStack(UndoCount)


    UndoCount = _
        UndoCount - 1


    If UndoCount = 0 Then

        Erase UndoStack

    Else

        ReDim Preserve _
            UndoStack( _
                1 To UndoCount _
            )

    End If


    RestoreSnapshot snapshotText

End Sub


Private Sub RestoreSnapshot( _
    ByVal snapshotText As String _
)

    Dim ws As Worksheet

    Dim parts As Variant
    Dim boardCodes As Variant

    Dim r As Long
    Dim c As Long

    Dim indexValue As Long
    Dim codeValue As Long

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    Application.EnableEvents = False
    Application.ScreenUpdating = False


    parts = _
        Split(snapshotText, "|")

    boardCodes = _
        Split(parts(0), ",")


    indexValue = 0


    For r = BOARD_TOP To BOARD_BOTTOM

        For c = BOARD_LEFT To BOARD_RIGHT


            codeValue = _
                CLng( _
                    boardCodes(indexValue) _
                )


            If codeValue = 0 Then

                ws.Cells(r, c).ClearContents

            Else

                ws.Cells(r, c).Value = _
                    ChrW(codeValue)

            End If


            indexValue = _
                indexValue + 1

        Next c

    Next r


    CurrentTurn = _
        CStr(parts(1))

    GameStatus = _
        CStr(parts(2))

    GameOver = _
        ParseBool(parts(3))


    HalfmoveClock = _
        CLng(parts(4))


    EnPassantRow = _
        CLng(parts(5))

    EnPassantCol = _
        CLng(parts(6))


    WhiteKingMoved = _
        ParseBool(parts(7))

    BlackKingMoved = _
        ParseBool(parts(8))


    WhiteRookAMoved = _
        ParseBool(parts(9))

    WhiteRookHMoved = _
        ParseBool(parts(10))


    BlackRookAMoved = _
        ParseBool(parts(11))

    BlackRookHMoved = _
        ParseBool(parts(12))


    MoveCount = _
        CLng(parts(13))


    PositionHistoryCount = _
        CLng(parts(14))


    LastFromRow = _
        CLng(parts(15))

    LastFromCol = _
        CLng(parts(16))

    LastToRow = _
        CLng(parts(17))

    LastToCol = _
        CLng(parts(18))


    If MoveCount = 0 Then

        Erase MoveList

    Else

        ReDim Preserve _
            MoveList( _
                1 To MoveCount _
            )

    End If


    If PositionHistoryCount > 0 Then

        ReDim Preserve _
            PositionHistory( _
                1 To PositionHistoryCount _
            )

    End If


    RebuildPositionCounts


    SelectedRow = 0
    SelectedCol = 0


    ResetBoardColors

    HighlightLastMove
    HighlightCheckedKing


    RefreshHistory
    UpdatePanel


    ws.Activate
    ws.Range("A1").Select


    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub


Private Function ParseBool( _
    ByVal valueText As String _
) As Boolean

    ParseBool = _
        LCase$(Trim$(valueText)) = "true"

End Function


' ============================================================
' COLORS
' ============================================================

Public Sub ResetBoardColors()

    Dim ws As Worksheet

    Dim r As Long
    Dim c As Long

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    For r = BOARD_TOP To BOARD_BOTTOM

        For c = BOARD_LEFT To BOARD_RIGHT


            If (r + c) Mod 2 = 0 Then

                ' Soft warm light square
                ws.Cells(r, c).Interior.Color = _
                    RGB(248, 246, 241)

            Else

                ' Soft sage square
                ws.Cells(r, c).Interior.Color = _
                    RGB(218, 225, 220)

            End If

        Next c

    Next r

End Sub


Private Sub HighlightLastMove()

    Dim ws As Worksheet

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    If LastFromRow = 0 Then
        Exit Sub
    End If


    ws.Cells( _
        LastFromRow, _
        LastFromCol _
    ).Interior.Color = _
        RGB(220, 228, 239)


    ws.Cells( _
        LastToRow, _
        LastToCol _
    ).Interior.Color = _
        RGB(220, 228, 239)

End Sub


Private Sub HighlightCheckedKing()

    Dim ws As Worksheet

    Dim kingRow As Long
    Dim kingCol As Long

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    If CurrentTurn = "" Then
        Exit Sub
    End If


    If IsKingInCheck(CurrentTurn) Then


        If CurrentTurn = "WHITE" Then


            If FindKing( _
                True, _
                kingRow, _
                kingCol _
            ) Then

                ws.Cells( _
                    kingRow, _
                    kingCol _
                ).Interior.Color = _
                    RGB(242, 184, 184)

            End If


        Else


            If FindKing( _
                False, _
                kingRow, _
                kingCol _
            ) Then

                ws.Cells( _
                    kingRow, _
                    kingCol _
                ).Interior.Color = _
                    RGB(242, 184, 184)

            End If

        End If

    End If

End Sub


' ============================================================
' DASHBOARD
'
' Existing layout remains:
'
' K4 = Turn label
' L4:M4 = Turn value
'
' K5 = Status label
' L5:M5 = Status value
'
' K6 = Selected label
' L6:M6 = Selected value
'
' K8:M9 = NEW GAME
' K10:M11 = UNDO
'
' K14:M14 = Move History headings
' ============================================================

Private Sub UpdatePanel()

    Dim ws As Worksheet

    Dim displayStatus As String
    Dim claimText As String

    Set ws = _
        ThisWorkbook.Worksheets("Chess Board")


    ' ========================================================
    ' TURN
    ' ========================================================

    If CurrentTurn = "WHITE" Then

        ws.Range("L4").Value = _
            "White"

    ElseIf CurrentTurn = "BLACK" Then

        ws.Range("L4").Value = _
            "Black"

    Else

        ws.Range("L4").Value = _
            "-"

    End If


    ' ========================================================
    ' STATUS
    ' ========================================================

    displayStatus = _
        GameStatus


    If Not GameOver Then

        claimText = _
            GetDrawClaimText()


        If claimText <> "" Then


            If GameStatus = "Check" Then

                displayStatus = _
                    "Check - " & claimText

            Else

                displayStatus = _
                    claimText

            End If

        End If

    End If


    ws.Range("L5").Value = _
        displayStatus


    ' ========================================================
    ' SELECTED
    ' ========================================================

    If SelectedRow = 0 Then

        ws.Range("L6").Value = _
            "-"

    Else

        ws.Range("L6").Value = _
            SquareName( _
                SelectedRow, _
                SelectedCol _
            )

    End If

End Sub


' ============================================================
' UTILITY FUNCTIONS
' ============================================================

Private Function PieceCodeAt( _
    ByVal ws As Worksheet, _
    ByVal r As Long, _
    ByVal c As Long _
) As Long

    Dim cellText As String


    cellText = _
        CStr( _
            ws.Cells(r, c).Value2 _
        )


    If Len(cellText) = 0 Then

        PieceCodeAt = 0

    Else

        PieceCodeAt = _
            AscW( _
                Left$(cellText, 1) _
            )

    End If

End Function


Private Function IsWhitePieceCode( _
    ByVal pieceCode As Long _
) As Boolean

    IsWhitePieceCode = _
        pieceCode >= WHITE_KING And _
        pieceCode <= WHITE_PAWN

End Function


Private Function IsBlackPieceCode( _
    ByVal pieceCode As Long _
) As Boolean

    IsBlackPieceCode = _
        pieceCode >= BLACK_KING And _
        pieceCode <= BLACK_PAWN

End Function


Private Function BelongsToSide( _
    ByVal pieceCode As Long, _
    ByVal side As String _
) As Boolean

    If pieceCode = 0 Then

        BelongsToSide = False
        Exit Function

    End If


    If side = "WHITE" Then

        BelongsToSide = _
            IsWhitePieceCode(pieceCode)

    Else

        BelongsToSide = _
            IsBlackPieceCode(pieceCode)

    End If

End Function


Private Function IsInsideBoard( _
    ByVal r As Long, _
    ByVal c As Long _
) As Boolean

    IsInsideBoard = _
        r >= BOARD_TOP And _
        r <= BOARD_BOTTOM And _
        c >= BOARD_LEFT And _
        c <= BOARD_RIGHT

End Function


Private Function SquareName( _
    ByVal r As Long, _
    ByVal c As Long _
) As String

    SquareName = _
        FileLetter(c) & _
        CStr(10 - r)

End Function


' ============================================================
' EMERGENCY EVENT REPAIR
' ============================================================

Public Sub RepairEvents()

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox _
        "Excel events are enabled again.", _
        vbInformation

End Sub

