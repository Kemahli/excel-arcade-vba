Attribute VB_Name = "Module2"
Option Explicit

' ============================================================
' EXCEL ARCADE - BLACKJACK
' Sheet: Blackjack
' 2 decks = 104 real cards
' Players: 1-4
' ============================================================

Private Const BJ_MAX_CARDS As Long = 14
Private Const BJ_SHOE_SIZE As Long = 104


' ============================================================
' SHOE
' ============================================================

Public BJ_Shoe(1 To BJ_SHOE_SIZE) As String
Public BJ_ShoePosition As Long
Public BJ_ShoeReady As Boolean


' ============================================================
' ROUND
' ============================================================

Public BJ_RoundActive As Boolean
Public BJ_RoundComplete As Boolean

Public BJ_DealerActive As Boolean
Public BJ_DealerHidden As Boolean

Public BJ_CurrentPlayer As Long
Public BJ_CurrentHand As Long


' ============================================================
' DEALER
' ============================================================

Public BJ_DealerCards(1 To BJ_MAX_CARDS) As String
Public BJ_DealerCount As Long


' ============================================================
' PLAYERS
' ============================================================

Public BJ_PlayerActive(1 To 4) As Boolean

Public BJ_PlayerCards( _
    1 To 4, _
    1 To 2, _
    1 To BJ_MAX_CARDS _
) As String

Public BJ_PlayerCardCount(1 To 4, 1 To 2) As Long
Public BJ_PlayerHandCount(1 To 4) As Long

Public BJ_HandStatus(1 To 4, 1 To 2) As String
Public BJ_InitialNatural(1 To 4, 1 To 2) As Boolean
Public BJ_Doubled(1 To 4, 1 To 2) As Boolean


' ============================================================
' MERGED CELL HELPERS
' ============================================================

Private Sub BJ_ClearMergedCell(ByVal targetCell As Range)

    targetCell.MergeArea.ClearContents

End Sub


Private Sub BJ_SetMergedValue( _
    ByVal targetCell As Range, _
    ByVal newValue As Variant _
)

    targetCell.MergeArea.Cells(1, 1).Value = newValue

End Sub


' ============================================================
' INITIALIZE
' ============================================================

Public Sub BJ_Initialize()

    Dim ws As Worksheet

    On Error GoTo InitError

    Set ws = ThisWorkbook.Worksheets("Blackjack")

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    BJ_SetupLabels
    BJ_SetupDropdowns
    BJ_FormatSheet

    ' Defaults
    If Trim$(CStr(ws.Range("I4").Value)) = "" Then
        ws.Range("I4").Value = "IN"
    End If

    If Trim$(CStr(ws.Range("I9").Value)) = "" Then
        ws.Range("I9").Value = "IN"
    End If

    If Trim$(CStr(ws.Range("I14").Value)) = "" Then
        ws.Range("I14").Value = "OUT"
    End If

    If Trim$(CStr(ws.Range("I19").Value)) = "" Then
        ws.Range("I19").Value = "OUT"
    End If

    If Trim$(CStr(ws.Range("I24").Value)) = "" Then
        ws.Range("I24").Value = "OUT"
    End If

    BJ_NewShoe

    ws.Activate
    ws.Range("A1").Select

InitExit:

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

InitError:

    MsgBox "Blackjack initialization error: " & Err.Description, vbExclamation
    Resume InitExit

End Sub


' ============================================================
' LABELS
' ============================================================

Private Sub BJ_SetupLabels()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Blackjack")

    BJ_SetMergedValue ws.Range("B1"), "BLACKJACK"

    BJ_SetMergedValue ws.Range("B4"), "DEALER"
    BJ_SetMergedValue ws.Range("G4"), "IN PLAY"
    BJ_SetMergedValue ws.Range("K4"), "STATUS"
    BJ_SetMergedValue ws.Range("P5"), "TOTAL"

    BJ_SetMergedValue ws.Range("B9"), "PLAYER 1"
    BJ_SetMergedValue ws.Range("G9"), "IN PLAY"
    BJ_SetMergedValue ws.Range("K9"), "STATUS"
    BJ_SetMergedValue ws.Range("P10"), "TOTAL"
    BJ_SetMergedValue ws.Range("P11"), "SPLIT"

    BJ_SetMergedValue ws.Range("B14"), "PLAYER 2"
    BJ_SetMergedValue ws.Range("G14"), "IN PLAY"
    BJ_SetMergedValue ws.Range("K14"), "STATUS"
    BJ_SetMergedValue ws.Range("P15"), "TOTAL"
    BJ_SetMergedValue ws.Range("P16"), "SPLIT"

    BJ_SetMergedValue ws.Range("B19"), "PLAYER 3"
    BJ_SetMergedValue ws.Range("G19"), "IN PLAY"
    BJ_SetMergedValue ws.Range("K19"), "STATUS"
    BJ_SetMergedValue ws.Range("P20"), "TOTAL"
    BJ_SetMergedValue ws.Range("P21"), "SPLIT"

    BJ_SetMergedValue ws.Range("B24"), "PLAYER 4"
    BJ_SetMergedValue ws.Range("G24"), "IN PLAY"
    BJ_SetMergedValue ws.Range("K24"), "STATUS"
    BJ_SetMergedValue ws.Range("P25"), "TOTAL"
    BJ_SetMergedValue ws.Range("P26"), "SPLIT"

    BJ_SetMergedValue ws.Range("U3"), "TABLE CONTROL"

    BJ_SetMergedValue ws.Range("U5"), "SHOE"
    BJ_SetMergedValue ws.Range("W5"), "2 DECKS"

    BJ_SetMergedValue ws.Range("U6"), "CARDS LEFT"
    BJ_SetMergedValue ws.Range("U7"), "TURN"

    BJ_SetMergedValue ws.Range("U9"), "DEAL / NEXT ROUND"
    BJ_SetMergedValue ws.Range("U12"), "HIT"
    BJ_SetMergedValue ws.Range("U15"), "STAND"
    BJ_SetMergedValue ws.Range("U18"), "DOUBLE"
    BJ_SetMergedValue ws.Range("U21"), "SPLIT"
    BJ_SetMergedValue ws.Range("U24"), "NEW SHOE"

End Sub


' ============================================================
' DROPDOWNS
' ============================================================

Private Sub BJ_SetupDropdowns()

    Dim ws As Worksheet
    Dim separatorText As String

    Set ws = ThisWorkbook.Worksheets("Blackjack")

    separatorText = Application.International(xlListSeparator)

    BJ_AddInOutValidation ws.Range("I4"), separatorText
    BJ_AddInOutValidation ws.Range("I9"), separatorText
    BJ_AddInOutValidation ws.Range("I14"), separatorText
    BJ_AddInOutValidation ws.Range("I19"), separatorText
    BJ_AddInOutValidation ws.Range("I24"), separatorText

End Sub


Private Sub BJ_AddInOutValidation( _
    ByVal targetCell As Range, _
    ByVal separatorText As String _
)

    On Error Resume Next

    targetCell.Validation.Delete

    targetCell.Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:="IN" & separatorText & "OUT"

    targetCell.Validation.IgnoreBlank = True
    targetCell.Validation.InCellDropdown = True

    On Error GoTo 0

End Sub


' ============================================================
' FORMATTING
' ============================================================

Private Sub BJ_FormatSheet()

    Dim ws As Worksheet
    Dim buttonRange As Range

    Set ws = ThisWorkbook.Worksheets("Blackjack")

    With ws.Range("B1:S2")
        .Font.Bold = True
        .Font.Size = 20
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    BJ_FormatCardRange ws.Range("B5:O5")

    BJ_FormatCardRange ws.Range("B10:O10")
    BJ_FormatCardRange ws.Range("B11:O11")

    BJ_FormatCardRange ws.Range("B15:O15")
    BJ_FormatCardRange ws.Range("B16:O16")

    BJ_FormatCardRange ws.Range("B20:O20")
    BJ_FormatCardRange ws.Range("B21:O21")

    BJ_FormatCardRange ws.Range("B25:O25")
    BJ_FormatCardRange ws.Range("B26:O26")

    Set buttonRange = Union( _
        ws.Range("U9:X10"), _
        ws.Range("U12:X13"), _
        ws.Range("U15:X16"), _
        ws.Range("U18:X19"), _
        ws.Range("U21:X22"), _
        ws.Range("U24:X25") _
    )

    With buttonRange

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Font.Bold = True
        .Font.Size = 11

        .Interior.Color = RGB(229, 235, 231)

        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(175, 185, 178)

    End With

End Sub


Private Sub BJ_FormatCardRange(ByVal cardRange As Range)

    With cardRange

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Font.Name = "Segoe UI Symbol"
        .Font.Size = 18
        .Font.Bold = True

        .Interior.Color = RGB(250, 250, 248)

        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(220, 220, 215)

    End With

End Sub


' ============================================================
' SHEET CLICK
' ============================================================

Public Sub BJ_HandleClick(ByVal Target As Range)

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Blackjack")

    If Not BJ_ShoeReady Then

        BJ_Initialize
        Exit Sub

    End If


    If Not Intersect(Target, ws.Range("U9:X10")) Is Nothing Then

        BJ_DealNextRound
        BJ_ResetSelection
        Exit Sub

    End If


    If Not Intersect(Target, ws.Range("U12:X13")) Is Nothing Then

        BJ_Hit
        BJ_ResetSelection
        Exit Sub

    End If


    If Not Intersect(Target, ws.Range("U15:X16")) Is Nothing Then

        BJ_Stand
        BJ_ResetSelection
        Exit Sub

    End If


    If Not Intersect(Target, ws.Range("U18:X19")) Is Nothing Then

        BJ_Double
        BJ_ResetSelection
        Exit Sub

    End If


    If Not Intersect(Target, ws.Range("U21:X22")) Is Nothing Then

        BJ_Split
        BJ_ResetSelection
        Exit Sub

    End If


    If Not Intersect(Target, ws.Range("U24:X25")) Is Nothing Then

        BJ_NewShoe
        BJ_ResetSelection
        Exit Sub

    End If

End Sub


Private Sub BJ_ResetSelection()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Blackjack")

    On Error GoTo ResetExit

    Application.EnableEvents = False

    ws.Range("A1").Select

ResetExit:

    Application.EnableEvents = True

End Sub


' ============================================================
' NEW SHOE
' ============================================================

Public Sub BJ_NewShoe()

    Dim ranks As Variant
    Dim suits As Variant

    Dim deckNumber As Long
    Dim r As Long
    Dim s As Long

    Dim indexValue As Long

    Dim i As Long
    Dim j As Long

    Dim tempCard As String

    ranks = Array( _
        "A", "2", "3", "4", "5", "6", "7", _
        "8", "9", "10", "J", "Q", "K" _
    )

    suits = Array("S", "H", "D", "C")

    indexValue = 1


    ' Exactly two complete decks
    For deckNumber = 1 To 2

        For s = LBound(suits) To UBound(suits)

            For r = LBound(ranks) To UBound(ranks)

                BJ_Shoe(indexValue) = _
                    CStr(ranks(r)) & "-" & CStr(suits(s))

                indexValue = indexValue + 1

            Next r

        Next s

    Next deckNumber


    ' Fisher-Yates shuffle
    Randomize Timer

    For i = BJ_SHOE_SIZE To 2 Step -1

        j = Int(Rnd() * i) + 1

        tempCard = BJ_Shoe(i)
        BJ_Shoe(i) = BJ_Shoe(j)
        BJ_Shoe(j) = tempCard

    Next i


    BJ_ShoePosition = 1
    BJ_ShoeReady = True

    BJ_ClearRoundState
    BJ_ClearRoundDisplay

    BJ_UpdateCardsLeft

    BJ_SetDealerStatus "READY"
    BJ_SetPlayerReadyStatuses

    BJ_UpdateTurnDisplay "-"

End Sub


' ============================================================
' CLEAR ROUND STATE
' ============================================================

Private Sub BJ_ClearRoundState()

    Erase BJ_DealerCards
    BJ_DealerCount = 0

    Erase BJ_PlayerActive
    Erase BJ_PlayerCards
    Erase BJ_PlayerCardCount
    Erase BJ_PlayerHandCount
    Erase BJ_HandStatus
    Erase BJ_InitialNatural
    Erase BJ_Doubled

    BJ_DealerActive = False
    BJ_DealerHidden = False

    BJ_CurrentPlayer = 0
    BJ_CurrentHand = 0

    BJ_RoundActive = False
    BJ_RoundComplete = False

End Sub


' ============================================================
' CLEAR ROUND DISPLAY
'
' IMPORTANT:
' All merged ranges are cleared as COMPLETE merged areas.
' ============================================================

Private Sub BJ_ClearRoundDisplay()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Blackjack")

    ' Dealer
    ws.Range("B5:O5").ClearContents
    ws.Range("R5:S5").ClearContents
    ws.Range("B6:S6").ClearContents
    ws.Range("M4:S4").ClearContents

    ' Player 1
    ws.Range("B10:O10").ClearContents
    ws.Range("B11:O11").ClearContents
    ws.Range("R10:S10").ClearContents
    ws.Range("R11:S11").ClearContents
    ws.Range("M9:S9").ClearContents
    ws.Range("B12:S12").ClearContents

    ' Player 2
    ws.Range("B15:O15").ClearContents
    ws.Range("B16:O16").ClearContents
    ws.Range("R15:S15").ClearContents
    ws.Range("R16:S16").ClearContents
    ws.Range("M14:S14").ClearContents
    ws.Range("B17:S17").ClearContents

    ' Player 3
    ws.Range("B20:O20").ClearContents
    ws.Range("B21:O21").ClearContents
    ws.Range("R20:S20").ClearContents
    ws.Range("R21:S21").ClearContents
    ws.Range("M19:S19").ClearContents
    ws.Range("B22:S22").ClearContents

    ' Player 4
    ws.Range("B25:O25").ClearContents
    ws.Range("B26:O26").ClearContents
    ws.Range("R25:S25").ClearContents
    ws.Range("R26:S26").ClearContents
    ws.Range("M24:S24").ClearContents
    ws.Range("B27:S27").ClearContents

End Sub


' ============================================================
' DEAL NEXT ROUND
' ============================================================

Public Sub BJ_DealNextRound()

    Dim ws As Worksheet

    Dim p As Long
    Dim activePlayers As Long
    Dim requiredCards As Long

    Set ws = ThisWorkbook.Worksheets("Blackjack")


    If BJ_RoundActive And Not BJ_RoundComplete Then

        MsgBox _
            "Finish the current round first.", _
            vbInformation

        Exit Sub

    End If


    activePlayers = 0

    For p = 1 To 4

        If BJ_PlayerIsSelectedIn(p) Then
            activePlayers = activePlayers + 1
        End If

    Next p


    If activePlayers = 0 Then

        MsgBox _
            "At least one player must be IN.", _
            vbInformation

        Exit Sub

    End If


    requiredCards = activePlayers * 2


    If UCase$(Trim$(CStr(ws.Range("I4").Value))) = "IN" Then
        requiredCards = requiredCards + 2
    End If


    If BJ_CardsLeft() < requiredCards Then

        MsgBox _
            "Not enough cards remain for a new round. Press NEW SHOE.", _
            vbInformation

        Exit Sub

    End If


    BJ_ClearRoundState
    BJ_ClearRoundDisplay


    ' Dealer selection
    BJ_DealerActive = _
        UCase$(Trim$(CStr(ws.Range("I4").Value))) = "IN"


    ' Players
    For p = 1 To 4

        BJ_PlayerActive(p) = BJ_PlayerIsSelectedIn(p)

        If BJ_PlayerActive(p) Then

            BJ_PlayerHandCount(p) = 1
            BJ_HandStatus(p, 1) = "PLAYING"

        Else

            BJ_PlayerHandCount(p) = 0

        End If

    Next p


    BJ_RoundActive = True
    BJ_RoundComplete = False


    ' First card to players
    For p = 1 To 4

        If BJ_PlayerActive(p) Then

            BJ_AddCardToPlayer _
                p, _
                1, _
                BJ_DrawCard()

        End If

    Next p


    ' Dealer first card:
    ' already determined but hidden
    If BJ_DealerActive Then

        BJ_AddCardToDealer BJ_DrawCard()

    End If


    ' Second cards
    For p = 1 To 4

        If BJ_PlayerActive(p) Then

            BJ_AddCardToPlayer _
                p, _
                1, _
                BJ_DrawCard()

        End If

    Next p


    If BJ_DealerActive Then

        BJ_AddCardToDealer BJ_DrawCard()

        BJ_DealerHidden = True

    End If


    ' Natural blackjacks
    For p = 1 To 4

        If BJ_PlayerActive(p) Then

            If BJ_PlayerTotal(p, 1) = 21 Then

                BJ_InitialNatural(p, 1) = True
                BJ_HandStatus(p, 1) = "BLACKJACK"

            End If

        End If

    Next p


    BJ_RenderAll


    If BJ_DealerActive Then

        BJ_SetDealerStatus "HOLE CARD HIDDEN"


        If BJ_IsDealerBlackjack() Then

            BJ_DealerHidden = False

            BJ_RenderDealer

            BJ_SetDealerStatus "BLACKJACK"

            BJ_SettleRound

            Exit Sub

        End If

    Else

        BJ_SetDealerStatus "OUT"

    End If


    BJ_StartFirstTurn

End Sub


' ============================================================
' PLAYER IN / OUT
' ============================================================

Private Function BJ_PlayerIsSelectedIn( _
    ByVal playerNumber As Long _
) As Boolean

    Dim ws As Worksheet
    Dim rowNumber As Long

    Set ws = ThisWorkbook.Worksheets("Blackjack")

    rowNumber = BJ_PlayerHeaderRow(playerNumber)

    BJ_PlayerIsSelectedIn = _
        UCase$(Trim$(CStr(ws.Cells(rowNumber, 9).Value))) = "IN"

End Function


' ============================================================
' DRAW CARD
' ============================================================

Private Function BJ_DrawCard() As String

    If BJ_ShoePosition > BJ_SHOE_SIZE Then

        BJ_DrawCard = ""
        Exit Function

    End If


    BJ_DrawCard = BJ_Shoe(BJ_ShoePosition)

    BJ_ShoePosition = BJ_ShoePosition + 1

    BJ_UpdateCardsLeft

End Function


Private Function BJ_CardsLeft() As Long

    BJ_CardsLeft = _
        BJ_SHOE_SIZE - BJ_ShoePosition + 1

    If BJ_CardsLeft < 0 Then
        BJ_CardsLeft = 0
    End If

End Function


Private Sub BJ_UpdateCardsLeft()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Blackjack")

    BJ_SetMergedValue _
        ws.Range("W6"), _
        BJ_CardsLeft()

End Sub


' ============================================================
' ADD CARDS
' ============================================================

Private Sub BJ_AddCardToPlayer( _
    ByVal playerNumber As Long, _
    ByVal handNumber As Long, _
    ByVal cardText As String _
)

    If cardText = "" Then Exit Sub

    If BJ_PlayerCardCount( _
        playerNumber, _
        handNumber _
    ) >= BJ_MAX_CARDS Then

        Exit Sub

    End If


    BJ_PlayerCardCount( _
        playerNumber, _
        handNumber _
    ) = _
        BJ_PlayerCardCount( _
            playerNumber, _
            handNumber _
        ) + 1


    BJ_PlayerCards( _
        playerNumber, _
        handNumber, _
        BJ_PlayerCardCount( _
            playerNumber, _
            handNumber _
        ) _
    ) = cardText

End Sub


Private Sub BJ_AddCardToDealer(ByVal cardText As String)

    If cardText = "" Then Exit Sub

    If BJ_DealerCount >= BJ_MAX_CARDS Then Exit Sub


    BJ_DealerCount = BJ_DealerCount + 1

    BJ_DealerCards(BJ_DealerCount) = cardText

End Sub


' ============================================================
' FIRST PLAYER TURN
' ============================================================

Private Sub BJ_StartFirstTurn()

    Dim p As Long
    Dim h As Long

    BJ_CurrentPlayer = 0
    BJ_CurrentHand = 0


    For p = 1 To 4

        If BJ_PlayerActive(p) Then

            For h = 1 To BJ_PlayerHandCount(p)

                If BJ_HandStatus(p, h) = "PLAYING" Then

                    BJ_CurrentPlayer = p
                    BJ_CurrentHand = h

                    BJ_UpdateTableState

                    Exit Sub

                End If

            Next h

        End If

    Next p


    BJ_FinishPlayersPhase

End Sub


' ============================================================
' HIT
' ============================================================

Public Sub BJ_Hit()

    Dim cardText As String
    Dim handTotal As Long

    If Not BJ_ValidCurrentTurn() Then Exit Sub


    cardText = BJ_DrawCard()


    If cardText = "" Then

        MsgBox _
            "The shoe is empty. Press NEW SHOE.", _
            vbExclamation

        Exit Sub

    End If


    BJ_AddCardToPlayer _
        BJ_CurrentPlayer, _
        BJ_CurrentHand, _
        cardText


    handTotal = _
        BJ_PlayerTotal( _
            BJ_CurrentPlayer, _
            BJ_CurrentHand _
        )


    If handTotal > 21 Then

        BJ_HandStatus( _
            BJ_CurrentPlayer, _
            BJ_CurrentHand _
        ) = "BUST"

        BJ_RenderPlayer BJ_CurrentPlayer

        BJ_AdvanceTurn


    ElseIf handTotal = 21 Then

        BJ_HandStatus( _
            BJ_CurrentPlayer, _
            BJ_CurrentHand _
        ) = "STAND"

        BJ_RenderPlayer BJ_CurrentPlayer

        BJ_AdvanceTurn


    Else

        BJ_RenderPlayer BJ_CurrentPlayer

        BJ_UpdateTableState

    End If

End Sub


' ============================================================
' STAND
' ============================================================

Public Sub BJ_Stand()

    If Not BJ_ValidCurrentTurn() Then Exit Sub


    BJ_HandStatus( _
        BJ_CurrentPlayer, _
        BJ_CurrentHand _
    ) = "STAND"


    BJ_RenderPlayer BJ_CurrentPlayer

    BJ_AdvanceTurn

End Sub


' ============================================================
' DOUBLE
' ============================================================

Public Sub BJ_Double()

    Dim cardText As String
    Dim handTotal As Long

    If Not BJ_ValidCurrentTurn() Then Exit Sub


    If BJ_PlayerCardCount( _
        BJ_CurrentPlayer, _
        BJ_CurrentHand _
    ) <> 2 Then

        MsgBox _
            "DOUBLE is only available on a two-card hand.", _
            vbInformation

        Exit Sub

    End If


    cardText = BJ_DrawCard()


    If cardText = "" Then

        MsgBox _
            "The shoe is empty. Press NEW SHOE.", _
            vbExclamation

        Exit Sub

    End If


    BJ_Doubled( _
        BJ_CurrentPlayer, _
        BJ_CurrentHand _
    ) = True


    BJ_AddCardToPlayer _
        BJ_CurrentPlayer, _
        BJ_CurrentHand, _
        cardText


    handTotal = _
        BJ_PlayerTotal( _
            BJ_CurrentPlayer, _
            BJ_CurrentHand _
        )


    If handTotal > 21 Then

        BJ_HandStatus( _
            BJ_CurrentPlayer, _
            BJ_CurrentHand _
        ) = "BUST"

    Else

        BJ_HandStatus( _
            BJ_CurrentPlayer, _
            BJ_CurrentHand _
        ) = "STAND"

    End If


    BJ_RenderPlayer BJ_CurrentPlayer

    BJ_AdvanceTurn

End Sub


' ============================================================
' SPLIT
' ============================================================

Public Sub BJ_Split()

    Dim p As Long

    Dim firstCard As String
    Dim secondCard As String

    Dim firstRank As String

    Dim newCardOne As String
    Dim newCardTwo As String

    If Not BJ_ValidCurrentTurn() Then Exit Sub

    p = BJ_CurrentPlayer


    If BJ_CurrentHand <> 1 Then

        MsgBox _
            "This hand cannot be split again.", _
            vbInformation

        Exit Sub

    End If


    If BJ_PlayerHandCount(p) <> 1 Then

        MsgBox _
            "Only one split is supported per player.", _
            vbInformation

        Exit Sub

    End If


    If BJ_PlayerCardCount(p, 1) <> 2 Then

        MsgBox _
            "SPLIT requires exactly two cards.", _
            vbInformation

        Exit Sub

    End If


    firstCard = BJ_PlayerCards(p, 1, 1)
    secondCard = BJ_PlayerCards(p, 1, 2)


    If BJ_CardRank(firstCard) <> _
       BJ_CardRank(secondCard) Then

        MsgBox _
            "The two cards must have the same rank to split.", _
            vbInformation

        Exit Sub

    End If


    If BJ_CardsLeft() < 2 Then

        MsgBox _
            "Not enough cards remain to split.", _
            vbInformation

        Exit Sub

    End If


    firstRank = BJ_CardRank(firstCard)


    ' Hand 1 keeps first card
    BJ_PlayerCards(p, 1, 2) = ""
    BJ_PlayerCardCount(p, 1) = 1


    ' Hand 2 receives second card
    BJ_PlayerHandCount(p) = 2

    BJ_PlayerCards(p, 2, 1) = secondCard
    BJ_PlayerCardCount(p, 2) = 1


    BJ_HandStatus(p, 1) = "PLAYING"
    BJ_HandStatus(p, 2) = "PLAYING"

    BJ_InitialNatural(p, 1) = False
    BJ_InitialNatural(p, 2) = False


    newCardOne = BJ_DrawCard()
    newCardTwo = BJ_DrawCard()


    BJ_AddCardToPlayer p, 1, newCardOne
    BJ_AddCardToPlayer p, 2, newCardTwo


    ' Split aces: exactly one new card each
    If firstRank = "A" Then

        BJ_HandStatus(p, 1) = "STAND"
        BJ_HandStatus(p, 2) = "STAND"

        BJ_RenderPlayer p

        BJ_AdvanceTurn

        Exit Sub

    End If


    ' 21 after split is not natural blackjack
    If BJ_PlayerTotal(p, 1) = 21 Then
        BJ_HandStatus(p, 1) = "STAND"
    End If

    If BJ_PlayerTotal(p, 2) = 21 Then
        BJ_HandStatus(p, 2) = "STAND"
    End If


    BJ_RenderPlayer p


    If BJ_HandStatus(p, 1) = "PLAYING" Then

        BJ_CurrentHand = 1

        BJ_UpdateTableState

    Else

        BJ_AdvanceTurn

    End If

End Sub


' ============================================================
' CURRENT TURN VALID?
' ============================================================

Private Function BJ_ValidCurrentTurn() As Boolean

    BJ_ValidCurrentTurn = False


    If Not BJ_RoundActive Then

        MsgBox _
            "Press DEAL / NEXT ROUND first.", _
            vbInformation

        Exit Function

    End If


    If BJ_CurrentPlayer < 1 Or _
       BJ_CurrentPlayer > 4 Then

        MsgBox _
            "There is no active player turn.", _
            vbInformation

        Exit Function

    End If


    If BJ_HandStatus( _
        BJ_CurrentPlayer, _
        BJ_CurrentHand _
    ) <> "PLAYING" Then

        Exit Function

    End If


    BJ_ValidCurrentTurn = True

End Function


' ============================================================
' ADVANCE PLAYER TURN
' ============================================================

Private Sub BJ_AdvanceTurn()

    Dim p As Long
    Dim h As Long
    Dim startingHand As Long


    For p = BJ_CurrentPlayer To 4

        If BJ_PlayerActive(p) Then

            startingHand = 1

            If p = BJ_CurrentPlayer Then
                startingHand = BJ_CurrentHand + 1
            End If


            For h = startingHand To BJ_PlayerHandCount(p)

                If BJ_HandStatus(p, h) = "PLAYING" Then

                    BJ_CurrentPlayer = p
                    BJ_CurrentHand = h

                    BJ_UpdateTableState

                    Exit Sub

                End If

            Next h

        End If

    Next p


    BJ_CurrentPlayer = 0
    BJ_CurrentHand = 0

    BJ_FinishPlayersPhase

End Sub


' ============================================================
' PLAYER PHASE COMPLETE
' ============================================================

Private Sub BJ_FinishPlayersPhase()

    If BJ_DealerActive Then

        BJ_DealerHidden = False

        BJ_RenderDealer


        If BJ_DealerNeedsToPlay() Then

            BJ_PlayDealer

        Else

            BJ_SettleRound

        End If

    Else

        BJ_CompletePracticeRound

    End If

End Sub


Private Function BJ_DealerNeedsToPlay() As Boolean

    Dim p As Long
    Dim h As Long

    BJ_DealerNeedsToPlay = False


    For p = 1 To 4

        If BJ_PlayerActive(p) Then

            For h = 1 To BJ_PlayerHandCount(p)

                If BJ_HandStatus(p, h) = "STAND" Then

                    BJ_DealerNeedsToPlay = True
                    Exit Function

                End If

            Next h

        End If

    Next p

End Function


' ============================================================
' DEALER
'
' Dealer stands on soft 17.
' ============================================================

Private Sub BJ_PlayDealer()

    Dim dealerTotal As Long
    Dim isSoft As Boolean

    Dim cardText As String

    BJ_SetDealerStatus "DEALER TURN"
    BJ_UpdateTurnDisplay "DEALER"


    Do

        dealerTotal = BJ_DealerTotal(isSoft)

        If dealerTotal >= 17 Then Exit Do


        cardText = BJ_DrawCard()


        If cardText = "" Then

            BJ_SetDealerStatus "SHOE EMPTY"

            MsgBox _
                "The shoe became empty during the dealer hand.", _
                vbExclamation

            Exit Do

        End If


        BJ_AddCardToDealer cardText

        BJ_RenderDealer

    Loop


    BJ_SettleRound

End Sub


' ============================================================
' SETTLE ROUND
' ============================================================

Private Sub BJ_SettleRound()

    Dim ws As Worksheet

    Dim p As Long
    Dim h As Long

    Dim dealerTotal As Long
    Dim dealerSoft As Boolean
    Dim dealerBlackjack As Boolean

    Dim playerTotal As Long

    Dim resultText As String
    Dim handResult As String

    Set ws = ThisWorkbook.Worksheets("Blackjack")


    BJ_DealerHidden = False

    BJ_RenderDealer


    dealerTotal = BJ_DealerTotal(dealerSoft)

    dealerBlackjack = BJ_IsDealerBlackjack()


    If dealerBlackjack Then

        BJ_SetDealerStatus "BLACKJACK"

    ElseIf dealerTotal > 21 Then

        BJ_SetDealerStatus "BUST " & dealerTotal

    Else

        BJ_SetDealerStatus "STAND " & dealerTotal

    End If


    For p = 1 To 4

        If BJ_PlayerActive(p) Then

            resultText = ""


            For h = 1 To BJ_PlayerHandCount(p)

                playerTotal = BJ_PlayerTotal(p, h)


                handResult = _
                    BJ_ResultForHand( _
                        p, _
                        h, _
                        playerTotal, _
                        dealerTotal, _
                        dealerBlackjack _
                    )


                If BJ_PlayerHandCount(p) = 1 Then

                    resultText = handResult

                Else

                    If h = 1 Then

                        resultText = _
                            "H1: " & handResult

                    Else

                        resultText = _
                            resultText & _
                            " | H2: " & handResult

                    End If

                End If

            Next h


            BJ_SetMergedValue _
                ws.Cells( _
                    BJ_PlayerResultRow(p), _
                    2 _
                ), _
                resultText


            BJ_SetPlayerStatus p, resultText

        Else

            BJ_SetPlayerStatus p, "OUT"

        End If

    Next p


    BJ_RoundActive = False
    BJ_RoundComplete = True

    BJ_CurrentPlayer = 0
    BJ_CurrentHand = 0

    BJ_UpdateTurnDisplay "ROUND COMPLETE"

End Sub


Private Function BJ_ResultForHand( _
    ByVal playerNumber As Long, _
    ByVal handNumber As Long, _
    ByVal playerTotal As Long, _
    ByVal dealerTotal As Long, _
    ByVal dealerBlackjack As Boolean _
) As String


    If BJ_HandStatus( _
        playerNumber, _
        handNumber _
    ) = "BUST" Then

        BJ_ResultForHand = _
            "BUST - LOSS (" & playerTotal & ")"

        Exit Function

    End If


    If BJ_InitialNatural( _
        playerNumber, _
        handNumber _
    ) Then

        If dealerBlackjack Then

            BJ_ResultForHand = _
                "PUSH - BLACKJACK"

        Else

            BJ_ResultForHand = _
                "BLACKJACK - WIN"

        End If

        Exit Function

    End If


    If dealerBlackjack Then

        BJ_ResultForHand = _
            "LOSS (" & playerTotal & " vs BJ)"

        Exit Function

    End If


    If dealerTotal > 21 Then

        BJ_ResultForHand = _
            "WIN (" & playerTotal & " vs BUST)"

        Exit Function

    End If


    If playerTotal > dealerTotal Then

        BJ_ResultForHand = _
            "WIN (" & playerTotal & _
            " vs " & dealerTotal & ")"


    ElseIf playerTotal < dealerTotal Then

        BJ_ResultForHand = _
            "LOSS (" & playerTotal & _
            " vs " & dealerTotal & ")"


    Else

        BJ_ResultForHand = _
            "PUSH (" & playerTotal & ")"

    End If

End Function


' ============================================================
' DEALER OUT / PRACTICE
' ============================================================

Private Sub BJ_CompletePracticeRound()

    Dim ws As Worksheet

    Dim p As Long
    Dim h As Long

    Dim resultText As String
    Dim totalValue As Long

    Set ws = ThisWorkbook.Worksheets("Blackjack")


    For p = 1 To 4

        If BJ_PlayerActive(p) Then

            resultText = ""


            For h = 1 To BJ_PlayerHandCount(p)

                totalValue = BJ_PlayerTotal(p, h)


                If BJ_PlayerHandCount(p) = 1 Then

                    If BJ_HandStatus(p, h) = "BUST" Then

                        resultText = _
                            "BUST (" & totalValue & ")"

                    Else

                        resultText = _
                            "COMPLETE (" & totalValue & ")"

                    End If

                Else

                    If h = 1 Then

                        resultText = _
                            "H1: " & totalValue

                    Else

                        resultText = _
                            resultText & _
                            " | H2: " & totalValue

                    End If

                End If

            Next h


            BJ_SetMergedValue _
                ws.Cells( _
                    BJ_PlayerResultRow(p), _
                    2 _
                ), _
                resultText


            BJ_SetPlayerStatus p, "PRACTICE COMPLETE"

        Else

            BJ_SetPlayerStatus p, "OUT"

        End If

    Next p


    BJ_SetDealerStatus "OUT"

    BJ_RoundActive = False
    BJ_RoundComplete = True

    BJ_CurrentPlayer = 0
    BJ_CurrentHand = 0

    BJ_UpdateTurnDisplay "ROUND COMPLETE"

End Sub


' ============================================================
' PLAYER TOTAL
' ============================================================

Private Function BJ_PlayerTotal( _
    ByVal playerNumber As Long, _
    ByVal handNumber As Long, _
    Optional ByRef isSoft As Boolean _
) As Long

    Dim i As Long

    Dim totalValue As Long
    Dim aceCount As Long

    Dim rankText As String

    totalValue = 0
    aceCount = 0


    For i = 1 To _
        BJ_PlayerCardCount( _
            playerNumber, _
            handNumber _
        )


        rankText = _
            BJ_CardRank( _
                BJ_PlayerCards( _
                    playerNumber, _
                    handNumber, _
                    i _
                ) _
            )


        If rankText = "A" Then

            totalValue = totalValue + 11
            aceCount = aceCount + 1

        Else

            totalValue = _
                totalValue + _
                BJ_RankValue(rankText)

        End If

    Next i


    Do While totalValue > 21 And aceCount > 0

        totalValue = totalValue - 10
        aceCount = aceCount - 1

    Loop


    isSoft = (aceCount > 0)

    BJ_PlayerTotal = totalValue

End Function


' ============================================================
' DEALER TOTAL
' ============================================================

Private Function BJ_DealerTotal( _
    Optional ByRef isSoft As Boolean _
) As Long

    Dim i As Long

    Dim totalValue As Long
    Dim aceCount As Long

    Dim rankText As String

    totalValue = 0
    aceCount = 0


    For i = 1 To BJ_DealerCount

        rankText = _
            BJ_CardRank( _
                BJ_DealerCards(i) _
            )


        If rankText = "A" Then

            totalValue = totalValue + 11
            aceCount = aceCount + 1

        Else

            totalValue = _
                totalValue + _
                BJ_RankValue(rankText)

        End If

    Next i


    Do While totalValue > 21 And aceCount > 0

        totalValue = totalValue - 10
        aceCount = aceCount - 1

    Loop


    isSoft = (aceCount > 0)

    BJ_DealerTotal = totalValue

End Function


Private Function BJ_RankValue( _
    ByVal rankText As String _
) As Long

    Select Case rankText

        Case "A"
            BJ_RankValue = 11

        Case "K", "Q", "J", "10"
            BJ_RankValue = 10

        Case Else
            BJ_RankValue = CLng(rankText)

    End Select

End Function


Private Function BJ_IsDealerBlackjack() As Boolean

    Dim softValue As Boolean

    BJ_IsDealerBlackjack = _
        BJ_DealerCount = 2 And _
        BJ_DealerTotal(softValue) = 21

End Function


' ============================================================
' CARD TEXT
' ============================================================

Private Function BJ_CardRank( _
    ByVal cardText As String _
) As String

    Dim parts As Variant

    parts = Split(cardText, "-")

    BJ_CardRank = CStr(parts(0))

End Function


Private Function BJ_CardSuit( _
    ByVal cardText As String _
) As String

    Dim parts As Variant

    parts = Split(cardText, "-")

    BJ_CardSuit = CStr(parts(1))

End Function


Private Function BJ_CardDisplay( _
    ByVal cardText As String _
) As String

    Dim rankText As String
    Dim suitText As String
    Dim suitSymbol As String

    rankText = BJ_CardRank(cardText)
    suitText = BJ_CardSuit(cardText)


    Select Case suitText

        Case "S"
            suitSymbol = ChrW(&H2660)

        Case "H"
            suitSymbol = ChrW(&H2665)

        Case "D"
            suitSymbol = ChrW(&H2666)

        Case "C"
            suitSymbol = ChrW(&H2663)

    End Select


    BJ_CardDisplay = _
        rankText & suitSymbol

End Function


' ============================================================
' RENDER ALL
' ============================================================

Private Sub BJ_RenderAll()

    Dim p As Long

    BJ_RenderDealer

    For p = 1 To 4
        BJ_RenderPlayer p
    Next p

    BJ_UpdateCardsLeft

    BJ_UpdateTableState

End Sub


' ============================================================
' RENDER DEALER
' ============================================================

Private Sub BJ_RenderDealer()

    Dim ws As Worksheet

    Dim i As Long
    Dim totalValue As Long
    Dim softValue As Boolean

    Set ws = ThisWorkbook.Worksheets("Blackjack")

    ws.Range("B5:O5").ClearContents


    For i = 1 To BJ_DealerCount

        If i = 1 And BJ_DealerHidden Then

            BJ_RenderHiddenCard _
                ws.Cells(5, 1 + i)

        Else

            BJ_RenderCard _
                ws.Cells(5, 1 + i), _
                BJ_DealerCards(i)

        End If

    Next i


    BJ_ClearMergedCell ws.Range("R5")


    If BJ_DealerCount = 0 Then

        BJ_SetMergedValue ws.Range("R5"), ""

    ElseIf BJ_DealerHidden Then

        BJ_SetMergedValue ws.Range("R5"), "?"

    Else

        totalValue = BJ_DealerTotal(softValue)

        BJ_SetMergedValue _
            ws.Range("R5"), _
            totalValue

    End If

End Sub


' ============================================================
' RENDER PLAYER
' ============================================================

Private Sub BJ_RenderPlayer(ByVal playerNumber As Long)

    Dim ws As Worksheet

    Dim rowOne As Long
    Dim rowTwo As Long

    Dim i As Long

    Dim totalValue As Long
    Dim softValue As Boolean

    Set ws = ThisWorkbook.Worksheets("Blackjack")

    rowOne = BJ_PlayerHandRow(playerNumber, 1)
    rowTwo = BJ_PlayerHandRow(playerNumber, 2)


    ws.Range( _
        ws.Cells(rowOne, 2), _
        ws.Cells(rowOne, 15) _
    ).ClearContents


    ws.Range( _
        ws.Cells(rowTwo, 2), _
        ws.Cells(rowTwo, 15) _
    ).ClearContents


    BJ_ClearMergedCell ws.Cells(rowOne, 18)
    BJ_ClearMergedCell ws.Cells(rowTwo, 18)


    If Not BJ_PlayerActive(playerNumber) Then

        BJ_SetPlayerStatus playerNumber, "OUT"
        Exit Sub

    End If


    ' Hand 1
    For i = 1 To _
        BJ_PlayerCardCount(playerNumber, 1)

        BJ_RenderCard _
            ws.Cells(rowOne, 1 + i), _
            BJ_PlayerCards( _
                playerNumber, _
                1, _
                i _
            )

    Next i


    totalValue = _
        BJ_PlayerTotal( _
            playerNumber, _
            1, _
            softValue _
        )


    BJ_SetMergedValue _
        ws.Cells(rowOne, 18), _
        totalValue


    ' Hand 2
    If BJ_PlayerHandCount(playerNumber) = 2 Then

        For i = 1 To _
            BJ_PlayerCardCount(playerNumber, 2)

            BJ_RenderCard _
                ws.Cells(rowTwo, 1 + i), _
                BJ_PlayerCards( _
                    playerNumber, _
                    2, _
                    i _
                )

        Next i


        totalValue = _
            BJ_PlayerTotal( _
                playerNumber, _
                2, _
                softValue _
            )


        BJ_SetMergedValue _
            ws.Cells(rowTwo, 18), _
            totalValue

    End If


    BJ_UpdatePlayerStatus playerNumber

End Sub


' ============================================================
' CARD VISUALS
' ============================================================

Private Sub BJ_RenderCard( _
    ByVal targetCell As Range, _
    ByVal cardText As String _
)

    Dim suitText As String

    suitText = BJ_CardSuit(cardText)

    targetCell.Value = BJ_CardDisplay(cardText)

    targetCell.Interior.Color = _
        RGB(252, 251, 248)

    targetCell.Font.Name = _
        "Segoe UI Symbol"

    targetCell.Font.Size = 18
    targetCell.Font.Bold = True


    If suitText = "H" Or _
       suitText = "D" Then

        targetCell.Font.Color = _
            RGB(180, 72, 72)

    Else

        targetCell.Font.Color = _
            RGB(45, 45, 45)

    End If


    targetCell.HorizontalAlignment = xlCenter
    targetCell.VerticalAlignment = xlCenter

End Sub


Private Sub BJ_RenderHiddenCard( _
    ByVal targetCell As Range _
)

    targetCell.Value = "??"

    targetCell.Interior.Color = _
        RGB(90, 98, 94)

    targetCell.Font.Color = _
        RGB(245, 245, 242)

    targetCell.Font.Name = "Arial"
    targetCell.Font.Size = 16
    targetCell.Font.Bold = True

    targetCell.HorizontalAlignment = xlCenter
    targetCell.VerticalAlignment = xlCenter

End Sub


' ============================================================
' UPDATE TABLE
' ============================================================

Private Sub BJ_UpdateTableState()

    Dim p As Long

    BJ_UpdateCardsLeft


    If BJ_CurrentPlayer >= 1 And _
       BJ_CurrentPlayer <= 4 Then


        If BJ_PlayerHandCount( _
            BJ_CurrentPlayer _
        ) = 2 Then

            BJ_UpdateTurnDisplay _
                "PLAYER " & _
                BJ_CurrentPlayer & _
                " - HAND " & _
                BJ_CurrentHand

        Else

            BJ_UpdateTurnDisplay _
                "PLAYER " & _
                BJ_CurrentPlayer

        End If


    ElseIf BJ_RoundActive Then

        BJ_UpdateTurnDisplay "DEALER"

    Else

        BJ_UpdateTurnDisplay "-"

    End If


    For p = 1 To 4

        BJ_UpdatePlayerStatus p

    Next p

End Sub


Private Sub BJ_UpdatePlayerStatus( _
    ByVal playerNumber As Long _
)

    Dim textValue As String

    If Not BJ_PlayerActive(playerNumber) Then

        BJ_SetPlayerStatus _
            playerNumber, _
            "OUT"

        Exit Sub

    End If


    If BJ_CurrentPlayer = playerNumber Then

        If BJ_PlayerHandCount(playerNumber) = 2 Then

            textValue = _
                "YOUR TURN - H" & _
                BJ_CurrentHand

        Else

            textValue = "YOUR TURN"

        End If


    ElseIf BJ_PlayerHandCount(playerNumber) = 2 Then

        textValue = _
            "H1 " & _
            BJ_HandStatus(playerNumber, 1) & _
            " / H2 " & _
            BJ_HandStatus(playerNumber, 2)


    Else

        textValue = _
            BJ_HandStatus(playerNumber, 1)

    End If


    BJ_SetPlayerStatus _
        playerNumber, _
        textValue

End Sub


Private Sub BJ_SetPlayerStatus( _
    ByVal playerNumber As Long, _
    ByVal statusText As String _
)

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Blackjack")


    BJ_SetMergedValue _
        ws.Cells( _
            BJ_PlayerHeaderRow(playerNumber), _
            13 _
        ), _
        statusText

End Sub


Private Sub BJ_SetDealerStatus( _
    ByVal statusText As String _
)

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Blackjack")

    BJ_SetMergedValue _
        ws.Range("M4"), _
        statusText

End Sub


Private Sub BJ_UpdateTurnDisplay( _
    ByVal turnText As String _
)

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Blackjack")

    BJ_SetMergedValue _
        ws.Range("W7"), _
        turnText

End Sub


Private Sub BJ_SetPlayerReadyStatuses()

    Dim p As Long

    For p = 1 To 4

        If BJ_PlayerIsSelectedIn(p) Then

            BJ_SetPlayerStatus _
                p, _
                "READY"

        Else

            BJ_SetPlayerStatus _
                p, _
                "OUT"

        End If

    Next p

End Sub


' ============================================================
' ROW MAPS
' ============================================================

Private Function BJ_PlayerHeaderRow( _
    ByVal playerNumber As Long _
) As Long

    BJ_PlayerHeaderRow = _
        9 + ((playerNumber - 1) * 5)

End Function


Private Function BJ_PlayerHandRow( _
    ByVal playerNumber As Long, _
    ByVal handNumber As Long _
) As Long

    BJ_PlayerHandRow = _
        BJ_PlayerHeaderRow(playerNumber) + handNumber

End Function


Private Function BJ_PlayerResultRow( _
    ByVal playerNumber As Long _
) As Long

    BJ_PlayerResultRow = _
        BJ_PlayerHeaderRow(playerNumber) + 3

End Function


' ============================================================
' EMERGENCY EVENT RESET
' ============================================================

Public Sub BJ_RepairEvents()

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox _
        "Excel events are enabled.", _
        vbInformation

End Sub

