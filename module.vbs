' Project:  Coin Cost Basis
' Author:   Alan Hettinger
' Version:  0.6 (Beta) (See changelog at bottom of file)
' Purpose:  Automate cost basis calculations of cryptocurrency

' constants
Public Const TX_DATE = "B"
Public Const BUY_COIN = "C"
Public Const BUY_COST = "D"
Public Const SELL_COIN = "E"
Public Const SELL_RECD = "F"
Public Const TX_STATUS = "G"
Public Const COST_BASIS = "H"
Public Const GAIN_LOSS = "I"
Public Const FIRST_ROW = 5

' globals
Public lots() As Variant
Public sales() As Variant
Public lastRow As Integer

' validate ---------------------------------------------------------------------------------------
Function validate()
  Dim lastDate As Date
  Dim coinCheck As Double
  coinCheck = 0
  lastDate = 0

  ' find last row with data 
  lastRow = Cells.Find(What:="*", After:=Range("A1"), LookAt:=xlPart, LookIn:=xlFormulas, _
            SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row

  ' validate date order
  For row = FIRST_ROW To lastRow
    If Not IsDate(ActiveSheet.Range(TX_DATE & row).Value) _
    And ActiveSheet.Range(TX_DATE & row).Value <> "" Then
      MsgBox("Invalid date in row " & row & ".")
      End
    End If
    If Not IsEmpty(ActiveSheet.Range(TX_DATE & row).Value) Then
      If DateDiff("d", ActiveSheet.Range(TX_DATE & row).Value, lastDate) <= 0 Then
        lastDate = ActiveSheet.Range(TX_DATE & row).Value
      Else
        MsgBox("Date out of order in row " & row & ".")
        End
      End If
    End If
  Next row

  ' validate coin running totals
  For row = FIRST_ROW To lastRow
    If ActiveSheet.Range(BUY_COIN & row).Value > 0 _
    Or ActiveSheet.Range(SELL_COIN & row).Value > 0 Then
      If coinCheck - ActiveSheet.Range(SELL_COIN & row).Value < 0 Then
        MsgBox("There were not enough coin buys to support all of your coin sales. " _
        & "Ensure that you have recorded all of your coin buys.")
        End
      Else
        coinCheck = coinCheck _
                  + ActiveSheet.Range(BUY_COIN & row).Value _
                  - ActiveSheet.Range(SELL_COIN & row).Value
      End If
    End If
  Next row

  For row = FIRST_ROW To lastRow
    If ActiveSheet.Range(BUY_COIN & row).Value > 0 Then
      If IsDate(ActiveSheet.Range(TX_DATE & row).Value) And _
      (IsNumeric(ActiveSheet.Range(BUY_COST & row).Value) Or _
      ActiveSheet.Range(BUY_COST & row).Value = 0) And _
      ActiveSheet.Range(SELL_COIN & row).Value = 0 And _
      ActiveSheet.Range(SELL_RECD & row).Value = 0 Then
        ' data is valid
      Else
        MsgBox("Invalid data in row " & row & ".")
        End
      End If
    End If
    If ActiveSheet.Range(SELL_COIN & row).Value > 0 Then
      If IsDate(ActiveSheet.Range(TX_DATE & row).Value) And _
      (IsNumeric(ActiveSheet.Range(SELL_RECD & row).Value) Or _
      ActiveSheet.Range(SELL_RECD & row).Value = 0) And _
      ActiveSheet.Range(BUY_COIN & row).Value = 0 And _
      ActiveSheet.Range(BUY_COST & row).Value = 0 Then
        ' data is valid
      Else
        MsgBox("Invalid data in row " & row & ".")
        End
      End If
    End If
  Next row
  ' if all data is valid, clear sheet
  ActiveSheet.Range(TX_STATUS & FIRST_ROW & ":" & GAIN_LOSS & lastRow ).Value = ""
End Function

' set global variables ---------------------------------------------------------------------------
Function getLots()
  lot = 0
  For row = FIRST_ROW To lastRow
    If ActiveSheet.Range(BUY_COIN & row).Value > 0 Then
      ReDim Preserve lots(3, lot)
      lots(0, lot) = ActiveSheet.Range(TX_DATE & row).Value
      lots(1, lot) = ActiveSheet.Range(BUY_COIN & row).Value
      lots(2, lot) = ActiveSheet.Range(BUY_COST & row).Value
      lots(3, lot) = row
      lot = lot + 1
    End If
  Next row
End Function

Function getSales()
  sale = 0
  For row = FIRST_ROW To lastRow
    If ActiveSheet.Range(SELL_COIN & row).Value > 0 Then
      ReDim Preserve sales(3, sale)
      sales(0, sale) = ActiveSheet.Range(TX_DATE & row).Value
      sales(1, sale) = ActiveSheet.Range(SELL_COIN & row).Value
      sales(2, sale) = ActiveSheet.Range(SELL_RECD & row).Value
      sales(3, sale) = row
      sale = sale + 1
    End If
  Next row
End Function

' calculate fifo ---------------------------------------------------------------------------------
Function calculateFifo()
  Dim shift As Integer
  Dim lotCount As Integer
  Dim lotCoinRemain As Double
  Dim costBasis As Double
  Dim gainLoss As Double
  Dim termSplit As Boolean
  Dim splitFactor As Double
  Dim totalCoin As Double
  Dim totalCost As Double
  Dim sellCoinRemain As Double
  Dim thisTerm As Date
  Dim nextTerm As Date
  Dim originalDate As Date
  Dim originalCoin As Double
  Dim originalRecd As Double
  Dim percentSold As Double
  Dim sellDate As Date
  Dim sellCoin As Double
  Dim sellRecd As Double
  Dim sellRow As Integer
  Dim lotDate As Date
  Dim lotCoin As Double
  Dim lotRecd As Double
  Dim lotRow As Integer

  shift = 0
  lotCount = 0
  lotCoinRemain = lots(1, 0)

  For sale = 0 To UBound(sales, 2)

    sellDate = sales(0, sale)
    sellCoin = sales(1, sale)
    sellRecd = sales(2, sale)
    sellRow = sales(3, sale)

    termSplit = False
    splitFactor = 0
    totalCoin = 0 ' running total of coins for basis
    totalCost = 0 ' running total of dollar cost for basis
    sellCoinRemain = sales(1, sale)

    For lot = lotCount To UBound(lots, 2)

      lotDate = lots(0, lot)
      lotCoin = lots(1, lot)
      lotCost = lots(2, lot)
      lotRow = lots(3, lot)

      ' if the remaining coin to sell is less than what is in the lot,
      ' calculate and post the cost basis and the gain or loss
      Debug.Print sellCoinRemain
      Debug.Print lotCoinRemain
      If Round(sellCoinRemain, 6) <= Round(lotCoinRemain, 6) Then
        With ActiveSheet
          If sellCoinRemain = lotCoinRemain And lotCount < UBound(lots, 2) Then 
            .Range(TX_STATUS & lotRow).Value = "100% Sold"
            lotCount = lotCount + 1
            lotCoinRemain = lots(1, lotCount)
          ElseIf sellCoinRemain = lotCoinRemain Then
            .Range(TX_STATUS & lotRow).Value = "100% Sold"
          Else
            lotCoinRemain = lotCoinRemain - sellCoinRemain
            percentSold = 1 - (lotCoinRemain / lotCoin)
            .Range(TX_STATUS & lotRow).Value = Round(percentSold*100, 0) & "% Sold"
          End If

          ' calculate and post results
          thisTerm = DateAdd("yyyy", 1, lotDate)
          If DateDiff("d", thisTerm, sellDate) >= 0 And termSplit = False Then
            .Range(TX_STATUS & sellRow + shift).Value = "Long-term"
          ElseIf termSplit = False Then
            .Range(TX_STATUS & sellRow + shift).Value = "Short-term"
          End If

          totalCoin = totalCoin + sellCoinRemain
          totalCost = totalCost + (lotCost * (sellCoinRemain / lotCoin))
          costBasis = sellCoin * (totalCost / totalCoin) * (1 - splitFactor)
          gainLoss = (sellRecd * (1 - splitFactor)) - costBasis
          .Range(COST_BASIS & sellRow + shift).Value = costBasis 
          .Range(GAIN_LOSS & sellRow + shift).Value = gainLoss
          
        End With
        Exit For

      ' if the remaining coin to sell is greater than what is in the lot,
      ' determine if there is a term split, and calculate running totals
      Else

        ' look ahead for a term split, and if a split exists, 
        ' set the split factor (% to allocate to either side of the split),
        ' and calculate and post the first half of the split
        Debug.Print "Look Ahead!"
        thisTerm = DateAdd("yyyy", 1, lotDate)
        nextTerm = DateAdd("yyyy", 1, lots(0, lot + 1))
        If DateDiff("d", thisTerm, sellDate) >= 0 _
        And DateDiff("d", nextTerm, sellDate) < 0 Then

          termSplit = True

          totalCoin = totalCoin + lotCoinRemain
          totalCost = totalCost + (lotCost * (lotCoinRemain / lotCoin))

          ' calculate the split factor
          splitFactor = totalCoin / sellCoin

          ' post the long-term split and continue
          costBasis = sellCoin * (totalCost / totalCoin) * splitFactor ' average price
          gainLoss = (sellRecd * splitFactor) - costBasis

          With ActiveSheet

            originalDate = .Range(TX_DATE & sellRow + shift).Value
            originalCoin = .Range(SELL_COIN & sellRow + shift).Value
            originalRecd = .Range(SELL_RECD & sellRow + shift).Value

            .Range(COST_BASIS & sellRow + shift).Value = costBasis
            .Range(GAIN_LOSS & sellRow + shift).Value = gainLoss

            If Not .Range(TX_DATE & sellRow + shift).Comment Is Nothing Then
              .Range(TX_DATE & sellRow + shift).Comment.Delete
            End If

            .Range(TX_DATE & sellRow + shift).AddComment _
                "This sale was split into two sales (rows " & sellRow + shift _
              & " and " & sellRow + shift + 1 _
              & ") because it included both long-term and short-term cost components." _
              & Chr(10) & "The original amount of coin sold was " & Round(originalCoin, 6) & ", " _
              & "and the original amount received was " & Round(originalRecd, 2) & "."
            .Range(TX_DATE & sellRow + shift).Comment.Shape.TextFrame.AutoSize = True
            .Range(SELL_COIN & sellRow + shift).Value = originalCoin * splitFactor
            .Range(SELL_RECD & sellRow + shift).Value = originalRecd * splitFactor
            .Range(TX_STATUS & sellRow + shift).Value = "Long-term"

            .Range("A" & sellRow + shift + 1).EntireRow.Insert
            shift = shift + 1

            If Not .Range(TX_DATE & sellRow + shift).Comment Is Nothing Then
              .Range(TX_DATE & sellRow + shift).Comment.Delete
            End If

            .Range(TX_DATE & sellRow + shift).Value = originalDate
            .Range(TX_DATE & sellRow + shift).AddComment _
                "This sale was split into two sales (rows " & sellRow + shift - 1 _
              & " and " & sellRow + shift _
              & ") because it included both long-term and short-term cost components." _
              & Chr(10) & "The original amount of coin sold was " & Round(originalCoin, 6) & ", " _
              & "and the original amount received was " & Round(originalRecd, 2) & "."
            .Range(TX_DATE & sellRow + shift).Comment.Shape.TextFrame.AutoSize = True
            .Range(SELL_COIN & sellRow + shift).Value = originalCoin * (1 - splitFactor)
            .Range(SELL_RECD & sellRow + shift).Value = originalRecd * (1 - splitFactor)
            .Range(TX_STATUS & sellRow + shift).Value = "Short-term"
          End With

          totalCoin = 0
          totalCost = 0

        ' if there isn't a term split, add to the running totals 
        ' and continue on to the next lot
        Else 
          totalCoin = totalCoin + lotCoinRemain
          totalCost = totalCost + (lotCost * (lotCoinRemain / lotCoin))
        End If

        ' subtract the lot amount from the remaining coin to be sold,
        ' and set up variables for the next lot, since this lot is completely used up
        sellCoinRemain = sellCoinRemain - lotCoinRemain
        ActiveSheet.Range(TX_STATUS & lotRow).Value = "100% Sold"
        lotCount = lotCount + 1
        lotCoinRemain = lots(1, lotCount)
      End If
    Next lot
  Next sale
End Function

Sub fifo()
  validate
  getLots
  getSales
  calculateFifo
End Sub