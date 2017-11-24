' Project:  Coin Cost Basis
' Author:   Alan Hettinger
' Version:  0.5 (Beta) (See changelog at bottom of file)
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
  Dim termTest1 As Date
  Dim termTest2 As Date
  Dim originalDate As Date
  Dim originalCoin As Double
  Dim originalRecd As Double
  Dim percentSold As Double

  shift = 0
  lotCount = 0
  lotCoinRemain = lots(1, 0)

  For sale = 0 To UBound(sales, 2)
    termSplit = False
    splitFactor = 0
    totalCoin = 0 ' running total of coins for basis
    totalCost = 0 ' running total of dollar cost for basis
    sellCoinRemain = sales(1, sale)
    
    For lot = lotCount To UBound(lots, 2)

      ' if the remaining coin to sell is less than what is in the lot,
      ' calculate and post the cost basis and the gain or loss
      If sellCoinRemain <= lotCoinRemain Then
        With ActiveSheet
          If sellCoinRemain = lotCoinRemain And lotCount < UBound(lots, 2) Then 
            .Range(TX_STATUS & lots(3, lot)).Value = "100% Sold"
            lotCount = lotCount + 1
            lotCoinRemain = lots(1, lotCount)
          ElseIf sellCoinRemain = lotCoinRemain Then
            .Range(TX_STATUS & lots(3, lot)).Value = "100% Sold"
          Else
            lotCoinRemain = lotCoinRemain - sellCoinRemain
            percentSold = 1 - (lotCoinRemain / lots(1, lot))
            .Range(TX_STATUS & lots(3, lot)).Value = Round(percentSold*100, 0) & "% Sold"
          End If

          ' calculate and post results
          termTest1 = DateAdd("yyyy", 1, lots(0, lot))
          If DateDiff("d", termTest1, sales(0, sale)) >= 0 And termSplit = False Then
            .Range(TX_STATUS & sales(3, sale) + shift).Value = "Long-term"
          ElseIf termSplit = False Then
            .Range(TX_STATUS & sales(3, sale) + shift).Value = "Short-term"
          End If

          totalCoin = totalCoin + sellCoinRemain
          totalCost = totalCost + (lots(2, lot) * (sellCoinRemain / lots(1, lot)))
          costBasis = sales(1, sale) * (totalCost / totalCoin) * (1 - splitFactor)
          gainLoss = (sales(2, sale) * (1 - splitFactor)) - costBasis
          .Range(COST_BASIS & sales(3, sale) + shift).Value = costBasis 
          .Range(GAIN_LOSS & sales(3, sale) + shift).Value = gainLoss
          
        End With
        Exit For

      ' if the remaining coin to sell is greater than what is in the lot,
      ' determine if there is a term split, and calculate running totals
      Else

        ' look ahead for a term split, and if a split exists, 
        ' set the split factor (% to allocate to either side of the split),
        ' and calculate and post the first half of the split
        termTest1 = DateAdd("yyyy", 1, lots(0, lot))
        termTest2 = DateAdd("yyyy", 1, lots(0, lot + 1))
        If DateDiff("d", termTest1, sales(0, sale)) >= 0 _
        And DateDiff("d", termTest2, sales(0, sale)) < 0 Then

          termSplit = True

          totalCoin = totalCoin + lotCoinRemain
          totalCost = totalCost + (lots(2, lot) * (lotCoinRemain / lots(1, lot)))

          ' calculate the split factor
          splitFactor = totalCoin / sales(1, sale)

          ' post the long-term split and continue
          costBasis = sales(1, sale) * (totalCost / totalCoin) * splitFactor ' average price
          gainLoss = (sales(2, sale) * splitFactor) - costBasis

          With ActiveSheet

            originalDate = .Range(TX_DATE & sales(3, sale) + shift).Value
            originalCoin = .Range(SELL_COIN & sales(3, sale) + shift).Value
            originalRecd = .Range(SELL_RECD & sales(3, sale) + shift).Value

            .Range(COST_BASIS & sales(3, sale) + shift).Value = costBasis
            .Range(GAIN_LOSS & sales(3, sale) + shift).Value = gainLoss

            If Not .Range(TX_DATE & sales(3, sale) + shift).Comment Is Nothing Then
              .Range(TX_DATE & sales(3, sale) + shift).Comment.Delete
            End If

            .Range(TX_DATE & sales(3, sale) + shift).AddComment _
                "This sale was split into two sales (rows " & sales(3, sale) + shift _
              & " and " & sales(3, sale) + shift + 1 _
              & ") because it included both long-term and short-term cost components." _
              & Chr(10) & "The original amount of coin sold was " & Round(originalCoin, 6) & ", " _
              & "and the original amount received was " & Round(originalRecd, 2) & "."
            .Range(TX_DATE & sales(3, sale) + shift).Comment.Shape.TextFrame.AutoSize = True
            .Range(SELL_COIN & sales(3, sale) + shift).Value = originalCoin * splitFactor
            .Range(SELL_RECD & sales(3, sale) + shift).Value = originalRecd * splitFactor
            .Range(TX_STATUS & sales(3, sale) + shift).Value = "Long-term"

            .Range("A" & sales(3, sale) + shift + 1).EntireRow.Insert
            shift = shift + 1

            If Not .Range(TX_DATE & sales(3, sale) + shift).Comment Is Nothing Then
              .Range(TX_DATE & sales(3, sale) + shift).Comment.Delete
            End If

            .Range(TX_DATE & sales(3, sale) + shift).Value = originalDate
            .Range(TX_DATE & sales(3, sale) + shift).AddComment _
                "This sale was split into two sales (rows " & sales(3, sale) + shift - 1 _
              & " and " & sales(3, sale) + shift _
              & ") because it included both long-term and short-term cost components." _
              & Chr(10) & "The original amount of coin sold was " & Round(originalCoin, 6) & ", " _
              & "and the original amount received was " & Round(originalRecd, 2) & "."
            .Range(TX_DATE & sales(3, sale) + shift).Comment.Shape.TextFrame.AutoSize = True
            .Range(SELL_COIN & sales(3, sale) + shift).Value = originalCoin * (1 - splitFactor)
            .Range(SELL_RECD & sales(3, sale) + shift).Value = originalRecd * (1 - splitFactor)
            .Range(TX_STATUS & sales(3, sale) + shift).Value = "Short-term"
          End With

          totalCoin = 0
          totalCost = 0

        ' if there isn't a term split, add to the running totals 
        ' and continue on to the next lot
        Else 
          totalCoin = totalCoin + lotCoinRemain
          totalCost = totalCost + (lots(2, lot) * (lotCoinRemain / lots(1, lot)))
        End If

        ' subtract the lot amount from the remaining coin to be sold,
        ' and set up variables for the next lot, since this lot is completely used up
        sellCoinRemain = sellCoinRemain - lotCoinRemain
        ActiveSheet.Range(TX_STATUS & lots(3, lot)).Value = "100% Sold"
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