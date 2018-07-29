
Public Function StockPriceClosing(DateOfClose As String, Ticker As String) As Variant

    Dim BeginPos As Integer
    Dim EndOfStringPos As Integer
    Dim LengthOfClosePrice As Integer
    Dim ApiKey As String
    Dim CommasBegin As Integer
    Dim CommasEnd As Integer
    Dim InitializeBeginPos As Integer
    Dim InitializeEndofStringPos As Integer
    Dim FindString As Integer
    Dim ConvertString As Currency

'The below represent constant values that should not change, but are variables in order to explain their nature
    InitializeBeginPos = 5
    InitializeEndofStringPos = 8
    ApiKey = "" 'Enter your API Key here. Must be obtained (For free) from Quandl.com

'If statement that checks if date entered is a weekend, and breaks function if so. Else, converts the date into the API designated form
    If Weekday(DateOfClose, 1) = 1 Or Weekday(DateOfClose, 1) = 7 Then
        StockPriceClosing = "Markets are not open on weekends, please enter a weekday"
        Exit Function
    Else
        DateOfClose = Format(DateOfClose, "yyyy-mm-dd")
    End If

    TextReturn = Application.WorksheetFunction.WebService("https://www.quandl.com/api/v3/datatables/WIKI/PRICES.json?date=" & DateOfClose & "&ticker=" & Ticker & "&api_key=" & ApiKey)

'For Loop from 1 to 5, which represents number of commas untill close price in string
    For CommasBegin = 1 To InitializeBeginPos
        BeginPos = InStr(BeginPos + 1, TextReturn, ",") + 1
    Next

'For Loop from 1 to 8 in reverse(instrev function to reverse), which represents number of commas untill close price from right to left in string
    EndOfStringPos = InStr(1, TextReturn, "]") 'Initialized the end point of desired return data
    
    For CommasEnd = 1 To InitializeEndofStringPos
        EndOfStringPos = InStrRev(TextReturn, ",", EndOfStringPos - 1)
    Next


'Finsds the numbers of characters at which the string ends
    EndOfStringPosDelete = InStrRev(TextReturn, ",", InStr(1, TextReturn, "]")) - 1

'Finds the length of the exchange rate based off of the begining, and ending variables stated above
    LengthOfClosePrice = EndOfStringPos - BeginPos

'Finds the exchange rate wtihin the given API Json string return, based upon Begining value position, and ending value position
    ConvertString = Mid(TextReturn, BeginPos, LengthOfClosePrice)
    StockPriceClosing = ConvertString


End Function

Sub FunctionDescriptionForStockPriceClosing()
'Running this program will code the StockPriceClosing function to the "Financial" area of functions, and will add descriptions
Dim FuncName As String
Dim FuncDesc As String
Dim FuncCat As Variant

Dim ArgDesc(1 To 2) As String '(the function has 2 arguments)
FuncName = "PERSONAL.XLSB!StockPriceClosing" '(function's name)
FuncDesc = "Close Price of Stock Given a Date, and Ticker" '(function's description)
FuncCat = 1 '(function category)
ArgDesc(1) = "The date of close" '(description of the first argument)
ArgDesc(2) = "The stock ticker"

Application.MacroOptions Macro:=FuncName, Description:=FuncDesc, Category:=FuncCat, ArgumentDescriptions:=ArgDesc()

End Sub
