Option Explicit
Dim xlApp, xlBook

Set xlBook=CreateObject("Excel.Application")
xlApp.DisplayAlerts=False
xlApp.Visible=True
Set xlBook=xlApp.Workbooks.Open("C:\Users\gaurvin\Data.xlsm",0,True)
xlApp.Run "RunMacro" ' Name of Macro Function to run
xlBook.Saveas "C:\Users\gaurvin\Data_Result.xls",-4143
xlBook.Close
xlApp.Quit

Set xlBook=nothing
set xlApp= nothing

Wscript.Quit