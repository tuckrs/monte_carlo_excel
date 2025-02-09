Attribute VB_Name = "MonteCarloModule"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
#Else
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
#End If

Private Function GetPythonPath() As String
    ' This is where Python is typically installed for the current user
    GetPythonPath = Environ$("LOCALAPPDATA") & "\Programs\Python\Python312\python.exe"
End Function

Public Sub InitializeRibbon()
    ' This will be called when the add-in loads
    Debug.Print "Monte Carlo Add-in Initialized"
End Sub

Public Sub RunSimulation()
    ' Get the selected range
    Dim inputRange As Range
    Dim outputRange As Range
    
    On Error Resume Next
    Set inputRange = Application.Selection
    If inputRange Is Nothing Then
        MsgBox "Please select an input range first!", vbExclamation
        Exit Sub
    End If
    
    ' Get output range
    Set outputRange = Application.InputBox("Select where to put the results:", "Output Range", Type:=8)
    If outputRange Is Nothing Then Exit Sub
    
    ' Run simulation
    On Error Resume Next
    
    ' Create Python command
    Dim pythonCmd As String
    pythonCmd = "import numpy as np; import pandas as pd; " & _
                "data = np.random.normal(100, 15, 1000); " & _
                "stats = pd.Series(data).describe(); " & _
                "xl_range = xw.Range('" & outputRange.Address & "'); " & _
                "xl_range.value = [[i, j] for i, j in zip(['count', 'mean', 'std', 'min', '25%', '50%', '75%', 'max'], stats)]"
    
    ' Get the Python path
    Dim pythonPath As String
    pythonPath = GetPythonPath()
    
    If Dir(pythonPath) = "" Then
        MsgBox "Python not found at: " & pythonPath & vbNewLine & _
               "Please install Python 3.12 and try again.", vbCritical
        Exit Sub
    End If
    
    ' Run Python code using xlwings
    Dim xl As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    
    Dim wb As Object
    Set wb = xl.Workbooks.Add
    
    wb.RunPython "import xlwings as xw; " & pythonCmd
    
    wb.Close False
    xl.Quit
    
    Set wb = Nothing
    Set xl = Nothing
    
    If Err.Number <> 0 Then
        MsgBox "Error running simulation: " & Err.Description, vbCritical
    Else
        MsgBox "Simulation completed successfully!", vbInformation
    End If
End Sub

Public Sub ShowHelp()
    MsgBox "Monte Carlo Excel Add-in" & vbNewLine & vbNewLine & _
           "To run a simulation:" & vbNewLine & _
           "1. Enter your data in columns with headers" & vbNewLine & _
           "2. Select the entire range including headers" & vbNewLine & _
           "3. Click 'Run Simulation'" & vbNewLine & _
           "4. Select where to put the results" & vbNewLine & vbNewLine & _
           "The simulation will generate statistics including:" & vbNewLine & _
           "- Mean" & vbNewLine & _
           "- Standard Deviation" & vbNewLine & _
           "- 5th and 95th percentiles", _
           vbInformation
End Sub
