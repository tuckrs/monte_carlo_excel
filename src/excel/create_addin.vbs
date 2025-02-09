Set xl = CreateObject("Excel.Application")
xl.Visible = True
xl.Workbooks.Add

' Add VBA code
Set vbProj = xl.ActiveWorkbook.VBProject
Set vbComp = vbProj.VBComponents.Add(1) ' 1 = vbext_ct_StdModule
vbComp.Name = "MonteCarloModule"

' Read the VBA code from file
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile("monte_carlo.bas", 1)
vbCode = file.ReadAll()
file.Close

' Add the code to the module
vbComp.CodeModule.AddFromString vbCode

' Save as add-in
xl.ActiveWorkbook.SaveAs WScript.Arguments(0), 55 ' 55 = xlOpenXMLAddIn (.xlam)
xl.ActiveWorkbook.Close
xl.Quit
