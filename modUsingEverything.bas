Attribute VB_Name = "modUsingEverything"
Option Explicit

Private Everything As String
Private Const VB6Compat As String = "Microsoft.VisualBasic.Compatibility.VB6"


Public Function UsingEverything(Optional ByVal PackageName As String) As String
  Dim List As String, Path As String, Name As String
  Dim E As String, L
  Dim R As String, N As String, M As String
  E = ""
  R = "": N = vbCrLf: M = ""
  
  If PackageName <> "" Then
'    R = R & N & "package " & PackagePrefix & PackageName & ";"
    R = R & N & ""
  End If
  
  If Everything = "" Then
    E = E & M & "using VB6 = " & VB6Compat & ";"
    E = E & N & "using static VBExtension;"
    
    E = E & N & "using static System.DateTime;"
    E = E & N & "using static System.Math;"
    
    E = E & N & "using static Microsoft.VisualBasic.Information;"
    E = E & N & "using static Microsoft.VisualBasic.Conversion;"
    E = E & N & "using static Microsoft.VisualBasic.Strings;"
    E = E & N & "using static Microsoft.VisualBasic.VBMath;"
    
    E = E & N & "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.ColorConstants;"
    E = E & N & "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.DrawStyleConstants;"
    E = E & N & "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.FillStyleConstants;"
    E = E & N & "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.GlobalModule;"
    E = E & N & "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.Printer;"
    E = E & N & "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.PrinterCollection;"
    E = E & N & "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.PrinterObjectConstants;"
    E = E & N & "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.ScaleModeConstants;"
    E = E & N & "using static Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6.SystemColorConstants;"
    E = E & N & "using ADODB;"
    
    Path = FilePath(vbpFile)
    For Each L In Split(VBPModules(vbpFile) & vbCrLf & VBPForms(vbpFile), vbCrLf)
      If L <> "" Then
        Name = ModuleName(ReadEntireFile(Path & L))
        E = E & N & "using static " & PackagePrefix & Name & ";"
      End If
    Next
'    For Each L In Split(VBPClasses(vbpFile), vbCrLf)  ' controls?
'      If L <> "" Then
'        Name = ModuleName(ReadEntireFile(Path & L))
'        E = E & N & "using " & PackagePrefix & Name & ";"
'      End If
'    Next
    Everything = E
  End If
  
  R = Everything & N & R
  UsingEverything = R
End Function

