Attribute VB_Name = "Módulo2"
Sub cordenada()
Attribute cordenada.VB_ProcData.VB_Invoke_Func = "l\n14"
' cordenada Macro
'
    Range("C58").Select
    Selection.Cut
    Range("D57").Select
    ActiveSheet.Paste
    Range("C57").Select
    Selection.Cut
    Range("C58").Select
    ActiveSheet.Paste
    Range("D57").Select
    Selection.Cut
    Range("C57").Select
    ActiveSheet.Paste
End Sub
Sub Extract()

    Dim RarIt As String
    Dim Source As String
    Dim Desti As String
    Dim WinRarPath As String
    ' ubica ruta en directorio de trabajo
    Ruta = "C:\Users\ERICSSON\Downloads\DESCARGAS_MPA"
    ChDir Ruta
    'se posiciona en directorio de trabajo
    trabajo = Dir(Ruta & "\*.rar")
    WinRarPath = "C:\Users\ERICSSON\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\WinRAR"
    Source = Ruta & "\" & trabajo
    Desti = "C:\Users\ERICSSON\Downloads\DESCARGAS_MPA"

    RarIt = Shell(WinRarPath & "WinRar.exe e " & Source & " " & Desti, vbNormalFocus)

End Sub

 Sub Extract2()
 Dim strFilePath As String
    Ruta = "C:\Users\ERICSSON\Downloads\DESCARGAS_MPA"
    ChDir Ruta
    trabajo = Dir(Ruta & "\*.rar")
 strFilePath = Ruta & "\" & trabajo
 Unzip (strFilePath)
 End Sub


