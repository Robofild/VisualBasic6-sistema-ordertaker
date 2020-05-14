Attribute VB_Name = "ProcurarArquivoPendente"
Option Explicit


Declare Function SearchTreeForFile Lib "IMAGEHLP.DLL" (ByVal lpRootPath As String, _
ByVal lpInputName As String, ByVal lpOutputName As String) As Long

Public Const MAX_PATH = 260
