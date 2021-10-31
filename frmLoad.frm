VERSION 5.00
Begin VB.Form frmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Associate File Type and Icon with your program"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFile 
      Height          =   4695
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function fCreateShellLink Lib "VB6STKIT.DLL" (ByVal _
        lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal _
        lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long

'NOTE: If you get any GPFs, use this one and not the one above:
'Or if you are using VB5 or earlier, use this instead:
'Private Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal _
 '    lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal _
 '    lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long
 
'To update windows Icon Cache
Private Declare Sub SHChangeNotify Lib "shell32" (ByVal wEventId As Long, _
                        ByVal uFlags As Long, ByVal dwItem1 As Long, _
                        ByVal dwItem2 As Long)

' A file type association has changed.
Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0

Private Sub Form_Load()

    Dim strString As String
    Dim lngDword As Long
    Dim Record As String


    If Command$ <> "%1" And Command$ <> "" Then
        'Command$ is the file you need To open!

        'Load the file
        Open Command$ For Input As #1
        Do While Not EOF(1)
            Line Input #1, Record
            txtFile = txtFile & Record & vbCrLf
        Loop

        'Add your file to the Recent file folder:
        lReturn = fCreateShellLink("..\..\Recent", _
                Command$, Command$, "")

    End If


    'See if our file extension already exists:
    If GetString(HKEY_CLASSES_ROOT, ".xyz", "Content Type") = "" Then
        'Nope - not added yet. Register the file type:
        
        'create an entry in the class key
        Call SaveString(HKEY_CLASSES_ROOT, ".xyz", "", "xyzfile")
        'content type
        Call SaveString(HKEY_CLASSES_ROOT, ".xyz", "Content Type", "text/plain")
        'name
        Call SaveString(HKEY_CLASSES_ROOT, "xyzfile", "", "This is where you type the description for your files")
        'edit flags
        Call SaveDWord(HKEY_CLASSES_ROOT, "xyzfile", "EditFlags", "0000")
        'file's icon (can be an icon file, or an icon located within a dll file)
        'in this example, I am using a resource icon in this exe, 0 (app icon).
        Call SaveString(HKEY_CLASSES_ROOT, "xyzfile\DefaultIcon", "", App.Path & "\" & App.EXEName & ".exe,0")
        'Shell
        Call SaveString(HKEY_CLASSES_ROOT, "xyzfile\Shell", "", "")
        'Shell Open
        Call SaveString(HKEY_CLASSES_ROOT, "xyzfile\Shell\Open", "", "")
        'Shell open command
        Call SaveString(HKEY_CLASSES_ROOT, "xyzfile\Shell\Open\Command", "", App.Path & "\" & App.EXEName & ".exe %1")
        'Update the Windows Icon Cache to see our icon right away:
        SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0

    End If
    

End Sub

