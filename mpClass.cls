VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mpClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpszCommand As String, ByVal lpszReturnString As String, ByVal cchReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function IsWindowVisible Lib "user32.dll" (ByVal hWnd As Long) As Long
Private mFile As String
Private k1 As RECT

Public Property Get mHeight() As Variant
mGetPos
        mHeight = k1.Bottom - k1.Top
End Property

Public Property Get mWidth() As Variant
mGetPos
        mWidth = k1.Right - k1.Left
End Property

Public Property Get mLeft() As Variant
mGetPos
        mLeft = k1.Left
End Property

Public Property Get mTop() As Variant
mGetPos
        mTop = k1.Top
End Property

Public Sub mLoadFile(FileName As String)
mciSendString "open " & FileName & " alias my1 style popup", "", 0, 0 'Open the file with an Alias, so we can call it by name
mFile = FileName
End Sub

Public Sub mPlay()
mciSendString "play my1", "", 0, 0 'Play it
End Sub

Public Sub mStop()
mciSendString "stop my1", "", 0, 0 'Stop it
End Sub

Public Sub mUnload()
mciSendString "close my1", "", 0, 0 'Close it. Remember to put this into your "form_unload" event!!
End Sub

Public Sub mPause()
mciSendString "pause my1", "", 0, 0 'Pause it
End Sub

Public Sub mSetPos(Top As Long, Left As Long, Width As Long, Height As Long)
Dim playwin As Long
playwin = FindWindow(vbNullString, Right(mFile, Len(mFile) - InStrRev(mFile, "\"))) 'The window's caption is the filename, so we need to find a win with this caption.
SetWindowPos playwin, 0&, Left, Top, Width, Height 'now set its pos
End Sub

Private Sub mGetPos()
Dim playwin As Long
playwin = FindWindow(vbNullString, Right(mFile, Len(mFile) - InStrRev(mFile, "\"))) ' Find that window, again.
GetWindowRect playwin, k1 'Get its position
End Sub

Public Property Get StillPlaying() As Boolean
Dim playwin As Long
playwin = FindWindow(vbNullString, Right(mFile, Len(mFile) - InStrRev(mFile, "\"))) ' Find that window, again.

If IsWindowVisible(playwin) = 1 Then 'if we can see the win, it's working
StillPlaying = True
Else
StillPlaying = False 'if not....
End If
End Property
