Attribute VB_Name = "modActionKeystate"
' Declares
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Sub keybd_event Lib "user32" (ByVal bVK As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
   
' Function to get or set keyboard states
Public Function GetSetKS(ByVal bolGetSetKS As Boolean, _
                                 Optional ByVal intGetSetKSNumLock As Integer = 0, _
                                 Optional ByVal intGetSetKSScrollLock As Integer = 0, _
                                 Optional ByVal intGetSetKSCapsLock As Integer = 0) As String

   ' Check if get or set keystate
   If bolGetSetKS = True Then ' Get keystate
      GetSetKS = IIf(CBool(GetKeyState(vbKeyNumlock) And 1) = True, 1, 0) & "," & _
                         IIf(CBool(GetKeyState(vbKeyScrollLock) And 1) = True, 1, 0) & "," & _
                         IIf(CBool(GetKeyState(vbKeyCapital) And 1) = True, 1, 0)
   Else ' Set keystate
      ' Num lock
      Select Case intGetSetKSNumLock
         Case 1 ' On
            If CBool(GetKeyState(vbKeyNumlock) And 1) = False Then
               Call keybd_event(&H90, &H45, &H1 Or 0, 0)
               Call keybd_event(&H90, &H45, &H1 Or &H2, 0)
            End If
         Case 2 ' Off
            If CBool(GetKeyState(vbKeyNumlock) And 1) = True Then
               Call keybd_event(&H90, &H45, &H1 Or 0, 0)
               Call keybd_event(&H90, &H45, &H1 Or &H2, 0)
            End If
         Case 3 ' Toggle
            Call keybd_event(&H90, &H45, &H1 Or 0, 0)
            Call keybd_event(&H90, &H45, &H1 Or &H2, 0)
      End Select
      ' Scroll lock
      Select Case intGetSetKSScrollLock
         Case 1 ' On
            If CBool(GetKeyState(vbKeyScrollLock) And 1) = False Then
               Call keybd_event(&H91, &H45, &H1 Or 0, 0)
               Call keybd_event(&H91, &H45, &H1 Or &H2, 0)
            End If
         Case 2 ' Off
            If CBool(GetKeyState(vbKeyScrollLock) And 1) = True Then
               Call keybd_event(&H91, &H45, &H1 Or 0, 0)
               Call keybd_event(&H91, &H45, &H1 Or &H2, 0)
            End If
         Case 3 ' Toggle
            Call keybd_event(&H91, &H45, &H1 Or 0, 0)
            Call keybd_event(&H91, &H45, &H1 Or &H2, 0)
      End Select
      ' Caps lock
      Select Case intGetSetKSCapsLock
         Case 1 ' On
            If CBool(GetKeyState(vbKeyCapital) And 1) = False Then
               Call keybd_event(&H14, &H45, &H1 Or 0, 0)
               Call keybd_event(&H14, &H45, &H1 Or &H2, 0)
            End If
         Case 2 ' Off
            If CBool(GetKeyState(vbKeyCapital) And 1) = True Then
               Call keybd_event(&H14, &H45, &H1 Or 0, 0)
               Call keybd_event(&H14, &H45, &H1 Or &H2, 0)
            End If
         Case 3 ' Toggle
            Call keybd_event(&H14, &H45, &H1 Or 0, 0)
            Call keybd_event(&H14, &H45, &H1 Or &H2, 0)
      End Select
      ' Let windows have resources
      DoEvents
   End If
End Function

