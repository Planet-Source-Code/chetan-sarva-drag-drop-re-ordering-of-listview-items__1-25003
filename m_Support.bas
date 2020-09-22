Attribute VB_Name = "m_Support"
'   Copyright (c) 2001, Chetan Sarva. All rights reserved.
'
'   Redistribution and use in source and binary forms, with or without
'   modification, are permitted provided that the following conditions are
'   met:
'
'   -Redistributions of source code must retain the above copyright notice,
'    this list of conditions and the following disclaimer.
'
'   -Redistributions in binary form must reproduce the above copyright
'    notice, this list of conditions and the following disclaimer in the
'    documentation and/or other materials provided with the distribution.
'
'   -Neither the name of pixelcop.com nor the names of its contributors may
'    be used to endorse or promote products derived from this software
'    without specific prior written permission.
'
'   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
'   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
'   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
'   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE
'   CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
'   EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
'   PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
'   PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
'   LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
'   NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
'   SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

Option Explicit

' For redrawing/refreshing the desktop
Public Const RDW_ALLCHILDREN = &H80
Public Const RDW_ERASE = &H4
Public Const RDW_ERASENOW = &H200
Public Const RDW_FRAME = &H400
Public Const RDW_INTERNALPAINT = &H2
Public Const RDW_INVALIDATE = &H1
Public Const RDW_NOCHILDREN = &H40
Public Const RDW_NOERASE = &H20
Public Const RDW_NOFRAME = &H800
Public Const RDW_NOINTERNALPAINT = &H10
Public Const RDW_UPDATENOW = &H100
Public Const RDW_VALIDATE = &H8

Public Declare Function GetDesktopWindow Lib "User32" () As Long ' Changed the RECT to any
Public Declare Function RedrawWindowAny Lib "User32" Alias "RedrawWindow" _
            (ByVal hWnd As Long, lprcUpdate As Any, _
             ByVal hrgnUpdate As Long, _
             ByVal fuRedraw As Long) As Long
            
            
' Other functions

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Declare Function SendMessage Lib "User32" _
   Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
   lParam As Any) As Long

Public Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "User32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long


' For scrolling the listview
Public Const WM_VSCROLL& = &H115
Public Const SB_LINEDOWN& = 1
Public Const SB_LINEUP& = 0


' Listview functions ( Translated by Brad Martinez: http://www.mvps.org/btmtz/ )

Public Type LVHITTESTINFO   ' was LV_HITTESTINFO
  pt As POINTAPI
  Flags As LVHITTESTINFO_flags
  iItem As Long
#If (WIN32_IE >= &H300) Then
  iSubItem As Long    ' this is was NOT in win95.  valid only for LVM_SUBITEMHITTEST
#End If
End Type

Public Enum LVHITTESTINFO_flags
  LVHT_NOWHERE = &H1   ' in LV client area, but not over item
  LVHT_ONITEMICON = &H2
  LVHT_ONITEMLABEL = &H4
  LVHT_ONITEMSTATEICON = &H8
  LVHT_ONITEM = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)
 
  ' outside the LV's client area
  LVHT_ABOVE = &H8
  LVHT_BELOW = &H10
  LVHT_TORIGHT = &H20
  LVHT_TOLEFT = &H40
End Enum

Public Const LVM_FIRST = &H1000
Public Const LVM_GETITEMPOSITION = (LVM_FIRST + 16)
Public Const LVM_HITTEST = (LVM_FIRST + 18)

Public Sub RefreshDesktop()
    Call RedrawWindowAny(GetDesktopWindow(), 0&, 1&, _
            RDW_INVALIDATE Or RDW_UPDATENOW Or RDW_ALLCHILDREN)
End Sub

Public Sub RefreshWindow(hWnd As Long)

    Call RedrawWindowAny(hWnd, 0&, 1&, _
            RDW_INVALIDATE Or RDW_UPDATENOW Or RDW_ALLCHILDREN)

End Sub

Public Function ListView_GetItemPosition(hwndLV As Long, i As Long, ppt As POINTAPI) As Boolean
  ListView_GetItemPosition = SendMessage(hwndLV, LVM_GETITEMPOSITION, ByVal i, ppt)
End Function

Public Function ListView_HitTest(hwndLV As Long, pinfo As LVHITTESTINFO) As Long
  ListView_HitTest = SendMessage(hwndLV, LVM_HITTEST, 0, pinfo)
End Function

