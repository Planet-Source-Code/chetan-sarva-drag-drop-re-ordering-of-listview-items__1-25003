VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ListView Drag&Drop Example by Chetan Sarva"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picL_mask 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5040
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   3660
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picL 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4740
      Picture         =   "frmMain.frx":005C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   3660
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.VScrollBar vsOrder 
      Height          =   7455
      LargeChange     =   3
      Left            =   4380
      TabIndex        =   2
      Top             =   60
      Width           =   255
   End
   Begin MSComctlLib.ListView lvDragDrop 
      Height          =   7455
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   13150
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Test Items"
         Object.Width           =   2734
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Test Data"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblDesc 
      Caption         =   "This example shows us how to do 2 things. "
      Height          =   3495
      Left            =   4740
      TabIndex        =   1
      Top             =   60
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private cDib As New cDIBSection
Private cdibMask As New cDIBSection
Private LastY As Long

'The item being dragged
Private DragLV As ListItem

Private Sub Form_Load()

    lblDesc.Caption = "This example shows us how to do 2 things:" & _
        vbCrLf & vbCrLf & _
        "1) Re-order items in a listview by using drag & drop" & vbCrLf & _
        "2) Give a visual cue to the user for the placement of the item" & vbCrLf & _
        "3) also reposition the items using a scrollbar"
        
    ' Add some items that we can play with to the listview
    Dim i As Long
    Dim li As ListItem
    For i = 1 To 20
        Set li = lvDragDrop.ListItems.Add(, , "Random item " & i)
            li.SubItems(1) = String(5, 64 + i)
    Next i
    
    vsOrder.Min = 1
    vsOrder.Max = 20
    
    ' Load the images we will use for the drag notifier
    cDib.CreateFromPicture picL.Picture
    cdibMask.CreateFromPicture picL_mask.Picture
        
End Sub

Private Sub lvDragDrop_ItemClick(ByVal Item As MSComctlLib.ListItem)

    ' Update the value of the scroll bar
    vsOrder.Value = Item.Index

End Sub

Private Sub lvDragDrop_OLECompleteDrag(Effect As Long)

    ' Cleanup the window
    RefreshWindow Me.hWnd
    
End Sub

Private Sub lvDragDrop_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   
    Dim liNew As ListItem
    Dim pinfo As LVHITTESTINFO
    Dim pt As POINTAPI
    Dim pti As POINTAPI
    
    Dim hitItem As ListItem
    Set hitItem = lvDragDrop.HitTest(x, y)
    
    ' Move the item to its new position
    
    ' We dropped on an item, so move the dragged item above this one.
    If Not hitItem Is Nothing Then
        If hitItem.Index <> DragLV.Index Then
            Set liNew = lvDragDrop.ListItems.Add(hitItem.Index, , DragLV.Text)
                liNew.SubItems(1) = DragLV.SubItems(1)
                liNew.Selected = True
            lvDragDrop.ListItems.Remove DragLV.Index
            vsOrder.Value = liNew.Index
        End If
    
    ' We're not over an item but we may be below the last item
    Else
        
        GetCursorPos pt
        Call ListView_GetItemPosition(lvDragDrop.hWnd, lvDragDrop.ListItems.Item(lvDragDrop.ListItems.Count - 1).Index, pti)
        If pt.y > Me.Top / Screen.TwipsPerPixelY + pti.y Then
            Set liNew = lvDragDrop.ListItems.Add(, , DragLV.Text)
                liNew.SubItems(1) = DragLV.SubItems(1)
                liNew.Selected = True
            lvDragDrop.ListItems.Remove DragLV.Index
            vsOrder.Value = liNew.Index
        End If
        
    End If

End Sub

Private Sub lvDragDrop_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    ' If we are at the top or bottom of the listview with our cursor
    ' scroll the items if need be.
    With lvDragDrop
        If y <= 100 Then
            SendMessage .hWnd, WM_VSCROLL, SB_LINEUP, .hWnd
        ElseIf y >= .Height - 100 Then
            SendMessage .hWnd, WM_VSCROLL, SB_LINEDOWN, .hWnd
        End If
    End With
    
    Dim pti As POINTAPI
    Dim pt As POINTAPI
    Dim pinfo As LVHITTESTINFO
    Dim hScreenDC As Long
    
    Dim hitItem As ListItem
    Set hitItem = lvDragDrop.HitTest(x, y)
    
    GetCursorPos pt
    
    If Not hitItem Is Nothing Then
        ' We need to subtract 1 because in the API the listview starts at item 0
        Call ListView_GetItemPosition(lvDragDrop.hWnd, hitItem.Index - 1, pti)

        ' Why add 22? It centers the arrow on the gridline perfectly
        ' (not sure if it will be thrown off by different resolution screens
        ' but it can be easily fixed by playing with the number...)
        If LastY <> Me.Top / Screen.TwipsPerPixelY + pti.y + 22 Then
            LastY = Me.Top / Screen.TwipsPerPixelY + pti.y + 22
            
            ' Get the DC of the desktop window
            ' We need to draw right to the desktop because we want the image
            ' to appear over controls as well (in this case the listview).
            ' If we use the DC of the form the image will appear beneath the LV.
            hScreenDC = GetDC(0)
            
            RefreshWindow Me.hWnd
            
            ' First we paint the mask
            cdibMask.PaintPicture hScreenDC, Me.Left / Screen.TwipsPerPixelX + pti.x, LastY, , , , , vbSrcAnd
            ' Then we paint the real picture. This allows us to get a transparent image
            cDib.PaintPicture hScreenDC, Me.Left / Screen.TwipsPerPixelX + pti.x, LastY, , , , , vbSrcPaint
            
            ' Release the DC for the desktop
            ReleaseDC 0, hScreenDC
        End If
        
    Else
        
        ' See if we're dragging below the last item
        Call ListView_GetItemPosition(lvDragDrop.hWnd, lvDragDrop.ListItems.Item(lvDragDrop.ListItems.Count - 1).Index, pti)
        If pt.y > Me.Top / Screen.TwipsPerPixelY + pti.y Then
        
            ' We are, so draw the notifier below the very last item
            
            If LastY <> Me.Top / Screen.TwipsPerPixelY + pti.y + 39 Then
                LastY = Me.Top / Screen.TwipsPerPixelY + pti.y + 39
                
                hScreenDC = GetDC(0)
                
                RefreshWindow Me.hWnd
    
                cdibMask.PaintPicture hScreenDC, Me.Left / Screen.TwipsPerPixelX + pti.x, LastY, , , , , vbSrcAnd
                cDib.PaintPicture hScreenDC, Me.Left / Screen.TwipsPerPixelX + pti.x, LastY, , , , , vbSrcPaint
                
                ReleaseDC 0, hScreenDC
            End If
        End If
        
    End If

End Sub

Private Sub lvDragDrop_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)

    'The Item to drag
    Set DragLV = lvDragDrop.SelectedItem
    
End Sub

Private Sub vsOrder_Change()

    Dim liMoving As ListItem
    Dim liNew As ListItem
    Dim newPos As Integer

    Set liMoving = lvDragDrop.SelectedItem
    If liMoving.Index = vsOrder.Value Then Exit Sub ' No change, exit
    
    ' See if we are moving up or down
    If vsOrder.Value > liMoving.Index Then
        ' We need to add one to get to the right position
        ' under the current item
        newPos = vsOrder.Value + 1
    Else
        newPos = vsOrder.Value
    End If
    
    ' Add the item in its new position
    Set liNew = lvDragDrop.ListItems.Add(newPos, , liMoving.Text)
        liNew.SubItems(1) = liMoving.SubItems(1)
        liNew.Selected = True
    
    ' Remove the item from its old position
    lvDragDrop.ListItems.Remove (liMoving.Index)
    
End Sub

Private Sub vsOrder_Scroll()

    ' Same as vsOrder_Change
    
    Dim liMoving As ListItem
    Dim liNew As ListItem
    Dim newPos As Integer

    Set liMoving = lvDragDrop.SelectedItem
    If liMoving.Index = vsOrder.Value Then Exit Sub ' No change, exit
    
    ' See if we are moving up or down
    If vsOrder.Value > liMoving.Index Then
        ' We need to add one to get to the right position
        ' under the current item
        newPos = vsOrder.Value + 1
    Else
        newPos = vsOrder.Value
    End If
    
    ' Add the item in its new position
    Set liNew = lvDragDrop.ListItems.Add(newPos, , liMoving.Text)
        liNew.SubItems(1) = liMoving.SubItems(1)
        liNew.Selected = True
    
    ' Remove the item from its old position
    lvDragDrop.ListItems.Remove (liMoving.Index)
End Sub
