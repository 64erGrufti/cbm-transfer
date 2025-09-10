VERSION 5.00
Begin VB.Form frmMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   195
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Menu mnuL 
      Caption         =   "Local"
      Begin VB.Menu mnuLocal 
         Caption         =   "Explore &LEFT Directory"
         Index           =   1
      End
      Begin VB.Menu mnuLocal 
         Caption         =   "Explore &RIGHT Directory"
         Index           =   2
      End
      Begin VB.Menu mnuLocal 
         Caption         =   "S&wap LEFT and RIGHT Paths"
         Index           =   3
      End
      Begin VB.Menu mnuLocal 
         Caption         =   "S&et LEFT Path same as RIGHT"
         Index           =   4
      End
      Begin VB.Menu mnuLocal 
         Caption         =   "Set RIGHT Path same as LEFT"
         Index           =   5
      End
      Begin VB.Menu mnuLocal 
         Caption         =   "Add Current Path to History"
         Index           =   6
      End
      Begin VB.Menu mnuLocal 
         Caption         =   "Remove Current Path from History"
         Index           =   7
      End
      Begin VB.Menu mnuLocal 
         Caption         =   "&Clear History"
         Index           =   8
      End
      Begin VB.Menu mnuLocal 
         Caption         =   "Create &New Folder"
         Index           =   9
      End
   End
   Begin VB.Menu mnuI 
      Caption         =   "Image"
      Begin VB.Menu mnuImg 
         Caption         =   "Save as Text"
         Index           =   1
      End
      Begin VB.Menu mnuImg 
         Caption         =   "Add to catalog"
         Index           =   2
      End
      Begin VB.Menu mnuImg 
         Caption         =   "Show Catalog"
         Index           =   3
      End
      Begin VB.Menu mnuImg 
         Caption         =   "Validate Image"
         Index           =   4
      End
      Begin VB.Menu mnuImg 
         Caption         =   "Backup Image"
         Index           =   5
      End
      Begin VB.Menu mnuImg 
         Caption         =   "Sort by Name A-Z"
         Index           =   6
      End
      Begin VB.Menu mnuImg 
         Caption         =   "Sort by Name Z-A"
         Index           =   7
      End
      Begin VB.Menu mnuImg 
         Caption         =   "Sort by Size 0-9"
         Index           =   8
      End
      Begin VB.Menu mnuImg 
         Caption         =   "Sort by Size 9-0"
         Index           =   9
      End
   End
   Begin VB.Menu mnuF 
      Caption         =   "Font"
      Begin VB.Menu mnuFont 
         Caption         =   "Edit Mode"
         Index           =   1
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Screen Designer"
         Index           =   2
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Multi-Colour"
         Index           =   3
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Border"
         Index           =   4
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Outline"
         Index           =   5
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Hi-light Select Box"
         Index           =   6
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Save as Bitmap..."
         Index           =   7
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Save Font As..."
         Index           =   8
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Save Range As..."
         Index           =   9
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Convert"
         Index           =   10
         Begin VB.Menu mnuConvert 
            Caption         =   "Convert 8x16 to 8x8 font"
            Index           =   0
         End
         Begin VB.Menu mnuConvert 
            Caption         =   "Convert 8x8 to 8x16 font"
            Index           =   1
         End
         Begin VB.Menu mnuConvert 
            Caption         =   "Convert 5x7 sideways font"
            Index           =   2
         End
         Begin VB.Menu mnuConvert 
            Caption         =   "Convert 5x7 upright font"
            Index           =   3
         End
         Begin VB.Menu mnuConvert 
            Caption         =   "Convert 8x14 (EGA ) font"
            Index           =   4
         End
         Begin VB.Menu mnuConvert 
            Caption         =   "Convert 8x32 to  8x16 font"
            Index           =   5
         End
         Begin VB.Menu mnuConvert 
            Caption         =   "Convert 6x8 sideways font"
            Index           =   6
         End
         Begin VB.Menu mnuConvert 
            Caption         =   "Convert 128 CHR font to 256 CHR with RVS"
            Index           =   7
         End
         Begin VB.Menu mnuConvert 
            Caption         =   "Convert Galaksija"
            Index           =   8
         End
      End
   End
   Begin VB.Menu mnuT 
      Caption         =   "Theme"
      Begin VB.Menu mnuTheme 
         Caption         =   "load"
         Index           =   0
      End
   End
   Begin VB.Menu mnuE 
      Caption         =   "Encoding"
      Begin VB.Menu mnuEnc 
         Caption         =   "PETSCII Upper"
         Index           =   1
      End
      Begin VB.Menu mnuEnc 
         Caption         =   "PETSCII Lower"
         Index           =   2
      End
      Begin VB.Menu mnuEnc 
         Caption         =   "Screen Upper"
         Index           =   3
      End
      Begin VB.Menu mnuEnc 
         Caption         =   "Screen Lower"
         Index           =   4
      End
      Begin VB.Menu mnuEnc 
         Caption         =   "ASCII Upper"
         Index           =   5
      End
      Begin VB.Menu mnuEnc 
         Caption         =   "ASCII Lower"
         Index           =   6
      End
   End
   Begin VB.Menu mnuD 
      Caption         =   "Device"
      Begin VB.Menu mnuDev 
         Caption         =   "Initialize"
         Index           =   1
      End
      Begin VB.Menu mnuDev 
         Caption         =   "Validate"
         Index           =   2
      End
      Begin VB.Menu mnuDev 
         Caption         =   "Format Disk"
         Index           =   3
      End
      Begin VB.Menu mnuDev 
         Caption         =   "Change Device#"
         Index           =   4
      End
      Begin VB.Menu mnuDev 
         Caption         =   "Set  Single-Sided Mode"
         Index           =   5
      End
      Begin VB.Menu mnuDev 
         Caption         =   "Set Double-Sided Mode"
         Index           =   6
      End
      Begin VB.Menu mnuDev 
         Caption         =   "Re-Scan Devices"
         Index           =   7
      End
   End
   Begin VB.Menu mnuFontSize 
      Caption         =   "Font Size"
      Begin VB.Menu mnuFS 
         Caption         =   "80 Col 1x2"
         Index           =   1
      End
      Begin VB.Menu mnuFS 
         Caption         =   "40 Col 2x2"
         Index           =   2
      End
      Begin VB.Menu mnuFS 
         Caption         =   "20 Col 4x2"
         Index           =   3
      End
      Begin VB.Menu mnuFS 
         Caption         =   "40 Col Small 1x1"
         Index           =   4
      End
      Begin VB.Menu mnuFS 
         Caption         =   "20 Col Small 2x1"
         Index           =   5
      End
      Begin VB.Menu mnuFS 
         Caption         =   "80 Col Tall 2x3"
         Index           =   6
      End
      Begin VB.Menu mnuFS 
         Caption         =   "40 Col Tall 4x3"
         Index           =   7
      End
   End
   Begin VB.Menu mnuScrEd 
      Caption         =   "ScrEdit"
      Begin VB.Menu mnuSE 
         Caption         =   "Clear Screen"
         Index           =   1
      End
      Begin VB.Menu mnuSE 
         Caption         =   "Reset Machine"
         Index           =   2
      End
      Begin VB.Menu mnuSE 
         Caption         =   "Load Buffer"
         Index           =   3
      End
      Begin VB.Menu mnuSE 
         Caption         =   "Save Buffer"
         Index           =   4
      End
      Begin VB.Menu mnuSE 
         Caption         =   "Export"
         Index           =   5
      End
      Begin VB.Menu mnuSE 
         Caption         =   "Load Macro"
         Index           =   6
      End
      Begin VB.Menu mnuSE 
         Caption         =   "Save Macro"
         Index           =   7
      End
      Begin VB.Menu mnuSE 
         Caption         =   "Toggle Record"
         Index           =   8
      End
      Begin VB.Menu mnuSE 
         Caption         =   "Save Bitmap"
         Index           =   9
      End
   End
   Begin VB.Menu mnuHex 
      Caption         =   "Hex"
      Begin VB.Menu mnuH 
         Caption         =   "Load Compare File..."
         Index           =   1
      End
      Begin VB.Menu mnuH 
         Caption         =   "Export..."
         Index           =   2
      End
      Begin VB.Menu mnuH 
         Caption         =   "Calculate MD5"
         Index           =   3
      End
   End
   Begin VB.Menu mnuFileList 
      Caption         =   "File List"
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CBM-Transfer - Copyright (C) 2007-2017 Steve J. Gray
' ====================================================
' frmMenu - Menu selection Dispatch
'
' The GUI:
' - Set the lobal variable MenuForm to 1=Main, or 2=Viewer
' - Selects a menu to pop up at the mouse positon
' The MENU:
' - When the menu is selected it calls the subroutines here.
' - Menu items have a number from 1 to x which are passed back to the form specified with 'MenuForm'
' - Menu numbers are assigned values in blocks to group menus
' The Target FORM:
' - Has one routine to handle menu selection
' - Acts on the menu selection using the number

Option Explicit

'=========================
' MAIN FORM MENUS
'=========================
' Menu 0 - Local Options      (Items   1- 99)
' Menu 1 - Directory Options  (Items 101-199)
' Menu 2 - Theme Menu         (Items 201-299)
' Menu 3 - Character Encoding (Items 300-399) Shared
' Menu 4 - Device Control     (Items 400-499)
' Menu 5 - Font Size          (Items 501-599) Shared

'-- Local Options Menu (1-99)
Private Sub mnuLocal_Click(Index As Integer)
    Call frmMain.DoMenu(Index)                                              'Menu Starts at 1
End Sub

'-- Directory Options Menu (101-199)

Private Sub mnuImg_Click(Index As Integer)
    Call frmMain.DoMenu(Index + 100)                                        'Menu Starts at 101
End Sub

'-- Theme Menu (201-299)
Private Sub mnuTheme_Click(Index As Integer)
    Call frmMain.DoMenu(Index + 200)                                        'Menu Starts at 201
End Sub

'****************************
'-- Character Encoding Menu (301-399)
' Reserved. Handled Below
'****************************

'-- Device Control Menu (401-499)
' Includes: Initialize,Validate,Format, Change Device#, Set Single/Double sided, Re-Scan devices

Private Sub mnuDev_Click(Index As Integer)
    Call frmMain.DoMenu(Index + 400)                                        'Menu Starts at 401
End Sub


'=========================
' VIEWER MENUS
'=========================
' Menu 0 - Font Select          (  0- 99)
' Menu 1 - Font Convert submenu (101-199)
' Menu 2 - Screen Editor Theme  (201-299)
' Menu 4 - Hex Menu             (401-499)

'-- Font Menu (1-99)
Private Sub mnuFont_Click(Index As Integer)
    Call frmViewer.DoFMenu(Index)                                           'Menu Starts at 1
End Sub

'-- Font Convert Sub-Menu (101-199)
Private Sub mnuConvert_Click(Index As Integer)
    Call frmViewer.DoFMenu(Index + 100)                                     'Menu Starts at 101
End Sub
'-- Screen Editor - Select Theme Menu (201-299)
Private Sub mnuSE_Click(Index As Integer)

    Call frmViewer.DoFMenu(Index + 200)                                     'Menu Starts at 201

End Sub

'****************************
'--- Font Encoding Menu (301-399)
' Reserved. Handled below                                                   'Reserved Starts at 301
'****************************

'--- Hex Viewer Menu (401-499)
Private Sub mnuH_Click(Index As Integer)
    Call frmViewer.DoFMenu(Index + 400)                                     'Menu Starts at 401
End Sub

'****************************
'--- Font Size Menu (501-599)
' Reserved. Handled below                                                   'Reserved Starts at 501
'****************************

'=========================
' SHARED MENUS
'=========================
' Menu 3 - Font Encoding Menu  (301-399)
' Menu 5 - Font Size Menu      (501-599)

' These menu are used in both Main Form and Viewer form. The MenuForm Global variable is set to specify the form
' 1 = Main Form
' 2 = Viewer Form

'-- Character Encoding Menu (301-399)
' (PETSCII, Screen, ASCII etc)
Public Sub mnuEnc_Click(Index As Integer)
    
    If MenuForm = 1 Then Call frmMain.DoMenu(Index + 300)                   'Menu Starts at 301
    If MenuForm = 2 Then Call frmViewer.DoFMenu(Index + 300)                'Menu Starts at 301

End Sub

'--- Font Size Menu (501-599)
Private Sub mnuFS_Click(Index As Integer)
    If MenuForm = 1 Then Call frmMain.DoMenu(Index + 500)                    'Menu Starts at 1
    If MenuForm = 2 Then Call frmViewer.DoFMenu(Index + 500)                 'Menu Starts at 1
End Sub

