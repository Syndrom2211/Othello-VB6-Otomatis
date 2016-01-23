VERSION 5.00
Begin VB.Form Depan 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Othello Classic"
   ClientHeight    =   4650
   ClientLeft      =   3735
   ClientTop       =   3435
   ClientWidth     =   4650
   FillColor       =   &H0080FF80&
   Icon            =   "Depan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Depan.frx":014A
   MousePointer    =   99  'Custom
   Picture         =   "Depan.frx":029C
   ScaleHeight     =   4650
   ScaleWidth      =   4650
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P e t a k  V e r s i o n "
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
   End
End
Attribute VB_Name = "Depan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================Deklarasi Fade & Timer==========================
Option Explicit

Private Sub Info_Click()
MsgBox "Info : " & vbCrLf & vbCrLf & "Kelompok 12 - Analisis Algoritma 8" & vbCrLf & "DFS dan Minimax" & vbCrLf & vbCrLf & "1. Firdamdam.Sasmita (10114175)" & vbCrLf & "2. Fajar (10114495)" & vbCrLf & "3. GunGun Abdullah (10114197)" & vbCrLf & vbCrLf & "Universitas Komputer Indonesia" & vbCrLf & Chr(169) & " Othello Classic", vbInformation, "Info"
End Sub

Private Sub Mulai_Click()
Main.Show
End Sub

Private Sub Tentang_Click(Index As Integer)
MsgBox "Tentang Game : " & vbCrLf & vbCrLf & "Othello adalah permainan tradisional berbentuk papan untuk dimainkan oleh 2 pemain, yaitu hitam dan putih. Salah satu aturan dalam game ini pada awal permainan kamu akan diminta memilih warna hitam atau putih. Kondisi ketika salah seorang pemain menang, adalah dengan banyak nya kepingan yang dimilikinya ketika semua papan sudah terpenuhi, siapa yang banyak maka dia yang menang. Menaklukan pemain hanya dilakukan ketika lawan sudah terhimpit oleh kedua warna yang menyerang." & vbCrLf & vbCrLf & Chr(169) & " Othello Classic", vbInformation, "Tentang Game"
End Sub

Private Sub Keluar_Click()
'==========================Keluar==========================
Dim Keluar As String
Keluar = MsgBox("Keluar dari Games ?", vbExclamation + vbYesNo, Chr(169) & " Othello Classic")
If Keluar = vbYes Then
    Unload Me
End If
End Sub
