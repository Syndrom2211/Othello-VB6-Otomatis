VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Othello Classic"
   ClientHeight    =   5985
   ClientLeft      =   8490
   ClientTop       =   3435
   ClientWidth     =   8295
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Main.frx":014A
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Aksi"
      ForeColor       =   &H8000000E&
      Height          =   2655
      Index           =   0
      Left            =   5280
      TabIndex        =   7
      Top             =   3000
      Width           =   2295
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Keluar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CommandButton TentangKami 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Info"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton Tentang 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tentang"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Mulay 
         BackColor       =   &H8000000B&
         Caption         =   "Mulai"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Status"
      ForeColor       =   &H8000000E&
      Height          =   1335
      Index           =   1
      Left            =   5280
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
      Begin VB.Label PesanGiliran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   645
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label PesanGiliran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Timer Timer2 
      Left            =   7680
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Left            =   7680
      Top             =   2400
   End
   Begin VB.PictureBox BoardBG 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   720
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox BoardAktif 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4170
      Left            =   720
      ScaleHeight     =   278
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   290
      TabIndex        =   0
      Top             =   1200
      Width           =   4350
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "V e r s i  O t o m a t i s "
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   5280
      Width           =   4350
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   720
      Picture         =   "Main.frx":029C
      Top             =   120
      Width           =   6915
   End
   Begin VB.Image Image2 
      Height          =   4785
      Left            =   0
      Picture         =   "Main.frx":8496
      Top             =   1200
      Width           =   8295
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================
'GAME OTHELLO
'VERSI:OTOMATIS MAIN
'
'UNTUK MEMENUHI SYARAT TUGAS BESAR ANALISIS ALGORITMA 8
'KELOMPOK 12
'
'FIRDAMDAM SASMITA, FAJAR, GUNGUN ABDULLAH
'
'Keterangan :
'   Petak Hitam = Player / aku / MAX
'   Petak Putih = Musuh / kamu / MIN
'=======================================================================
Option Explicit

' Inisialisasi Definisi Variable dari A-Z
DefInt A-Z

' Inisialisasi Variable bertipe Boolean
Dim Info As Boolean 'Inisialisasi Variable Informasi untuk pengujian program

Dim Petak   'Inisialisasi Variable untuk 1 petak
Dim C(9, 9) 'Inisialisasi Array Minimax
Dim P(9, 9) 'Inisialisasi Array PetakHijau
Dim IndexKolom(9), IndexBaris(9), Jumlah 'Inisialisasi Variable Untuk Menyimpan kemungkinan menemukan petak
Dim BarisIndexX(8), BarisIndexY(8) 'Inisialisasi Variable mencari nilai arah dari x dan y
Dim warnamu, warnaku 'Inisialisasi Variable untuk memilih warna petak
Dim barisku, kolomku 'Inisialisasi Variable untuk baris dan kolom si player
Dim barismu, kolommu 'Inisialisasi Variable untuk baris dan kolom si musuh
Dim Giliran 'Inisialisasi Variable Giliran setiap pemain (Maksimal 60x)
Dim Baris1, Kolom1, Baris2, Kolom2 'Inisialisasi Variable pencarian batas frame selama permainan

'Inisialisasi Konstanta Skala Petak menggunakan bit DWORD
Const SRCCOPY = &HCC0020

'Inisialisasi Kondisi petak di board main
Const petakputih = 0
Const petakhitam = 3
Const petakhijau = 6

'Inisialisasi Konstanta Warna
Const Putih = 16777215 'Putih
Const Hitam = 0 'Hitam
Const Hijau = 2186785 'Hijau

'===================================BAGIANCODINGUMUM===================================
Private Sub Keluar_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Mulay_Click()
MulaiGame
End Sub

Private Sub Tentang_Click(Index As Integer)
MsgBox "Tentang Game : " & vbCrLf & vbCrLf & "Othello adalah permainan tradisional berbentuk papan untuk dimainkan oleh 2 pemain, yaitu hitam dan putih. Salah satu aturan dalam game ini pada awal permainan kamu akan diminta memilih warna hitam atau putih. Kondisi ketika salah seorang pemain menang, adalah dengan banyak nya kepingan yang dimilikinya ketika semua papan sudah terpenuhi, siapa yang banyak maka dia yang menang. Menaklukan pemain hanya dilakukan ketika lawan sudah terhimpit oleh kedua warna yang menyerang." & vbCrLf & vbCrLf & Chr(169) & " Othello Classic", vbInformation, "Tentang Game"
End Sub

Private Sub TentangKami_Click(Index As Integer)
MsgBox "Info : " & vbCrLf & vbCrLf & "Kelompok 12 - Analisis Algoritma 8" & vbCrLf & "DFS dan Minimax" & vbCrLf & vbCrLf & "1. Firdamdam.Sasmita (10114175)" & vbCrLf & "2. Fajar (10114495)" & vbCrLf & "3. GunGun Abdullah (10114197)" & vbCrLf & vbCrLf & "Universitas Komputer Indonesia" & vbCrLf & Chr(169) & " Othello Classic", vbInformation, "Info"
End Sub

Private Sub Form_Load()
    Petak = Int(BoardAktif.ScaleWidth / 9.5)
    SetKonstanta
    TampilField
    BoardAktif.BorderStyle = 0
    PesanGiliran(0).BorderStyle = 0
    PesanGiliran(1).BorderStyle = 0
End Sub
'===================================BAGIANCODINGUMUM===================================

'==================================BAGIANCODINGUTAMA1===================================
Private Sub SetKonstanta()
   Dim i, Baris, Kolom
   Dim txt As String
   Dim BarisDanKolom(1 To 8) As String
   
   ' Index Awal
   For i = 0 To 9
     C(i, 0) = petakhijau
     C(0, i) = petakhijau
     C(9, i) = petakhijau
     C(i, 9) = petakhijau
   Next i
   
   'Inisialisasi Pencarian
   Baris1 = 2: Kolom1 = 2: Baris2 = 7: Kolom2 = 7
   
   'Mobility
   For i = 1 To 8
     BarisIndexX(i) = Choose(i, 1, 1, 0, -1, -1, -1, 0, 1)
     BarisIndexY(i) = Choose(i, 0, 1, 1, 1, 0, -1, -1, -1)
   Next i
   
   'Petak Hijau
   BarisDanKolom(1) = "30 01 20 10 10 20 01 30"
   BarisDanKolom(2) = "01 01 03 03 03 03 01 01"
   BarisDanKolom(3) = "20 03 05 05 05 05 03 20"
   BarisDanKolom(4) = "10 03 05 00 00 05 03 10"
   BarisDanKolom(5) = "10 03 05 00 00 05 03 10"
   BarisDanKolom(6) = "20 03 05 05 05 05 03 20"
   BarisDanKolom(7) = "01 01 03 03 03 03 01 01"
   BarisDanKolom(8) = "30 01 20 10 10 20 01 30"
   
   For Baris = 1 To 8: For Kolom = 1 To 8
      P(Baris, Kolom) = Val(Mid(BarisDanKolom(Baris), (Kolom - 1) * 3 + 1, 3))
      C(Baris, Kolom) = petakhijau
   Next Kolom: Next Baris
      
   'Peletakan 4 petak pertama
   C(4, 4) = 3: C(4, 5) = 0
   C(5, 4) = 0: C(5, 5) = 3
   Giliran = 0
End Sub

Private Sub TampilField()
   Dim Baris, Kolom
   Dim X, Y
   Dim txt As String
   
   With BoardAktif
      
      .FontSize = 18
      
      ' Nomer Horizontal Off
      For Kolom = 1 To 8
         txt = Format(Kolom)
         X = (Kolom * Petak + Petak / 2) - .TextWidth(txt) / 2
         Y = (Petak - .TextHeight(txt)) / 2
         TampilNomor BoardAktif, txt, X + 1, Y + 1, Putih
         TampilNomor BoardAktif, txt, X, Y, RGB(240, 248, 255) 'Warna Text Di Board
      Next Kolom
      
      ' Nomer Vertikal Off
      For Baris = 1 To 8
         txt = Format(Baris)
         X = (Petak - .TextWidth(txt)) / 2
         Y = (Baris * Petak + Petak / 2) - .TextHeight(txt) / 2
         TampilNomor BoardAktif, txt, X + 1, Y + 1, Putih
         TampilNomor BoardAktif, txt, X, Y, RGB(240, 248, 255) 'Warna Text Di Board
      Next Baris
      
      'Petak - Petak
      For Baris = 1 To 8: For Kolom = 1 To 8
         Select Case C(Baris, Kolom)
           Case petakhitam: .FillColor = Hitam
           Case petakputih: .FillColor = Putih
           Case petakhijau: .FillColor = Hijau
         End Select
         BoardAktif.Line (Petak * Kolom, Petak * Baris)-Step(Petak, Petak), .FillColor, BF
         BoardAktif.Line (Petak * Kolom, Petak * Baris)-Step(Petak, Petak), QBColor(7), B
      Next Kolom, Baris
         
   End With
End Sub

Private Sub TampilNomor(pic As PictureBox, txt As String, X, Y, Color As Long)
   With pic
      .ForeColor = Color: .CurrentX = X: .CurrentY = Y
   End With
   pic.Print txt
End Sub
'==================================BAGIANCODINGUTAMA1===================================

'==================================BAGIANCODINGUTAMA2===================================
Private Sub MulaiGame()
   If Mulay.Caption = "Mulai" Then
      warnamu = petakputih: warnaku = petakhitam
      Mulay.Caption = "Proses..."
      Mainkan
   End If
End Sub

' AUTO MAX
Private Sub CariPetakTerbaikMax(Max, TotalPetak)
   Dim i, Baris, Kolom
   Dim SetiapPetak, PerKolom
   Dim siKolom, siBaris
   
   If Not (Baris1 * Kolom1 = 1 And Baris2 * Kolom2 = 64) Then
      For i = 2 To 7
         If C(2, i) <> petakhijau Then Baris1 = 1
         If C(7, i) <> petakhijau Then Baris2 = 8
         If C(i, 2) <> petakhijau Then Kolom1 = 1
         If C(i, 7) <> petakhijau Then Kolom2 = 8
      Next i
      End If
      
   ' Mencari ke semua petak lalu memilih petak kosong
   PerKolom = 0
   Max = 0: TotalPetak = 0
   For Baris = Baris1 To Baris2: For Kolom = Kolom1 To Kolom2
      If C(Baris, Kolom) = petakhijau Then
         If P(Baris, Kolom) < Max Then GoTo PencarianBerikutnya:
         PerKolom = 0
         
         For i = 0 To 8
            SetiapPetak = 0
            siKolom = Kolom: siBaris = Baris
            siKolom = siKolom + BarisIndexX(i)
            siBaris = siBaris + BarisIndexY(i)
            While C(siBaris, siKolom) = warnaku
               SetiapPetak = SetiapPetak + 1
               siKolom = siKolom + BarisIndexX(i)
               siBaris = siBaris + BarisIndexY(i)
            Wend
            If C(siBaris, siKolom) <> petakhijau And SetiapPetak <> 0 Then PerKolom = PerKolom + SetiapPetak
         Next i
         
         If PerKolom <> 0 Then
            If P(Baris, Kolom) > Max Then
               Max = P(Baris, Kolom)
               Jumlah = 0
               TotalPetak = PerKolom
               IndexKolom(0) = Kolom: IndexBaris(0) = Baris
               GoTo PencarianBerikutnya:
               End If
            If TotalPetak > SetiapPetak Then GoTo PencarianBerikutnya:
            If TotalPetak < SetiapPetak Then
               Jumlah = 0
               TotalPetak = SetiapPetak
               IndexKolom(0) = Kolom
               IndexBaris(0) = Baris
               Else
               Jumlah = Jumlah + 1
               IndexKolom(Jumlah) = Kolom
               IndexBaris(Jumlah) = Baris
               End If
            End If
         End If
            
PencarianBerikutnya:
DoEvents
Next Kolom: Next Baris
      
End Sub

' AUTO MIN
Private Sub CariPetakTerbaikMin(Min, TotalPetak)
   Dim i, Baris, Kolom
   Dim SetiapPetak, PerKolom
   Dim siKolom, siBaris
   
   If Not (Baris1 * Kolom1 = 1 And Baris2 * Kolom2 = 64) Then
      For i = 2 To 7
         If C(2, i) <> petakhijau Then Baris1 = 1
         If C(7, i) <> petakhijau Then Baris2 = 8
         If C(i, 2) <> petakhijau Then Kolom1 = 1
         If C(i, 7) <> petakhijau Then Kolom2 = 8
      Next i
      End If
      
   ' Mencari ke semua petak lalu memilih petak kosong
   PerKolom = 0
   Min = 0: TotalPetak = 0
   For Baris = Baris1 To Baris2: For Kolom = Kolom1 To Kolom2
      If C(Baris, Kolom) = petakhijau Then
         If P(Baris, Kolom) < Min Then GoTo PencarianBerikutnya:
         PerKolom = 0
         For i = 0 To 8
            SetiapPetak = 0
            siKolom = Kolom: siBaris = Baris
            siKolom = siKolom + BarisIndexX(i)
            siBaris = siBaris + BarisIndexY(i)
            While C(siBaris, siKolom) = warnamu
               SetiapPetak = SetiapPetak + 1
               siKolom = siKolom + BarisIndexX(i)
               siBaris = siBaris + BarisIndexY(i)
            Wend
            If C(siBaris, siKolom) <> petakhijau And SetiapPetak <> 0 Then PerKolom = PerKolom + SetiapPetak
         Next i
         If PerKolom <> 0 Then
            If P(Baris, Kolom) > Min Then
               Min = P(Baris, Kolom)
               Jumlah = 0
               TotalPetak = PerKolom
               IndexKolom(0) = Kolom: IndexBaris(0) = Baris
               GoTo PencarianBerikutnya:
               End If
            If TotalPetak > SetiapPetak Then GoTo PencarianBerikutnya:
            If TotalPetak < SetiapPetak Then
               Jumlah = 0
               TotalPetak = SetiapPetak
               IndexKolom(0) = Kolom
               IndexBaris(0) = Baris
               Else
               Jumlah = Jumlah + 1
               IndexKolom(Jumlah) = Kolom
               IndexBaris(Jumlah) = Baris
               End If
            End If
         End If
            
PencarianBerikutnya:
DoEvents
Next Kolom: Next Baris
      
End Sub

' Dapat Petak
Private Function JumlahPetak(Baris, Kolom, CariWarna)
   Dim SetiapPetak
   Dim TotalPetak
   Dim siKolom, siBaris
   Dim i
   Dim UbahWarna
   
   UbahWarna = IIf(CariWarna = petakhitam, petakputih, petakhitam)
   
   For i = 1 To 8 ' Memeriksa 8 arah
      SetiapPetak = 0 ' Inisialisasi SetiapPetak dengan nilai 0
      siKolom = Kolom: siBaris = Baris
      siKolom = siKolom + BarisIndexX(i)
      siBaris = siBaris + BarisIndexY(i)
      
      While C(siBaris, siKolom) = UbahWarna
         SetiapPetak = SetiapPetak + 1
         siKolom = siKolom + BarisIndexX(i)
         siBaris = siBaris + BarisIndexY(i)
      Wend
      
      If C(siBaris, siKolom) <> petakhijau And SetiapPetak <> 0 Then
         TotalPetak = TotalPetak + SetiapPetak
         siKolom = siKolom - BarisIndexX(i)
         siBaris = siBaris - BarisIndexY(i)
         
         While C(siBaris, siKolom) <> petakhijau
            C(siBaris, siKolom) = CariWarna
            siKolom = siKolom - BarisIndexX(i)
            siBaris = siBaris - BarisIndexY(i)
         Wend
         
       End If
         
   Next i
   JumlahPetak = TotalPetak
End Function

' Tanda Silang
Private Sub TandaSilang(Baris, Kolom)
   Dim txt As String, Warna As Long
   Dim X, Y
   Warna = RGB(40, 142, 40)
   BoardAktif.Line (Petak * Kolom, Petak * Baris)-Step(Petak, Petak), Warna
   BoardAktif.Line (Petak * Kolom, Petak * Baris + Petak)-Step(Petak, -Petak), Warna
   BoardAktif.FontSize = 18
   BoardAktif.ForeColor = Warna
   
   ' Kolom Horizontal
   txt = Format(Kolom)
   X = (Kolom * Petak + Petak / 2) - BoardAktif.TextWidth(txt) / 2
   Y = (Petak - BoardAktif.TextHeight(txt)) / 2
   TampilNomor BoardAktif, txt, X + 1, Y + 1, Hitam
   TampilNomor BoardAktif, txt, X, Y, Warna
   
   ' Kolom Vertikal
   txt = Format(Baris)
   X = (Petak - BoardAktif.TextWidth(txt)) / 2
   Y = (Baris * Petak + Petak / 2) - BoardAktif.TextHeight(txt) / 2
   TampilNomor BoardAktif, txt, X + 1, Y + 1, Hitam
   TampilNomor BoardAktif, txt, X, Y, Warna
   KecRan 500
End Sub

' Kondisi Menang
Private Sub KondisiMenang(pesanpesan As String)
   Dim Baris, Kolom
   Dim total1
   Dim total2
   Dim kondisi As String
   Dim kondisi2 As String
   
   total1 = 0: total2 = 0
   For Baris = 1 To 8: For Kolom = 1 To 8
      If C(Baris, Kolom) = warnaku Then total1 = total1 + 1 Else total2 = total2 + 1
   Next Kolom, Baris
   
   ' Kondisi Player
   If total2 < total1 Then
      kondisi = "Jumlah petak musuh" & total2 & " petak, dan player " & total1 & "."
      kondisi2 = "Player Menang !"
      End If
      
   ' Kondisi Musuh
   If total2 > total1 Then
      kondisi = "Jumlah petak player " & total1 & " petak, dan musuh " & total2 & "."
      kondisi2 = "Musuh Menang!"
      End If
   MsgBox pesanpesan & vbCrLf & vbCrLf & kondisi & vbCrLf & vbCrLf & kondisi2
   Pesan kondisi2, Putih
      
End Sub

' Kondisi Dilewat
Private Function Dilewat() As Boolean
   Dim SetiapPetak, i, Baris, Kolom, siBaris, siKolom
   
   Dilewat = True
   For Baris = 1 To 8: For Kolom = 1 To 8
   
      ' Mencari yang kosong
      If C(Baris, Kolom) = petakhijau Then
         For i = 1 To 8
            SetiapPetak = 0: siKolom = Kolom: siBaris = Baris
            Do
               siBaris = siBaris + BarisIndexY(i): siKolom = siKolom + BarisIndexX(i)
               If C(siBaris, siKolom) = warnaku Then SetiapPetak = SetiapPetak + 1
            Loop Until C(siBaris, siKolom) <> warnaku
            If C(siBaris, siKolom) = warnamu And SetiapPetak > 0 Then
               Dilewat = False
            End If
         Next i
         End If
         
   Next Kolom, Baris
End Function

' Kecepatan Random
Private Sub KecRan(kr)
   Dim waktu As Variant
   waktu = Timer
   While Timer - waktu < kr / 1000: DoEvents: Wend
End Sub

' Giliran Bermain
Private Function GiliranMain() As Boolean
   Giliran = Giliran + 1
   If Giliran = 60 Then
      KondisiMenang "Permainan Berakhir."
      MulaiLagi
      GiliranMain = False
      Else
      GiliranMain = True
      End If
End Function

' MulaiLagi
Private Sub MulaiLagi()
   Mulay.Caption = "Selesai"
End Sub

' Pesan
Private Sub Pesan(txt As String, Warna As Long)
   PesanGiliran(0).ForeColor = Warna
   PesanGiliran(0).Caption = txt
   PesanGiliran(1).Caption = txt
End Sub
'==================================BAGIANCODINGUTAMA2===================================

'==================================BAGIANCODINGMAIN===================================
Private Sub Mainkan()
   Dim Min
   Dim Max
   Dim Untung
   Dim Lewat As Boolean
   Dim RandomIndex
   
   SetKonstanta
   TampilField
   
kamu:
   Pesan "Giliran Musuh", Hitam
   CariPetakTerbaikMax Max, Untung
   
   ' Kondisi menemukan petak terbaik
   If Max > 0 Then
      If Jumlah > 0 Then
         Jumlah = Jumlah + 1
         RandomIndex = Int(Rnd(1) * Jumlah)
         kolomku = IndexKolom(RandomIndex): barisku = IndexBaris(RandomIndex)
      Else
         kolomku = IndexKolom(0): barisku = IndexBaris(0)
      End If
      
      TandaSilang barisku, kolomku
      Untung = JumlahPetak(barisku, kolomku, warnamu)
      C(barisku, kolomku) = warnamu
      TampilField
      
      If GiliranMain() = False Then Exit Sub
      Lewat = Dilewat()
      End If
   
aku:
   Pesan "Giliran Player", Putih
   CariPetakTerbaikMin Min, Untung
   
   ' Kondisi menemukan petak terbaik
   If Min > 0 Then
      If Jumlah > 0 Then
         Jumlah = Jumlah + 1
         RandomIndex = Int(Rnd(1) * Jumlah)
         kolommu = IndexKolom(RandomIndex): barismu = IndexBaris(RandomIndex)
      Else
         kolommu = IndexKolom(0): barismu = IndexBaris(0)
      End If
      
      TandaSilang barismu, kolommu
      Untung = JumlahPetak(barismu, kolommu, warnaku)
      C(barismu, kolommu) = warnaku
      TampilField
      
      If GiliranMain() = False Then Exit Sub
      Lewat = Dilewat()
      End If
   
    ' Kondisi SELESAI
    Lewat = Dilewat()
    If Lewat = True Then
       KondisiMenang "Permainan Berakhir!"
       MulaiLagi
       Exit Sub
    End If
    GoTo kamu:
   
End Sub
