VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form BEGO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jebol Ms.Access 97/2000/XP Password"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   4920
      Top             =   4230
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstPass 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   3660
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   7065
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "&Get Password..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5655
      TabIndex        =   0
      Top             =   4230
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "kirim e-mail: mail42me@telkom.net"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4065
      MousePointer    =   10  'Up Arrow
      TabIndex        =   4
      Top             =   4785
      Width           =   3075
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Powered by"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   105
      TabIndex        =   3
      Top             =   4770
      Width           =   855
   End
   Begin VB.Image Image10 
      Height          =   585
      Left            =   90
      Picture         =   "Form1.frx":000C
      Top             =   4185
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password Editor:"
      Height          =   195
      Left            =   105
      TabIndex        =   2
      Top             =   135
      Width           =   1440
   End
End
Attribute VB_Name = "BEGO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'
' Visual Basic Source Code
' Situs mengenai seputar pemograman menggunakan visual basic.
' http://www.geocities.com/vb_bego
'
' Kirim pertanyaan anda ke:
'
' vbhelp_us@ yahoo.com
' mail42me@ telkom.net
' VanJava@ telkom.net
'
' Author: CyberBoy - BEGO admin
' Update terakhir, Maret - 2004
' » http://groups.yahoo.com/group/vb_bego 
'
' NB: Anda Boleh mengedit source code ini, tanpa menghilangkan
'     nama pembuat asli
'***************************************************************************

' VIVA PROGRAMMER INDONESIA
'

Option Explicit

 Sub GetMs97(Filename As String)
     'di procedure ini kita akan mencoba mengetahui password untuk
     'ms.access 97 terlebih dahulu
     
     'kita buat variable array sebanyak 20 dengan tipe byte.
     'kenapa harus 20? Ini disebabkan panjang maximal password
     'access adalah 20.

      Dim data(0 To 19) As Byte, Pwd As String

     'Sekarang kita buka dan ambil data dari file yang akan di crack
     'passwordnya.
      Open Filename For Binary As #1
          'Kita ambil data mulai dari posisi 67.
          'kenapa harus pd posisi 67? Ini dikarenakan password yang disimpan
          'oleh ms.access ada pada posisi tersebut
          Get #1, 67, data
      Close #1

      Dim MaxSize, I As Integer, TempPwd
      Dim EncDec As String, nKey

      'untuk enskripsi dibawah ini, saya tidak bisa menjelaskannya.
      'Karena terlalu panjang utk dijelaskan. Enskripsi ini hasil penelitian saya
      'jadi anda tinggal pake aja OK!

      'Panjang keseluruhan enskripsi ini tentunya sama dengan panjang max password (20).
      'Kemudian kita split ke variable nKey.
      EncDec = "86 FB EC 37 5D 44 9C FA C6 5E 28 E6 13 B6 8A 60 54 94 7B 36"
      nKey = Split(EncDec, " ")


      Dim spos As Integer
      'Nah sekarang kita gunakan metode/fungsi XOR untuk mendapatkan password aslinya
      'Nilai yang ada pada variable data, dibandingkan dengan nilai enskripsinya.
       For I = 0 To 19
          TempPwd = TempPwd & Chr(data(I) Xor ("&H" & nKey(spos)))
          'var ini digunakan untuk mengetahui panjang password yang ada pd file yang dicrack.
          spos = spos + 1
      Next I


      'hasilnya kita cetak ke listbox(lstpass)
      Dim inLen As Integer
          inLen = InStr(1, TempPwd, Chr(0))
          lstPass.AddItem "Nama File: " & Filename
          lstPass.AddItem "Ukuran   : " & FileLen(Filename) & " bytes"
          lstPass.AddItem "Panjang password: " & IIf(inLen = 0, 20, inLen - 1)
          lstPass.AddItem "---------------------"
          lstPass.AddItem TempPwd
  End Sub


' EOF For access 97 password


'Nah sekarang kita coba untuk access 2000/xp
 Sub GetMs2000XP(Filename As String)
     'kita buat variable array sebanyak 40 dengan tipe byte.
     'kenapa harus 40? Ini disebabkan panjang maximal password
     'access adalah 20, kemudian dikalikan 2 maka hasilnya 40.
     
     Dim data(39) As Byte, cek As Byte
     Open Filename For Binary As #1
          'Kita ambil data mulai dari posisi 67.
          'kenapa harus pd posisi 67? Ini dikarenakan password yang disimpan
          'oleh ms.access ada pada posisi tersebut
          Get #1, 67, data
          Get #1, 151, cek
     Close #1

     'Sebelum melanjutkan mecrack 2000, kita periksa dahulu versi dari file tersebut
     'jika versinya 97 maka kita panggil prosedur GetMs97 dan keluar dari rutin 2000
     If cek = 0 Then GetMs97 Filename: Exit Sub

     'Kita buat var2 pendukung
     Dim EncDec   As String
     Dim I        As Integer
     Dim H        As Integer, nKey
     Dim nHex     As String
     Dim Pwd      As String

     'untuk enskripsi dibawah ini, saya tidak bisa menjelaskannya.
     'Karena terlalu panjang utk dijelaskan. Enskripsi ini hasil penelitian saya
     'jadi anda tinggal pake aja OK!
    
     'Tentunya enskripsi berikut berbeda dengan enskripsi untuk msa97
      EncDec = "00 EC DB 9C 40 28 95 8A D2 7B 73 DF F1 13 49 B1 B2 79 14 7C"
      nKey = Split(EncDec, " ")

      'Kita cari tau panjang passwordnya, dengan metode XOR
      Dim inLen As Integer
      For H = 0 To UBound(nKey)
         If H Mod 2 <> 0 Then
            If (data(H * 2) Xor ("&H" & nKey(H))) = 0 Then
               inLen = H
               Exit For
            End If
      End If
      Next H

     'Hasil pencariannya kita cetak pada listpass
     lstPass.AddItem "Nama File: " & Filename
     lstPass.AddItem "Ukuran   : " & FileLen(Filename) & " bytes"
     lstPass.AddItem "Panjang password: " & IIf(inLen = 0, 20, inLen)
     lstPass.AddItem "---------------------"

     'Nah disini kita cari tau passwordnya
     'Kita gunakan looping sampai dengan 255 kali
     'ini dilakukan karena kita akan membadingkan
     'sampai dengan karakter pertama sampai karakter terakhir(255)
     For I = 0 To 255
         'looping kedua berfungsi untuk membandingkan nilai asli dari file
         'dengan nilai enskripsi
         For H = 0 To UBound(nKey)
            If H Mod 2 = 0 Then
               'membandingkan nilai
               nHex = Hex((("&H" & nKey(H)) Xor I))
            Else
               nHex = nKey(H)
            End If
               'membandingkan nilai
            Pwd = Pwd & Chr((data(H * 2) Xor ("&H" & nHex)))
         Next H
         'Cetak hasil enskripsi yang didapat ke lstpass
         If InStr(1, Pwd, String(20 - inLen, Chr(0))) Then
            If InStr(1, Pwd, String(20, Chr(0))) Then
               lstPass.List(2) = "nggak ada passwordnya"
            Else
               lstPass.AddItem Pwd
            End If
         ElseIf InStr(1, Pwd, Chr(0)) = 0 Then
            lstPass.AddItem Pwd
        End If
        Pwd = ""
     Next I
 End Sub


'Untuk tahap terakhir kita tulis code pada cmdExec
 Private Sub cmdExec_Click()
 With CD
      .CancelError = True
      On Error GoTo X
         .Filter = "Ms. Access 97/2000/XP|*.mdb"
         .ShowOpen
         lstPass.Clear
         GetMs2000XP .Filename
         Exit Sub
X:
 End With
End Sub


Private Sub Form_Load()
lstPass.AddItem "Jebol Password Ms.Access 97/2000/XP"
lstPass.AddItem "Copyright 2004 by CyberBoy - BEGO"
lstPass.AddItem "Update terakhir, April 2004"
lstPass.AddItem ""
lstPass.AddItem "* Open Source *"
lstPass.AddItem ""
lstPass.AddItem "Email: mail42@telkom.net"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Me.WindowState = vbMinimized
End
End Sub

Private Sub Label2_Click()
On Error Resume Next
Shell "explorer mailto:mail42me@telkom.net", 1
End Sub
