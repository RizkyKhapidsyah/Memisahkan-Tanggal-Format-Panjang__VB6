VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memisahkan Tanggal Format Panjang"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5715
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memisahkan komponen tanggal dari format tanggal
'panjang (dd mmmm yyyy) dan menampilkannya dengan
'menggunakan kata kunci di depan setiap komponen
'Masukkan tanggal dengan format: dd mmmm yyyy
'ke dalam Text1 (contoh: 22 Januari 2002)
'lalu klik Command1, maka akan menghasilkan:
'Tanggal 22 Bulan Januari Tahun 2002

Private Sub Command1_Click()
'Inisialisasi variabel yg digunakan
Dim strTanggal As String, i As Integer
Dim huruf As String * 1, Temp As String
Dim Lokasi1 As Integer, Lokasi2 As Integer
Dim Tanggal As String, Bulan As String, Tahun As String

Temp = "" 'Inisialisasi menampung huruf per huruf
  
  'Periksa, jika tanggal tidak valid, atau formatnya
  'tidak sama dengan format tanggal panjang...
  If Not IsDate(Text1.Text) Or Text1.Text <> Format(Text1.Text, "dd mmmm yyyy") Then
     'Tampilkan pesan...
     MsgBox "Tanggal/format-nya salah!", _
            vbCritical, "Tanggal Salah"
     'Kursor kembali ke Text1
     Text1.SetFocus
     SendKeys "{Home}+{End}"
     Exit Sub 'Keluar dari prosedur
  Else
  
  'Jika tanggal valid, tampung data tanggal
  strTanggal = Text1.Text
  
  'Ulangi huruf demi huruf dari awal sampai akhir
  For i = 1 To Len(strTanggal)
    
    'Tampung setiap satu huruf saja pada posisi ke-i
     huruf = Chr(Asc(Mid(strTanggal, i, 1)))
    
    'Tampung dan tambahkan huruf, demikian seterusnya..
     Temp = Temp + Chr(Asc(Mid(strTanggal, i, 1)))
    
    'Cari posisi karakter spasi pertama untuk
    'mendapatkan posisi string Bulan, yaitu posisi
    'spasi pertama + 1
    'Jika ada spasi dan panjang huruf yg sudah
    'ditampung masih lebih kecil dari 4, berarti itu
    'spasi I...
    If Len(Trim(huruf)) < 1 And Len(Temp) < 4 Then
       'Lokasi1 untuk mengambil posisi awal string
       'Bulan
       Lokasi1 = i + 1
  
       'Jangan lupa, tampung tanggalnya mulai dari
       'posisi
       Tanggal = Left(Temp, Lokasi1 - 2)
    End If
      
    'Jika terdapat lagi spasi berikutnya, di mana
    'panjang string Temp harus lebih besar dari 4 di
    'atas...
    If Len(Trim(huruf)) < 1 And Len(Temp) > 4 Then
       'awal sampai posisi Lokasi1 dikurangi 2
       'Dikurangi 2, karena bisa saja string Tanggal
       'hanya 1 digit, atau bisa juga 2 digit
       
       'Tampung posisi spasi tsb ditambah satu
       'untuk posisi string Tahun
       Lokasi2 = i + 1
       
       'Tampung string Bulan, mulai dari tengah pada
       'posisi Lokasi1, sebanyak (Lokasi2 dikurangi
       'dengan (Lokasi1 kurang 1))
       Bulan = Mid(Temp, Lokasi1, Lokasi2 - Lokasi1 - 1)
    End If
    
    'Jika counter lebih besar dari posisi Lokasi2
    'dan nilai counter sudah sama dengan panjang
    'strTanggal
    If i > Lokasi2 And i = Len(strTanggal) Then
       
       'Tampung string Tahun...
       Tahun = Mid(Temp, Lokasi2, 4)
    End If
  Next i  'Akhir pemeriksaan huruf per huruf
  'tampilkan hasilnya dalam bentuk string dengan
  'kata kunci penjelasan di depan setiap komponen...
  MsgBox "Tanggal " & Tanggal & _
         " Bulan " & Bulan & _
         " Tahun " & Tahun
    End If
End Sub


