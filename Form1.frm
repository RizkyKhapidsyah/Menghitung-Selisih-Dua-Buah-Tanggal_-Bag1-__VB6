VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menghitung Selisih Dua Buah Tanggal (1)"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   1800
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim hari As Integer, bulan As Integer, tahun As Integer
    hari = DateTime.DateDiff("d", _
           CDate("22/01/1973"), _
           CDate("22/01/2002")) 'Menghasilkan 10592
    
    bulan = DateTime.DateDiff("m", _
           CDate("22/01/1973"), _
           CDate("22/01/2002")) 'Menghasilkan 348
    
    tahun = DateTime.DateDiff("yyyy", _
           CDate("22/01/1973"), _
           CDate("22/01/2002")) 'Menghasilkan 29
    
    MsgBox "Selisih antara tanggal 22/01/1973" & _
           vbCrLf & _
           "dengan tanggal 22/01/2002 " & vbCrLf & _
           "menghasilkan sebagai berikut: " & _
           vbCrLf & "" & vbCrLf & _
           " " & Format(hari, "#,#") & _
           " hari, ATAU" & vbCrLf & _
           " " & Format(bulan, "#,#") & _
           " bulan, ATAU" & vbCrLf & _
           " " & Format(tahun, "#,#") & " tahun", _
           vbInformation, "DateDiff"
End Sub


