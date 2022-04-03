VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
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
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton minimizescrnn 
      Caption         =   "Minimize"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton fullscrn 
      Caption         =   "Full Screen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton selesai 
      Caption         =   "Selesai"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9360
      TabIndex        =   17
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton cetak 
      Caption         =   "Cetak"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4800
      TabIndex        =   16
      Top             =   6600
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Height          =   360
      Left            =   7560
      TabIndex        =   10
      Top             =   5640
      Width           =   3135
   End
   Begin VB.TextBox Text6 
      Height          =   360
      Left            =   7560
      TabIndex        =   9
      Top             =   5160
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      Height          =   360
      Left            =   7560
      TabIndex        =   8
      Top             =   4680
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Height          =   360
      Left            =   7560
      TabIndex        =   7
      Top             =   4080
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   7560
      TabIndex        =   6
      Top             =   3600
      Width           =   3135
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   7560
      TabIndex        =   5
      Top             =   3120
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   7560
      TabIndex        =   4
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   7560
      TabIndex        =   3
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label9 
      Caption         =   "INPUT DATA GAJI KARYAWAN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   19
      Top             =   1200
      Width           =   5415
   End
   Begin VB.Label Label8 
      Caption         =   "Total Gaji"
      Height          =   255
      Left            =   5640
      TabIndex        =   15
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Pajak"
      Height          =   255
      Left            =   5640
      TabIndex        =   14
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Bonus"
      Height          =   255
      Left            =   5640
      TabIndex        =   13
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Tunjangan"
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Gaji Pokok"
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Bagian"
      Height          =   255
      Left            =   5640
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Pegawai"
      Height          =   255
      Left            =   5640
      TabIndex        =   1
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "No. Pegawai"
      Height          =   255
      Left            =   5640
      TabIndex        =   0
      Top             =   2160
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cetak_Click()
 'Load Form 2
 Load Form2
 'Menampilkan Form 2
 Form2.Show
 'Form 2 adalah Full Screen
 Form2.WindowState = 2
 
 'Menginput dari Form 1 ke Form 2
 Form2.Text1 = Form1.Text1
 Form2.Text2 = Form1.Text2
 Form2.Text3 = Form1.Text3
 Form2.Text4 = Form1.Text4
 Form2.Text5 = Form1.Text5
 Form2.Text6 = Form1.Text6
 Form2.Text7 = Form1.Text7
 Form2.Combo1 = Form1.Combo1
 
 'Text Form 2 di nonaktifkan
 Form2.Combo1.Enabled = False
 Form2.Text1.Enabled = False
 Form2.Text2.Enabled = False
 Form2.Text3.Enabled = False
 Form2.Text4.Enabled = False
 Form2.Text5.Enabled = False
 Form2.Text6.Enabled = False
 Form2.Text7.Enabled = False
End Sub

Private Sub Form_Activate()
 'Form Window
 Form1.WindowState = 2
 
 'Combo 1
 Combo1.AddItem "Akuntan"
 Combo1.AddItem "Manajer"
 Combo1.AddItem "CEO"
 Combo1.AddItem "Satpam"
 Combo1.AddItem "Helper"
End Sub

Private Sub fullscrn_Click()
 'Jika Window State adalah Full Screen
 If Form1.WindowState = 2 Then
 'Lalu melakukan Restore Down
 Form1.WindowState = 0
 Else
 'Lainnya melakukan Full Screen lagi
 Form1.WindowState = 2
 End If
End Sub

Private Sub minimizescrnn_Click()
 'Tombol ini digunakan untuk minimize screen
 Form1.WindowState = 1
End Sub

Private Sub selesai_Click()
 End
End Sub

Private Sub Text3_LostFocus()
 'Menghitung untuk isi dari Text4 (Tunjangan)
 Text4 = Val(Text3.Text) * 0.25
 Text5.SetFocus
End Sub

Private Sub Text5_LostFocus()
 Text6 = Val(Val(Text3) + Val(Text4) + Val(Text5)) * 0.1
 Text7 = Val(Val(Text3) + Val(Text4) + Val(Text5)) - Val(Text6)
 cetak.SetFocus
End Sub
