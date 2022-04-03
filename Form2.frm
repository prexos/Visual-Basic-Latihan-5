VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
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
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton backbttn 
      Caption         =   "KEMBALI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7800
      TabIndex        =   17
      Top             =   6720
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   8040
      TabIndex        =   7
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   8040
      TabIndex        =   6
      Top             =   2640
      Width           =   3135
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   8040
      TabIndex        =   5
      Top             =   3120
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   8040
      TabIndex        =   4
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Height          =   360
      Left            =   8040
      TabIndex        =   3
      Top             =   4080
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      Height          =   360
      Left            =   8040
      TabIndex        =   2
      Top             =   4680
      Width           =   3135
   End
   Begin VB.TextBox Text6 
      Height          =   360
      Left            =   8040
      TabIndex        =   1
      Top             =   5160
      Width           =   3135
   End
   Begin VB.TextBox Text7 
      Height          =   360
      Left            =   8040
      TabIndex        =   0
      Top             =   5640
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "No. Pegawai"
      Height          =   255
      Left            =   6120
      TabIndex        =   16
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Pegawai"
      Height          =   255
      Left            =   6120
      TabIndex        =   15
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Bagian"
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Gaji Pokok"
      Height          =   255
      Left            =   6120
      TabIndex        =   13
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Tunjangan"
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Bonus"
      Height          =   255
      Left            =   6120
      TabIndex        =   11
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Pajak"
      Height          =   255
      Left            =   6120
      TabIndex        =   10
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Total Gaji"
      Height          =   255
      Left            =   6120
      TabIndex        =   9
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "RINCIAN DATA GAJI KARYAWAN"
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
      Left            =   5640
      TabIndex        =   8
      Top             =   1200
      Width           =   5895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub backbttn_Click()
 'Menampilkan Form 1
 Form1.Show
 'Fokus kepada Text 1
 Form1.cetak.SetFocus
End Sub
