VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
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
   Begin VB.CommandButton Command2 
      Caption         =   "Simpan dan Isi lagi"
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
      Left            =   5040
      TabIndex        =   29
      Top             =   8880
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   9240
      TabIndex        =   27
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   118226945
      CurrentDate     =   44722
   End
   Begin VB.Data Data3 
      Caption         =   "DStock"
      Connect         =   "Access"
      DatabaseName    =   "D:\Kuliah\Semester 4\Pemrograman\Visual-Basic-Latihan-10-main\penjualan1.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Left            =   9120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TSTOCK"
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Data Data2 
      Caption         =   "DCustomer"
      Connect         =   "Access"
      DatabaseName    =   "D:\Kuliah\Semester 4\Pemrograman\Visual-Basic-Latihan-10-main\penjualan1.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Left            =   9120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TCUSTOMER"
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Data Data1 
      Caption         =   "DJual"
      Connect         =   "Access"
      DatabaseName    =   "D:\Kuliah\Semester 4\Pemrograman\Visual-Basic-Latihan-10-main\penjualan1.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   9120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TJUAL"
      Top             =   3720
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Selesai"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8400
      TabIndex        =   26
      Top             =   8880
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Data1"
      Height          =   360
      Left            =   3840
      TabIndex        =   14
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identitas Customer"
      Height          =   2655
      Left            =   2280
      TabIndex        =   7
      Top             =   2640
      Width           =   5295
      Begin VB.TextBox txtnocust 
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Nomor"
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Nama"
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Alamat"
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   1800
         Width           =   615
      End
   End
   Begin VB.TextBox txtnostok 
      Alignment       =   2  'Center
      DataSource      =   "Data1"
      Height          =   360
      Left            =   2160
      TabIndex        =   6
      Top             =   6480
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataSource      =   "Data1"
      Height          =   360
      Left            =   4320
      TabIndex        =   5
      Top             =   6480
      Width           =   2535
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      DataSource      =   "Data1"
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Top             =   6480
      Width           =   2535
   End
   Begin VB.TextBox txtunit 
      Alignment       =   2  'Center
      DataSource      =   "Data1"
      Height          =   360
      Left            =   10080
      TabIndex        =   3
      Top             =   6480
      Width           =   1815
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   12360
      TabIndex        =   2
      Top             =   6480
      Width           =   2175
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      DataSource      =   "Data1"
      Height          =   360
      Left            =   12360
      TabIndex        =   1
      Top             =   7080
      Width           =   2175
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   12360
      TabIndex        =   0
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Label Label15 
      Caption         =   "Tanggal Penjualan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   28
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "CV BITFINEX"
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
      Left            =   7440
      TabIndex        =   25
      Top             =   0
      Width           =   2295
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2040
      X2              =   15360
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      Caption         =   "FAKTUR PENJUALAN"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   24
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Nomor Faktur"
      Height          =   255
      Left            =   2280
      TabIndex        =   23
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lbltgl 
      Caption         =   "Label Tanggal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11040
      TabIndex        =   22
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2040
      X2              =   15360
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label8 
      Caption         =   "Nomor Stok"
      Height          =   255
      Left            =   2400
      TabIndex        =   21
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Nama Stok"
      Height          =   255
      Left            =   5160
      TabIndex        =   20
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "Harga Jual"
      Height          =   255
      Left            =   7920
      TabIndex        =   19
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Unit Jual"
      Height          =   255
      Left            =   10560
      TabIndex        =   18
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Nilai Jual"
      Height          =   255
      Left            =   12960
      TabIndex        =   17
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Besaran Potongan"
      Height          =   255
      Left            =   10200
      TabIndex        =   16
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label Label14 
      Caption         =   "Nilai Penjualan Bersih"
      Height          =   255
      Left            =   9960
      TabIndex        =   15
      Top             =   7680
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    End
End Sub

Private Sub Command2_Click()
    'Data disimpan ke tabel Jual
        Data1.Recordset.AddNew
        Data1.Recordset!NOFAKTUR = Text1
        Data1.Recordset!TGLTRANS = DTPicker1.Value
        Data1.Recordset!NOCUST = txtnocust
        Data1.Recordset!NOSTOK = txtnostok
        Data1.Recordset!UNITJUAL = txtunit
        Data1.Recordset!HARGAJUAL = Text8
        Data1.Recordset!POTONGAN = Text11
        Data1.Recordset.Update
    'Data customer diedit
        Data2.Recordset.Edit
        Data2.Recordset!SALDOHUTANG = Data2.Recordset!SALDOHUTANG + Val(Text12.Text)
        Data2.Recordset.Update
    'Data Stok
        Data3.Recordset.Edit
        Data3.Recordset!UNITSTOCK = Data3.Recordset!UNITSTOCK - Val(txtunit.Text)
        Data3.Recordset.Update
    'Lain-lain
        Text1.Text = ""
        txtnocust.Text = ""
        Text3.Enabled = True
        Text3.Text = ""
        Text4.Enabled = True
        Text4.Text = ""
        txtnostok.Text = ""
        Text7.Enabled = True
        Text7.Text = ""
        txtunit.Text = ""
        Text8.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
    'Fokus text nomor faktur
        Text1.SetFocus
End Sub

Private Sub DTPicker1_Click()
    lbltgl.Caption = DTPicker1.Value
End Sub

Private Sub Form_Activate()
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    'Fullscreen
    Form1.WindowState = 2
End Sub

Private Sub Text11_LostFocus()
    If txtunit.Text = "" Then
        txtunit.SetFocus
    Else
        Text12.Text = Val(Text10.Text) * Val(Text11.Text)
        Command1.SetFocus
    End If
End Sub

Private Sub Text8_LostFocus()
    If txtunit.Text = "" Then
        txtunit.SetFocus
    Else
        Text10.Text = Val(Text8.Text) * Val(txtunit.Text)
        txtunit.SetFocus
    End If
End Sub

Private Sub txtnocust_LostFocus()
    Cari = "NOCUST='" + txtnocust.Text + "'"
    Data2.Recordset.FindFirst Cari
    If Data2.Recordset.NoMatch Then
        Respon = MsgBox("Data tidak ditemukan, cari lainnya?", vbYesNo, "Pencarian Customer")
            If Respon = vbYes Then
                txtnocust.Text = ""
                txtnocust.SetFocus
            Else
                Command1.SetFocus
            End If
    Else
        Text3.Text = Data2.Recordset!NAMACUST
        Text4.Text = Data2.Recordset!ALAMATCUST
        Text3.Enabled = False
        Text4.Enabled = False
        txtnostok.SetFocus
    End If
End Sub

Private Sub txtnostok_LostFocus()
    Cari = "NOSTOCK='" + txtnostok.Text + "'"
    Data3.Recordset.FindFirst Cari
    If Data3.Recordset.NoMatch Then
        Respon = MsgBox("Data Stok tidak ditemukan, cari lainnya?", vbYesNo, "Pencarian Stok")
            If Respon = vbYes Then
                txtnostok.Text = ""
                txtnostok.SetFocus
            Else
                Command1.SetFocus
            End If
    Else
        Text7.Text = Data3.Recordset!NAMASTOCK
        Text7.Enabled = False
        Text8.SetFocus
    End If
End Sub

Private Sub txtunit_LostFocus()
    If Text8.Text = "" Then
        Text8.SetFocus
    Else
        Text10.Text = Val(Text8.Text) * Val(txtunit.Text)
        Text11.SetFocus
    End If
End Sub
