VERSION 5.00
Object = "{9BD6A640-CE75-11D1-AF04-204C4F4F5020}#2.0#0"; "Mo20.ocx"
Begin VB.Form Form1 
   Caption         =   "Koordniat Peta Dunia"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   12405
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Cari"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15720
      TabIndex        =   5
      Top             =   9840
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   10080
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   12600
      TabIndex        =   4
      Top             =   9840
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7680
      TabIndex        =   3
      Top             =   9840
      Width           =   2775
   End
   Begin MapObjects2.Map Map1 
      Height          =   9255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   19935
      _Version        =   131072
      _ExtentX        =   35163
      _ExtentY        =   16325
      _StockProps     =   225
      BackColor       =   16711680
      BorderStyle     =   1
      BackColor       =   16711680
      Contents        =   "KoordinatPeta.frx":0000
   End
   Begin VB.Label Label2 
      Caption         =   "Koordinat Y :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   2
      Top             =   9840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Koordinat X :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   9840
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_collPoints As VBA.Collection
Private m_symPoint As MapObjects2.Symbol

Sub Map_ku()
Dim layer As New MapLayer
    Dim dc As New MapObjects2.DataConnection

    dc.Database = App.Path & "\WORLD"
    dc.Connect

    Set layer = New MapLayer
    Map1.Layers.Clear

    Set layer.GeoDataset = dc.FindGeoDataset("country")
    layer.Symbol.Color = moGreen
    Map1.Layers.Add layer
End Sub

Private Sub Command1_Click()
Dim X As Single
Dim Y As Single

If Text1.Text = "" And Text2.Text = "" Then MsgBox "Koordinat X Dan Y Belum Dimasukkan": Exit Sub
If Text1.Text = "" Then MsgBox "Koordinat X Belum Dimasukkan": Exit Sub
If Text2.Text = "" Then MsgBox "Koordinat Y Belum Dimasukkan": Exit Sub

Dim pt As MapObjects2.Point
Set pt = Map1.ToMapPoint(X, Y)
    pt.X = Text1.Text
    pt.Y = Text2.Text
m_collPoints.Add pt
Map1.Refresh
End Sub

Private Sub Form_Load()
Call Map_ku
  Dim rect As MapObjects2.Rectangle
Set rect = Map1.FullExtent
rect.ScaleRectangle 1.2
Set Map1.FullExtent = rect
Set Map1.Extent = rect

Set m_symPoint = New MapObjects2.Symbol
With m_symPoint
.SymbolType = moPointSymbol
.Style = moCircleMarker
.Color = moRed
.Size = 6
End With

Set m_collPoints = New VBA.Collection
End Sub

Private Sub Map1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call getXY(Me.Text1, Me.Text2, X, Y)
Dim pt As MapObjects2.Point
Set pt = Map1.ToMapPoint(X, Y)
m_collPoints.Add pt
Map1.Refresh
End Sub

Sub getXY(ByVal lblX, ByVal lblY, ByVal X, ByVal Y)
Dim pt As MapObjects2.Point
    Set pt = Map1.ToMapPoint(X, Y)
    Text1.Text = Format(Map1.ToMapPoint(X, Y).X)
    Text2.Text = Format(Map1.ToMapPoint(X, Y).Y)
End Sub

Private Sub Map1_AfterLayerDraw(ByVal index As Integer, ByVal canceled As Boolean, ByVal hDC As stdole.OLE_HANDLE)
Dim i As Integer
If index = 0 Then
For i = 1 To m_collPoints.Count
Map1.DrawShape m_collPoints(i), m_symPoint
Next i
End If
Set m_collPoints = New VBA.Collection
Map1.Refresh
End Sub
  
  Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
          Dim Jawab As Integer
             Jawab = MsgBox("Anda yakin akan keluar dari program?", _
             vbQuestion + vbYesNo, "Konfirmasi")

             If Jawab = vbNo Then Cancel = -1
          End Sub
