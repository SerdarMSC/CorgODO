VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "CargoOdo V1.1"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   4140
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1055
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   11
      Text            =   "KM bilgisi giriniz"
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Text            =   "Text8"
      Top             =   5160
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Text            =   "Text7"
      Top             =   5160
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Text            =   "Text6"
      Top             =   5160
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Text            =   "Text5"
      Top             =   5160
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Kilometre OKU"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Kilometre YAZ"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   2010
      Left            =   120
      Picture         =   "main.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim dosya As String
    Label2.Caption = "Gösterge tipi"
    Label1.Caption = "Program hazýr..."
End Sub
Private Sub Command1_Click()
Dim xfile
Dim yazilacakKM, yazilacakKMHEX, boy
Dim d1 As Byte
Dim d2 As Byte
Dim d3 As Byte
Dim d4 As Byte

If Text9.Text = "KM bilgisi giriniz" Or Text9.Text = "" Then
    MsgBox ("Lütfen kilometre bilgisi giriniz...")
Else
    If Left(Label2.Caption, 2) = "v2" Then
        yazilacakKM = (Hex(Round(Text9.Text / 8, 0)))
        boy = (Len(yazilacakKM))
        If boy = 1 Then
            yazilacakKMHEX = "0000000" & yazilacakKM
        ElseIf boy = 2 Then
            yazilacakKMHEX = "000000" & yazilacakKM
        ElseIf boy = 3 Then
            yazilacakKMHEX = "00000" & yazilacakKM
        ElseIf boy = 4 Then
            yazilacakKMHEX = "0000" & yazilacakKM
        ElseIf boy = 5 Then
            yazilacakKMHEX = "000" & yazilacakKM
        ElseIf boy = 6 Then
            yazilacakKMHEX = "00" & yazilacakKM
        ElseIf boy = 7 Then
            yazilacakKMHEX = "0" & yazilacakKM
        End If
        d4 = CLng("&H" & hexlookup(Right(yazilacakKMHEX, 2)))
        d3 = CLng("&H" & hexlookup(Right(Left(yazilacakKMHEX, 6), 2)))
        d2 = CLng("&H" & hexlookup(Right(Left(yazilacakKMHEX, 4), 2)))
        d1 = CLng("&H" & hexlookup(Right(Left(yazilacakKMHEX, 2), 2)))
        
        xfile = CommonDialog1.FileName & ".mod"
        FileCopy CommonDialog1.FileName, xfile
        For i = 1 To 32 Step 4
            Open xfile For Binary Access Write As #2
            Put #2, i, d4
            Put #2, (i + 1), d3
            Put #2, (i + 2), d2
            Put #2, (i + 3), d1
            Close #2
        Next i
        
        Label1.Caption = "Ýþlem baþarýlý..."
        Label2.Caption = "v2 tipi gösterge yazýldý."
        Label2.BackColor = &H0&
        Label1.BackColor = &HFF00&
        MsgBox ("Yeni kilometre bilgisi v2 tip göstergeye baþarýyla yazýldý...")
    
        ElseIf Left(Label2.Caption, 2) = "v1" Then
        xfile = CommonDialog1.FileName & ".mod"
        FileCopy CommonDialog1.FileName, xfile
        
        If Text9.Text < 10000 Then
            v1KMyaz xfile, 33, checksumliKM(Text9.Text) ' 1. Blok
            v1KMyaz xfile, 41, checksumliKM(Text9.Text + 1) ' 2. Blok
            v1KMyaz xfile, 49, checksumliKM(Text9.Text + 2) ' 3. Blok
            v1KMyaz xfile, 57, checksumliKM(Text9.Text + 3) ' 4. Blok
            v1KMyaz xfile, 65, checksumliKM(Text9.Text + 4) ' 5. Blok
            v1KMyaz xfile, 73, checksumliKM(Text9.Text + 5) ' 6. Blok
            v1KMyaz xfile, 81, checksumliKM(Text9.Text + 6) ' 7. Blok
            v1KMyaz xfile, 89, checksumliKM(Text9.Text + 7) ' 8. Blok
            v1KMyaz xfile, 97, checksumliKM(Text9.Text + 8) ' 9. Blok
        Else
            v1KMyaz xfile, 33, checksumliKM(Text9.Text - 5) ' 1. Blok
            v1KMyaz xfile, 41, checksumliKM(Text9.Text - 4) ' 2. Blok
            v1KMyaz xfile, 49, checksumliKM(Text9.Text - 3) ' 3. Blok
            v1KMyaz xfile, 57, checksumliKM(Text9.Text - 2) ' 4. Blok
            v1KMyaz xfile, 65, checksumliKM(Text9.Text - 1) ' 5. Blok
            v1KMyaz xfile, 73, checksumliKM(Text9.Text) ' 6. Blok
            v1KMyaz xfile, 81, checksumliKM(Text9.Text - 8) ' 7. Blok
            v1KMyaz xfile, 89, checksumliKM(Text9.Text - 7) ' 8. Blok
            v1KMyaz xfile, 97, checksumliKM(Text9.Text - 6) ' 9. Blok
        End If

        Label1.Caption = "Ýþlem baþarýlý..."
        Label2.Caption = "v1 tipi gösterge yazýldý."
        Label2.BackColor = &H0&
        Label1.BackColor = &HFF00&
        MsgBox ("Yeni kilometre bilgisi v1 tip göstergeye baþarýyla yazýldý...")

    End If

End If
End Sub

Private Sub Command2_Click()
Dim sfile
Dim ssByte1 As Byte
Dim ssByte2 As Byte
Dim ssByte3 As Byte
Dim ssByte4 As Byte
Dim km As String

CommonDialog1.Filter = "Eprom Dosyasý (*.bin) yada (*.hex)"
CommonDialog1.DefaultExt = "bin"
CommonDialog1.DialogTitle = "Okunan dosyayý seçiniz"
CommonDialog1.ShowOpen

sfile = CommonDialog1.FileName
    If FileLen(sfile) = 8192 Then '24c64 dosyasý boyutu
        Label2.Caption = "v2 tipi göstergeden veri okundu."
        Label2.BackColor = &HFF00&
        Open sfile For Binary As #1
        Get #1, 5, ssByte1
        Text1.Text = Hex(ssByte1)
        Get #1, 6, ssByte2
        Text2.Text = Hex(ssByte2)
        Get #1, 7, ssByte3
        Text3.Text = Hex(ssByte3)
        Get #1, 8, ssByte4
        Text4.Text = Hex(ssByte4)
        Close #1
        Text5.Text = hexlookup(Text1.Text)
        Text6.Text = hexlookup(Text2.Text)
        Text7.Text = hexlookup(Text3.Text)
        Text8.Text = hexlookup(Text4.Text)
        km = Text8.Text & Text7.Text & Text6.Text & Text5.Text
        Label1.Visible = True
        Label1.Caption = "Araç " & FormatNumber(hex2dec(km) * 8, 0) & " kilometrededir."
        Text9.Text = FormatNumber(hex2dec(km) * 8, 0)
    ElseIf FileLen(sfile) = 512 Then '24c64 dosyasý boyutu
        Label2.Caption = "v1 tipi göstergeden veri okundu."
        Label2.BackColor = &HFF00&
        Open sfile For Binary As #1
        Get #1, 73, ssByte1
        Text1.Text = Hex(ssByte1)
        Get #1, 74, ssByte2
        Text2.Text = Hex(ssByte2)
        Get #1, 75, ssByte3
        Text3.Text = Hex(ssByte3)
        Get #1, 76, ssByte4
        Text4.Text = Hex(ssByte4)
        Close #1
        If Len(Text1.Text) = 1 Then
            Text1.Text = "0" + Text1.Text
        End If
        If Len(Text2.Text) = 1 Then
            Text2.Text = "0" + Text2.Text
        End If
        If Len(Text3.Text) = 1 Then
            Text3.Text = "0" + Text3.Text
        End If
        If Len(Text4.Text) = 1 Then
            Text4.Text = "0" + Text4.Text
        End If
        km = Text1.Text & Text2.Text & Text3.Text & Text4.Text
        Label1.Visible = True
        Label1.Caption = "Araç " & FormatNumber(hex2dec(km), 0) & " kilometrededir."
        Text9.Text = FormatNumber(hex2dec(km), 0)
        Else
        MsgBox ("Bilinmeyen dosya...")
        
    End If

End Sub
Public Function hex2dec(h)
Dim l As Long: l = Len(h)
If l < 16 Then
hex2dec = CDec("&h0" & h)
If hex2dec < 0 Then hex2dec = hex2dec + 4294967296#
ElseIf l < 25 Then
hex2dec = hex2dec(Left$(h, l - 9)) * 68719476736# + CDec("&h" & Right$(h, 9))
End If
End Function
Public Function checksumliKM(KilometreCRC As Variant) As String
Dim crc1, crc2, crc3, crc4 As Variant
Dim HesapBoy, HesapKM, HesapKMHEX As Variant

    HesapKM = (Hex(Round(KilometreCRC, 0)))
    HesapBoy = (Len(HesapKM))
    If HesapBoy = 1 Then
        HesapKMHEX = "0000000" & HesapKM
    ElseIf HesapBoy = 2 Then
        HesapKMHEX = "000000" & HesapKM
    ElseIf HesapBoy = 3 Then
        HesapKMHEX = "00000" & HesapKM
    ElseIf HesapBoy = 4 Then
        HesapKMHEX = "0000" & HesapKM
    ElseIf HesapBoy = 5 Then
        HesapKMHEX = "000" & HesapKM
    ElseIf HesapBoy = 6 Then
        HesapKMHEX = "00" & HesapKM
    ElseIf HesapBoy = 7 Then
        HesapKMHEX = "0" & HesapKM
    End If

crc1 = hexlookup(Left(HesapKMHEX, 2))
crc2 = hexlookup(Right(Left(HesapKMHEX, 4), 2))
crc3 = hexlookup(Right(Left(HesapKMHEX, 6), 2))
crc4 = hexlookup(Right(HesapKMHEX, 2))

checksumliKM = HesapKMHEX + crc1 + crc2 + crc3 + crc4

End Function
Public Function v1KMyaz(dosya, adres, data As Variant)
Dim tempd As Byte
    Open dosya For Binary As #3
    For dongu = 1 To Len(data) / 2
    tempd = CLng("&H" & Right(Left(data, 2 + ((dongu - 1) * 2)), 2))
    Put #3, adres + dongu - 1, tempd
    Next dongu
    Close #3
End Function


Public Function hexlookup(hl As Variant) As String
Dim one As Variant
Dim oneHL As String
Dim two As Variant
Dim twoHL As String
If hl = 0 Then
    hexlookup = 0
Else

one = Left(hl, 1)
two = Right(hl, 1)

End If

If one = 0 Then
oneHL = "F"
ElseIf one = 1 Then
oneHL = "E"
ElseIf one = 2 Then
oneHL = "D"
ElseIf one = 3 Then
oneHL = "C"
ElseIf one = 4 Then
oneHL = "B"
ElseIf one = 5 Then
oneHL = "A"
ElseIf one = 6 Then
oneHL = 9
ElseIf one = 7 Then
oneHL = 8
ElseIf one = 8 Then
oneHL = 7
ElseIf one = 9 Then
oneHL = 6
ElseIf one = "A" Then
oneHL = 5
ElseIf one = "B" Then
oneHL = 4
ElseIf one = "C" Then
oneHL = 3
ElseIf one = "D" Then
oneHL = 2
ElseIf one = "E" Then
oneHL = 1
ElseIf one = "F" Then
oneHL = 0
End If

If two = 0 Then
twoHL = "F"
ElseIf two = 1 Then
twoHL = "E"
ElseIf two = 2 Then
twoHL = "D"
ElseIf two = 3 Then
twoHL = "C"
ElseIf two = 4 Then
twoHL = "B"
ElseIf two = 5 Then
twoHL = "A"
ElseIf two = 6 Then
twoHL = 9
ElseIf two = 7 Then
twoHL = 8
ElseIf two = 8 Then
twoHL = 7
ElseIf two = 9 Then
twoHL = 6
ElseIf two = "A" Then
twoHL = 5
ElseIf two = "B" Then
twoHL = 4
ElseIf two = "C" Then
twoHL = 3
ElseIf two = "D" Then
twoHL = 2
ElseIf two = "E" Then
twoHL = 1
ElseIf two = "F" Then
twoHL = 0
End If

hexlookup = oneHL & twoHL

End Function

Private Sub Image1_Click()
frmAbout.Show
End Sub
