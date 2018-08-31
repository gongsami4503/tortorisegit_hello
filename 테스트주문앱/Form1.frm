VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14595
   LinkTopic       =   "Form1"
   ScaleHeight     =   11040
   ScaleWidth      =   14595
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox Text2 
      Height          =   510
      Left            =   3000
      TabIndex        =   18
      Text            =   "160700001"
      Top             =   180
      Width           =   2610
   End
   Begin VB.TextBox txtResult 
      Height          =   450
      Index           =   14
      Left            =   9840
      TabIndex        =   17
      Top             =   2190
      Width           =   4605
   End
   Begin VB.TextBox txtResult 
      Height          =   450
      Index           =   13
      Left            =   9825
      TabIndex        =   16
      Top             =   1635
      Width           =   4605
   End
   Begin VB.TextBox txtResult 
      Height          =   450
      Index           =   12
      Left            =   9810
      TabIndex        =   15
      Top             =   1065
      Width           =   4605
   End
   Begin VB.TextBox txtResult 
      Height          =   450
      Index           =   11
      Left            =   4995
      TabIndex        =   14
      Top             =   3585
      Width           =   4605
   End
   Begin VB.TextBox txtResult 
      Height          =   450
      Index           =   10
      Left            =   5010
      TabIndex        =   13
      Top             =   3105
      Width           =   4605
   End
   Begin VB.TextBox txtResult 
      Height          =   450
      Index           =   9
      Left            =   5025
      TabIndex        =   12
      Top             =   2610
      Width           =   4605
   End
   Begin VB.TextBox txtResult 
      Height          =   450
      Index           =   8
      Left            =   5010
      TabIndex        =   11
      Top             =   2115
      Width           =   4605
   End
   Begin VB.TextBox txtResult 
      Height          =   450
      Index           =   7
      Left            =   5010
      TabIndex        =   10
      Top             =   1590
      Width           =   4605
   End
   Begin VB.TextBox txtResult 
      Height          =   450
      Index           =   6
      Left            =   5010
      TabIndex        =   9
      Top             =   1050
      Width           =   4605
   End
   Begin VB.TextBox txtResult 
      Height          =   450
      Index           =   5
      Left            =   195
      TabIndex        =   8
      Top             =   3555
      Width           =   4605
   End
   Begin VB.TextBox txtResult 
      Height          =   450
      Index           =   4
      Left            =   195
      TabIndex        =   7
      Top             =   3060
      Width           =   4605
   End
   Begin VB.TextBox txtResult 
      Height          =   450
      Index           =   3
      Left            =   195
      TabIndex        =   6
      Top             =   2565
      Width           =   4605
   End
   Begin VB.TextBox txtResult 
      Height          =   450
      Index           =   2
      Left            =   210
      TabIndex        =   5
      Top             =   2085
      Width           =   4605
   End
   Begin VB.TextBox txtResult 
      Height          =   450
      Index           =   1
      Left            =   210
      TabIndex        =   4
      Top             =   1590
      Width           =   4605
   End
   Begin VB.TextBox txtResult 
      Height          =   450
      Index           =   0
      Left            =   210
      TabIndex        =   3
      Top             =   1050
      Width           =   4605
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   6495
      Left            =   195
      TabIndex        =   2
      Top             =   4485
      Width           =   14295
      _Version        =   524288
      _ExtentX        =   25215
      _ExtentY        =   11456
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   8
      MaxRows         =   1
      SpreadDesigner  =   "Form1.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "바코드 검색"
      Height          =   510
      Left            =   6510
      TabIndex        =   1
      Top             =   165
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   510
      Left            =   195
      TabIndex        =   0
      Text            =   "8888800001"
      Top             =   180
      Width           =   2610
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents sp1 As YsSp
Attribute sp1.VB_VarHelpID = -1

' 검색되는 상품 그룹이 1개 이상이라도 모두 가져옴..
'            string sql = "" +
'                        "SELECT " +
'                        " M.vcPdtManCd " +
'                        ",M.vcPdtManNm " +
'                        ",M.iPdtMdsClsSn " +
'                        ",M.iPdtDgr1stSn " +
'                        ",M.iPdtDgr2ndSn " +
'                        ",M.iPdtDgr3rdSn " +
'                        ",M.iPdtDgrUsrSn " +
'                        ",M.vcPharmCmpNm " +
'                        ",I.vcStdCd " +
'                        ",I.vcRegDt " +
'                        ",I.iInsSn " +
'                        ",(SELECT vcClsNm FROM TB_CodeBase WHERE iSecSn=1 AND iClsCdSn=4 AND iClsSn=I.iInsSn) as vcInsNm " +
'                        ",I.iTakSn " +
'                        ",(SELECT vcClsNm FROM TB_CodeBase WHERE iSecSn=1 AND iClsCdSn=5 AND iClsSn=I.iTakSn) as vcTakNm " +
'                        ",I.vcUntNm " +
'                        ",P.vcPdtCd " +
'                        ",P.vcPdtNm " +
'                        ",P.vcBarCd " +
'                        ",P.vcRptBarCd " +
'                        ",P.fTotCn " +
'                        ",P.iRepYn " +
'                        ",P.mManPr " +
'                        ",P.mPr " +
'                        "FROM TB_Products as P " +
'                        "INNER JOIN TB_ProductsMain as M ON P.vcPdtManCd=M.vcPdtManCd " +
'                        "LEFT JOIN TB_ProductsInsurance as I ON M.vcPdtManCd=I.vcPdtManCd " +
'                        "WHERE M.vcPdtManCd IN  " +
'                        "(SELECT vcPdtManCd FROM TB_Products WHERE vcBarCd=@바코드 OR vcRptBarCd=@중복바코드 GROUP BY vcPdtManCd) " +
'                        "ORDER BY M.vcPdtManCd, P.iRepYn, P.fTotCn, P.vcPdtCd ";

Private Sub Command1_Click()
    Dim tmpStr As String
    Dim ws As New SoapClient30
    
    Dim i As Long
    Dim aa() As String
    Dim bb() As String
    
    On Error GoTo ErrHandle
    
    ws.MSSoapInit ("http://www.bsjumun.net:45030/Service.asmx?WSDL")
    tmpStr = ws.S_POS_Order_BS(Text1.Text, Text2.Text)
    
    MsgBox tmpStr
    
    txtResult(0).Text = ""
    txtResult(1).Text = ""
    txtResult(2).Text = ""
    txtResult(3).Text = ""
    txtResult(4).Text = ""
    txtResult(5).Text = ""
    txtResult(6).Text = ""
    txtResult(7).Text = ""
    txtResult(8).Text = ""
    txtResult(9).Text = ""
    txtResult(10).Text = ""
    txtResult(11).Text = ""
    txtResult(12).Text = ""
    txtResult(13).Text = ""
    txtResult(14).Text = ""
    sp1.MaxRows = 0
    
    aa = Split(tmpStr, "$")
    If UBound(aa) > 0 Then
        For i = 0 To UBound(aa) - 1
            bb = Split(aa(i), "#")
            If UBound(bb) >= 23 Then '현재 23까지 있음
                sp1.MaxRows = sp1.MaxRows + 1
                
                If i = 0 Then '최초한번만 공통값 가져옴
                    txtResult(0).Text = bb(0)
                    txtResult(1).Text = bb(1)
                    txtResult(2).Text = bb(2)
                    txtResult(3).Text = bb(3)
                    txtResult(4).Text = bb(4)
                    txtResult(5).Text = bb(5)
                    txtResult(6).Text = bb(6)
                    txtResult(7).Text = bb(7)
                    txtResult(8).Text = bb(8)
                    txtResult(9).Text = bb(9)
                    txtResult(10).Text = bb(10)
                    txtResult(11).Text = bb(11)
                    txtResult(12).Text = bb(12)
                    txtResult(13).Text = bb(13)
                    txtResult(14).Text = bb(14)
                End If
                
                sp1.Cell(sp1.MaxRows, 1) = bb(15)
                sp1.Cell(sp1.MaxRows, 2) = bb(16)
                sp1.Cell(sp1.MaxRows, 3) = bb(17)
                sp1.Cell(sp1.MaxRows, 4) = bb(18)
                sp1.Cell(sp1.MaxRows, 5) = bb(19)
                sp1.Cell(sp1.MaxRows, 6) = bb(20)
                sp1.Cell(sp1.MaxRows, 7) = bb(21)
                sp1.Cell(sp1.MaxRows, 8) = bb(22)
            End If
        Next i
    End If
    
    Exit Sub
ErrHandle:
    MsgBox Err.Number & " | " & Err.Description
End Sub

Private Sub Form_Load()
    Set sp1 = New YsSp
    Set sp1.SpSet = fpSpread1
    sp1.MaxRows = 0
End Sub
