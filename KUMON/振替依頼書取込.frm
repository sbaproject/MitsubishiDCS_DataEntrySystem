VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmFurikaeReqImport 
   Caption         =   "振替依頼書(取込)"
   ClientHeight    =   7965
   ClientLeft      =   2445
   ClientTop       =   2370
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11100
   Begin VB.CommandButton cmdDelete 
      Caption         =   "廃棄(&D)"
      Height          =   435
      Left            =   5340
      TabIndex        =   7
      Top             =   6900
      Width           =   1395
   End
   Begin VB.Frame fraProgressBar 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'なし
      Caption         =   "fraProgressBar"
      ForeColor       =   &H80000004&
      Height          =   290
      Left            =   1860
      TabIndex        =   12
      Top             =   7500
      Width           =   7060
      Begin MSComctlLib.ProgressBar pgrProgressBar 
         Height          =   255
         Left            =   15
         TabIndex        =   13
         Top             =   15
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
      End
   End
   Begin VB.ComboBox cboSort 
      Height          =   300
      ItemData        =   "振替依頼書取込.frx":0000
      Left            =   4500
      List            =   "振替依頼書取込.frx":000D
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   60
      Width           =   1695
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "マスタ反映(&U)"
      Height          =   435
      Left            =   6840
      TabIndex        =   8
      Top             =   6900
      Width           =   1395
   End
   Begin VB.CommandButton cmdErrList 
      Caption         =   "エラーリスト(&P)"
      Height          =   435
      Left            =   3480
      TabIndex        =   6
      Top             =   6900
      Width           =   1395
   End
   Begin FPSpread.vaSpread sprMeisai 
      Bindings        =   "振替依頼書取込.frx":0037
      Height          =   6315
      Left            =   180
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   420
      Width           =   10665
      _Version        =   196608
      _ExtentX        =   18812
      _ExtentY        =   11139
      _StockProps     =   64
      ButtonDrawMode  =   4
      ColsFrozen      =   6
      DAutoCellTypes  =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   51
      MaxRows         =   1000000
      SpreadDesigner  =   "振替依頼書取込.frx":004F
      UserResize      =   0
      VirtualMode     =   -1  'True
      VirtualScrollBuffer=   -1  'True
   End
   Begin VB.ComboBox cboImpDate 
      Height          =   300
      Left            =   1200
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   60
      Width           =   1935
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "チェック(&C)"
      Height          =   435
      Left            =   1980
      TabIndex        =   5
      Top             =   6900
      Width           =   1395
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "取込(&I)"
      Height          =   435
      Left            =   480
      TabIndex        =   4
      Top             =   6900
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "終了(&X)"
      Height          =   435
      Left            =   9360
      TabIndex        =   0
      Top             =   6900
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   8580
      Top             =   6900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ORADCLibCtl.ORADC dbcImport 
      Height          =   315
      Left            =   9120
      Top             =   7320
      Visible         =   0   'False
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   207
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DatabaseName    =   "dcssvr03"
      Connect         =   "kumon/kumon"
      RecordSource    =   "select * from tchogoshaimport"
   End
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  '下揃え
      Height          =   315
      Left            =   0
      TabIndex        =   11
      Top             =   7650
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "残り 9,999 件"
            TextSave        =   "残り 9,999 件"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblModoriCount 
      Caption         =   "【 口座戻り件数： 9,999 件 】"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6540
      TabIndex        =   15
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "表示順"
      Height          =   180
      Left            =   3780
      TabIndex        =   14
      Top             =   120
      Width           =   600
   End
   Begin VB.Label Label8 
      Caption         =   "取込日時"
      Height          =   180
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Width           =   780
   End
   Begin VB.Label lblSysDate 
      Caption         =   "Label26"
      Height          =   195
      Left            =   9540
      TabIndex        =   9
      Top             =   0
      Width           =   1275
   End
   Begin VB.Menu mnuFile 
      Caption         =   "ﾌｧｲﾙ(&F)"
      Begin VB.Menu mnuEnd 
         Caption         =   "終了(&X)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "ﾍﾙﾌﾟ(&H)"
      Begin VB.Menu mnuVersion 
         Caption         =   "ﾊﾞｰｼﾞｮﾝ情報(&A)"
      End
   End
End
Attribute VB_Name = "frmFurikaeReqImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//仮想モードですると動きが変だ！！！
#Const VIRTUAL_MODE = True
#Const DATA_DUPLICATE = False   '//振替依頼書は重複をチェック

#Const BLOCK_CHECK = False           '//チェック時のブロックがいくつあるか？を表示：デバック時のみ
#If BLOCK_CHECK = True Then           '//チェック時のブロックがいくつあるか？を表示：デバック時のみ
Private mCheckBlocks As Integer
#End If

Private mCaption As String
Private mAbort As Boolean
Private mForm As New FormClass
Private mSpread As New SpreadClass
Private mReg As New RegistryClass
Private mRimp As New FurikaeReqImpClass
Public mEditRow As Long      '//修正中の行番号

Private Type tpErrorStatus
    Field   As String
    Error   As Integer
    Message As String
End Type
Private mErrStts() As tpErrorStatus

Private Type tpHogoshaImport
    MochikomiBi     As String * 8   '//持ち込み日 2006/03/24 項目追加
    Keiyakusha      As String * 5   '//契約者番号
    Kyoshittsu      As String * 1   '//教室番号
    HogoshaNo       As String * 4   '//保護者番号
    HogoshaKana     As String * 40  '//保護者名(カナ)=>口座名義人
    HogoshaKanji    As String * 30  '//保護者名(漢字)
    SeitoShimei     As String * 50  '//生徒氏名
    BankCode        As String * 4   '//金融機関コード
    BankName        As String * 30  '//金融機関名
    ShitenCode      As String * 3   '//支店コード
    ShitenName      As String * 30  '//支店名
    YokinShumoku    As String * 1   '//預金種目
    KouzaBango      As String * 7   '//口座番号
    TuuchoKigou     As String * 3   '//通帳記号
    TuuchoBango     As String * 8   '//通帳番号
    FurikaeGaku     As String * 6   '//振替金額
    CrLf            As String * 2   '// CR + LF
End Type

Private Const cBtnCancel As String = "中止(&A)"
Private Const cBtnImport As String = "取込(&I)"
Private Const cBtnDelete As String = "廃棄(&D)"
Private Const cBtnCheck  As String = "チェック(&C)"
Private Const cBtnUpdate As String = "マスタ反映(&U)"
Private Const cVisibleRows As Long = 25
'Private Const cImportToUpdate As String = "U"
'Private Const cImportToInsert As String = "I"
Private Const cEditDataMsg  As String = "修正 => チェック処理をして下さい。"
Private Const cImportMsg    As String = "取込 => チェック処理をして下さい。"

Private Enum eSprCol
    eErrorStts = 1  'エラー内容：異常、正常、警告
    eItakuName      '   CIITKB  委託者名
    eKeiyakuCode    '   CIKYCD  契約者コード
    eKeiyakuName    '           契約者名
    eKyoshitsuNo    '   CIKSCD  教室番号
    eHogoshaCode    '   CIHGCD  保護者コード
    eHogoshaName    '   CIKJNM  保護者名(漢字)
    eHogoshaKana    '   CIKNNM  保護者名(カナ)=>口座名義人名
    eSeitoName      '   CISTNM  生徒氏名
    eFurikaeGaku    '   CISKGK  振替金額
    eKinyuuKubun    '   CIKKBN  金融機関区分
    eBankCode       '   CIBANK  銀行コード
    eBankName_m     '           銀行名(マスター)
    eBankName_i     '   CIBKNM  銀行名(取込)
    eShitenCode     '   CISITN  支店コード
    eShitenName_m   '           支店名(マスター)
    eShitenName_i   '   CISINM  支店名(取込)
    eYokinShumoku   '   CIKZSB  預金種目
    eKouzaBango     '   CIKZNO  口座番号
    eYubinKigou     '   CIYBTK  郵便局:通帳記号
    eYubinBango     '   CIYBTN  郵便局:通帳番号
    eKouzaName      '   CIKZNM  口座名義人=>保護者名(カナ)
    eMstUpdate      '//マスター反映フラグ
    eImpDate        '取込日
    eImpSEQ         'ＳＥＱ
    eUseCols = eKouzaName  '//表示する列は此処まで
    eMaxCols = 50   '//エラー列も含めて！
End Enum

Private Enum eSort
    eImportSeq
    eKeiyakusha
    eKinnyuKikan
End Enum
Private mMainSQL As String

Private Sub pLockedControl(blMode As Boolean, Optional vButton As CommandButton = Nothing)
    cmdImport.Enabled = blMode
    cmdCheck.Enabled = blMode
    cmdErrList.Enabled = blMode
    cmdDelete.Enabled = blMode
    cmdUpdate.Enabled = blMode
    cmdEnd.Enabled = blMode     '//処理途中で終了するとおかしくなるので終了も殺す！
    If Not vButton Is Nothing Then
        vButton.Enabled = True
    End If
End Sub

Private Function pMakeSQLReadData(Optional vErrColomns As Boolean = False) As String
    Dim sql As String
    
    sql = "SELECT * FROM(" & vbCrLf
    sql = sql & "SELECT " & vbCrLf
    'sql = sql & " CIERROR," & vbCrLf
#If SHORT_MSG Then
    sql = sql & " DECODE(CIERROR,-3,'取込',-2,'修正',-1,decode(cimupd,1,'警告','異常'),0,'正常',1,'警告','例外') as CIERRNM," & vbCrLf
#Else
    sql = sql & " CASE WHEN CIERROR = -2 THEN " & gdDBS.ColumnDataSet(cEditDataMsg, vEnd:=True) & vbCrLf
    sql = sql & "      WHEN CIERROR = -3 THEN " & gdDBS.ColumnDataSet(cImportMsg, vEnd:=True) & vbCrLf
    sql = sql & "      WHEN CIERROR IN(-1,+0,+1) THEN " & vbCrLf
    sql = sql & "           DECODE(CIERROR," & vbCrLf
    sql = sql & "               -1,decode(cimupd,1,'警告','異常')," & vbCrLf
    sql = sql & "               +0,'正常'," & vbCrLf
    sql = sql & "               +1,'警告'," & vbCrLf
    sql = sql & "               NULL" & vbCrLf
    sql = sql & "           ) || ' => ' || " & vbCrLf
    sql = sql & "       DECODE(CIOKFG," & vbCrLf
    sql = sql & "               " & mRimp.updInvalid & ",'" & mRimp.mUpdateMessage(mRimp.updInvalid) & "'," & vbCrLf
    sql = sql & "               " & mRimp.updWarnErr & ",'" & mRimp.mUpdateMessage(mRimp.updWarnErr) & "'," & vbCrLf
    sql = sql & "               " & mRimp.updNormal & ",'" & mRimp.mUpdateMessage(mRimp.updNormal) & "'," & vbCrLf
    sql = sql & "               " & mRimp.updWarnUpd & ",'" & mRimp.mUpdateMessage(mRimp.updWarnUpd) & "'," & vbCrLf
    sql = sql & "               " & mRimp.updResetCancel & ",'" & mRimp.mUpdateMessage(mRimp.updResetCancel) & "'," & vbCrLf
    sql = sql & "               '処理結果が特定できません。'" & vbCrLf
    sql = sql & "           )" & vbCrLf
    sql = sql & "      ELSE                             '例外 => 処理結果が特定できません。'" & vbCrLf
    sql = sql & " END as CIERRNM," & vbCrLf
#End If
    'sql = sql & " CIITKB," & vbCrLf
    sql = sql & " (SELECT ABKJNM " & vbCrLf
    sql = sql & "  FROM taItakushaMaster" & vbCrLf
    sql = sql & "  WHERE ABITKB = a.CIITKB" & vbCrLf
    sql = sql & " ) as ABKJNM," & vbCrLf    '//通常の外部結合でするとややこしいので...(tcHogoshaImport Table は全件出したい！)
    sql = sql & " CIKYCD," & vbCrLf
    sql = sql & " (SELECT MAX(BAKJNM) BAKJNM " & vbCrLf
    sql = sql & "  FROM tbKeiyakushaMaster " & vbCrLf
    sql = sql & "  WHERE BAITKB = a.CIITKB" & vbCrLf
    sql = sql & "    AND BAKYCD = a.CIKYCD" & vbCrLf
    '//契約者は現在有効分：契約期間＆振替期間
'//2012/08/09 契約期間を復活：古い氏名が出てしまうバグ対応
'    sql = sql & "    AND TO_CHAR(SYSDATE,'yyyymmdd') BETWEEN BAKYST AND BAKYED" & vbCrLf
'    sql = sql & "    AND TO_CHAR(SYSDATE,'yyyymmdd') BETWEEN BAFKST AND BAFKED" & vbCrLf
    sql = sql & "     and basqno in(" & vbCrLf
    sql = sql & "       select max(basqno) from tbKeiyakushaMaster " & vbCrLf
    sql = sql & "       WHERE BAITKB = a.CIITKB" & vbCrLf
    sql = sql & "         AND BAKYCD = a.CIKYCD" & vbCrLf
    sql = sql & "   )"
    sql = sql & " ) as BAKJNM," & vbCrLf    '//通常の外部結合でするとややこしいので...(tcHogoshaImport Table は全件出したい！)
    sql = sql & " CIKSCD," & vbCrLf
    sql = sql & " CIHGCD," & vbCrLf
    sql = sql & " CIKJNM," & vbCrLf
    sql = sql & " CIKNNM," & vbCrLf
    sql = sql & " CISTNM," & vbCrLf
    sql = sql & " DECODE(CISKGK,NULL,'',TO_CHAR(CISKGK,'99,999,999')) as CISKGK," & vbCrLf
    sql = sql & " DECODE(CIKKBN," & eBankKubun.KinnyuuKikan & ",'民間'," & eBankKubun.YuubinKyoku & ",'郵便局',NULL)     as CIKKBN," & vbCrLf
    sql = sql & " CIBANK," & vbCrLf
    sql = sql & " (SELECT DAKJNM" & vbCrLf
    sql = sql & "  FROM tdBankMaster" & vbCrLf
    sql = sql & "  WHERE DABANK = a.CIBANK" & vbCrLf
    sql = sql & "    AND DARKBN = '0'"
    sql = sql & "    AND DASITN = '000'"
    sql = sql & "    AND DASQNO = ':'"      '//これが現在有効
    sql = sql & " ) as DABKNM,"                '//通常の外部結合でするとややこしいので...(tcHogoshaImport Table は全件出したい！)
    sql = sql & " CIBKNM," & vbCrLf
    sql = sql & " CISITN," & vbCrLf
    sql = sql & " (SELECT DAKJNM" & vbCrLf
    sql = sql & "  FROM tdBankMaster" & vbCrLf
    sql = sql & "  WHERE DABANK = a.CIBANK" & vbCrLf
    sql = sql & "    AND DASITN = a.CISITN"
    sql = sql & "    AND DARKBN = '1'"
    sql = sql & "    AND DASQNO = 'ｱ'"      '//これが現在有効
    sql = sql & " ) as DASTNM,"                '//通常の外部結合でするとややこしいので...(tcHogoshaImport Table は全件出したい！)
    sql = sql & " CISINM," & vbCrLf
    sql = sql & " DECODE(CIKKBN," & eBankKubun.KinnyuuKikan & ",DECODE(CIKZSB,'1','普通','2','当座',CIKZSB),NULL) as CIKZSB," & vbCrLf
    sql = sql & " CIKZNO," & vbCrLf
    sql = sql & " CIYBTK," & vbCrLf
    sql = sql & " CIYBTN," & vbCrLf
    sql = sql & " CIKZNM," & vbCrLf
    sql = sql & " CIMUPD," & vbCrLf     '//2006/04/04 マスタ反映ＯＫフラグ項目追加
    sql = sql & " TO_CHAR(CIINDT,'yyyy/mm/dd hh24:mi:ss') CIINDT," & vbCrLf
    If vErrColomns Then
        sql = sql & mRimp.StatusColumns("," & vbCrLf)
    End If
    sql = sql & " CISEQN " & vbCrLf
    '////////////////////////////////////////////////////////////////////
    '//これ以降のＳＱＬ (MainSQL) を修正画面で流用するので注意して変更のこと！！！
    mMainSQL = " FROM " & mRimp.TcHogoshaImport & " a" & vbCrLf
    mMainSQL = mMainSQL & " WHERE CIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss')" & vbCrLf
    '//2006/04/14 ORDER が思惑通りになっていなかった
    'mMainSQL = mMainSQL & " ORDER BY DECODE(CIERSR,-2, 1,-1,-12, 1,-11 ,CIERSR)"    '修正、エラー、警告、正常の順
    mMainSQL = mMainSQL & " ORDER BY DECODE(CIERSR,-2, -11, -1,-12, 1,-10 ,CIERSR)"    '修正、エラー、警告、正常の順
    '//以降のＯＲＤＥＲ句
    Select Case cboSort.ListIndex
    Case eSort.eImportSeq
        mMainSQL = mMainSQL & ",CIINDT,CISEQN" & vbCrLf
    Case eSort.eKeiyakusha
        mMainSQL = mMainSQL & ",CIITKB,CIKYCD,CIKSCD,CIHGCD,CISEQN" & vbCrLf
    Case eSort.eKinnyuKikan
        mMainSQL = mMainSQL & ",CIKKBN,CIBANK,CISITN,CIKZSB,CIKZNO,CIYBTK,CIYBTN,CISEQN" & vbCrLf
    Case Else
    End Select
    sql = sql & mMainSQL & ")"
    pMakeSQLReadData = sql
End Function

Private Sub pReadDataAndSetting()
    
    dbcImport.RecordSource = pMakeSQLReadData
    sprMeisai.VirtualMode = False   '//一旦仮想モード解除
    Call dbcImport.Refresh

#If True = VIRTUAL_MODE Then
    '//仮想モードにするとページが変わるとデータが入れ替わってしまうので注意！！！
    sprMeisai.VScrollSpecial = True
    sprMeisai.VScrollSpecialType = 0
    sprMeisai.VirtualMode = True    '//仮想モード再設定：行のリフレッシュ！
    '//2012/07/02 特定のデータに対して表示ができない？バグ？なので設定行をコメント化：SQLが悪かった？
    sprMeisai.VirtualMaxRows = dbcImport.Recordset.RecordCount
#Else
    sprMeisai.VScrollSpecial = True
    sprMeisai.VScrollSpecialType = 0
    sprMeisai.MaxRows = dbcImport.Recordset.RecordCount
#End If
    
    '//セル単位にエラー箇所をカラー表示
    Call pSpreadSetErrorStatus(True)
    '//ToolTip を有効にする為に強制的にフォーカスを移す
    'Call sprMeisai.SetFocus
'//2007/07/19 口座戻りの件数を表示
    Dim sql As String, dyn As OraDynaset
    sql = "select count(*) modori " & vbCrLf
    sql = sql & " from " & mRimp.TcHogoshaImport & " a," & vbCrLf
    sql = sql & "   (select " & vbCrLf
    sql = sql & "     distinct caitkb,cakycd,cakscd,cahgcd " & vbCrLf
    sql = sql & "     from tcHogoshaMaster" & vbCrLf
    sql = sql & "   ) b " & vbCrLf
    sql = sql & " where ciitkb = caitkb " & vbCrLf
    sql = sql & "   and cikycd = cakycd " & vbCrLf
    sql = sql & "   and cikscd = cakscd " & vbCrLf
    sql = sql & "   and cihgcd = cahgcd " & vbCrLf
    sql = sql & "   AND CIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss')"
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
    lblModoriCount.Caption = "【 口座戻り件数： " & Format(dyn.Fields("modori"), "#,0") & " 件 】"
    Call dyn.Close
    Set dyn = Nothing
End Sub

Private Function pCheckSubForm() As Boolean
    '//修正画面が表示されていたなら閉じてしまう！
    If Not gdFormSub Is Nothing Then
        '//効かない？
        'If gdFormSub.dbcImport.EditMode <> OracleConstantModule.ORADATA_EDITNONE Then
            If vbOK <> MsgBox("修正画面での現在編集中のデータは破棄されます." & vbCrLf & vbCrLf & "よろしいですか？", vbOKCancel + vbDefaultButton2 + vbInformation, mCaption) Then
                Exit Function
            End If
            'Call gdFormSub.dbcImport.UpdateControls   '//キャンセル
            Call gdFormSub.cmdEnd_Click
        'End If
        'Unload gdFormSub
        Set gdFormSub = Nothing
    End If
    pCheckSubForm = True
End Function
Private Sub cboImpDate_Click()
    If "" = Trim(cboImpDate.Text) Then
        '//有り得ない
        Exit Sub
    End If
    If False = pCheckSubForm Then
        Exit Sub
    End If
    Dim ms As New MouseClass
    Call ms.Start
    '//データ読み込み＆ Spread に設定反映
    Call pReadDataAndSetting
End Sub

Private Sub cboSort_Click()
    Call cboImpDate_Click
End Sub

Private Function pMoveTempRecords(vCondition As String, vMode As String) As Long
    Dim sql As String
    '//削除対象データを Temp にバックアップ
    sql = "INSERT INTO " & mRimp.TcHogoshaImport & "Temp" & vbCrLf
    sql = sql & " SELECT SYSDATE,'" & vMode & "',a.*"
    sql = sql & " FROM " & mRimp.TcHogoshaImport & " a " & vbCrLf
    sql = sql & " WHERE 1 = 1" & vbCrLf
    sql = sql & vCondition
    Call gdDBS.Database.ExecuteSQL(sql)
    
    sql = "DELETE " & mRimp.TcHogoshaImport & vbCrLf
    sql = sql & " WHERE 1 = 1" & vbCrLf
    sql = sql & vCondition
    pMoveTempRecords = gdDBS.Database.ExecuteSQL(sql)
End Function

Private Sub cmdDelete_Click()
    If False = pCheckSubForm Then
        Exit Sub
    End If
    If 0 = cboImpDate.ListCount Then
        Exit Sub
    ElseIf vbOK <> MsgBox("現在表示されているデータを破棄します." & vbCrLf & vbCrLf & _
                      "廃棄対象 = [" & cboImpDate.Text & "]" & vbCrLf & vbCrLf & _
                      "よろしいですか？", vbOKCancel + vbInformation, mCaption) Then
        Exit Sub
    End If
    If -1 <> pAbortButton(cmdDelete, cBtnDelete) Then
        Exit Sub
    End If
    cmdDelete.Caption = cBtnCancel
    '//コマンド・ボタン制御
    Call pLockedControl(False, cmdDelete)
    
    Dim ms As New MouseClass, recCnt As Long
    Call ms.Start
    
    Call gdDBS.AutoLogOut(mCaption, "[" & cboImpDate.Text & "] の廃棄が開始されました。")
    
    On Error GoTo cmdDelete_ClickErr:
    Call gdDBS.Database.BeginTrans
    
    '//マスタ反映時にも同じ事をするので共通化
    recCnt = pMoveTempRecords(" AND CIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss')", gcFurikaeImportToDelete)
    If recCnt < 0 Then
        GoTo cmdDelete_ClickErr:
    End If
    
    Call gdDBS.Database.CommitTrans
    
    Set ms = Nothing
    Call MsgBox("廃棄対象 = [" & cboImpDate.Text & "]" & vbCrLf & vbCrLf & _
                recCnt & " 件が廃棄されました.", vbOKOnly + vbInformation, mCaption)
    
    '//ステータス行の整列・調整
    stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "廃棄完了"
    Call gdDBS.AutoLogOut(mCaption, "[" & cboImpDate.Text & "] の " & recCnt & " 件の廃棄が完了しました。")
    
    Call pMakeComboBox
    '//ボタンを戻す
    cmdDelete.Caption = cBtnDelete
    '//コマンド・ボタン制御
    Call pLockedControl(True)
    Exit Sub
cmdDelete_ClickErr:
    Call gdDBS.Database.Rollback
    If err Then
        Dim errCode As Integer, errMsg As String
        If gdDBS.Database.LastServerErr Then
            errCode = gdDBS.Database.LastServerErr
            errMsg = gdDBS.Database.LastServerErrText
        Else
            errCode = err
            errMsg = Error
        End If
        fraProgressBar.Visible = False
        '//ステータス行の整列・調整
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "廃棄エラー(" & errCode & ")"
        Call gdDBS.AutoLogOut(mCaption, "エラーが発生したため廃棄は中止されました。(Error=" & errMsg & ")")
        Call MsgBox("廃棄対象 = [" & cboImpDate.Text & "]" & vbCrLf & _
                    "はエラーが発生したため廃棄は中止されました。" & vbCrLf & errMsg, _
                vbOKOnly + vbCritical, mCaption)
    Else
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "廃棄中断"
        Call gdDBS.AutoLogOut(mCaption, "[" & cboImpDate.Text & "] の廃棄は中止されました。")
    End If
    '//ボタンを戻す
    cmdDelete.Caption = cBtnDelete
    '//コマンド・ボタン制御
    Call pLockedControl(True)
End Sub

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Function pAbortButton(vButton As CommandButton, vCaption As String) As Integer
    pAbortButton = -1   '// -1 = 処理開始
    mAbort = False
    If vButton.Caption <> cBtnCancel Then
        Exit Function
    End If
    pAbortButton = MsgBox(Left(vCaption, InStr(vCaption, "(") - 1) & "を中止しますか？", vbInformation + vbOKCancel, mCaption)
    If vbOK <> pAbortButton Then
        Exit Function   '//中止をやめた！
    End If
    vButton.Caption = vCaption
    mAbort = True
End Function

Private Sub cmdErrList_Click()
    If False = pCheckSubForm Then
        Exit Sub
    End If
    Dim reg As New RegistryClass
    Dim sql As String
    Load rptFurikaeReqImport
    With rptFurikaeReqImport
        .lblSort.Caption = "表示順： " & cboSort.Text
        '.mTotalCnt = dbcImport.Recordset.RecordCount
        .documentName = mCaption
        .adoData.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=" & reg.DbPassword & _
                                    ";Persist Security Info=True;User ID=" & reg.DbUserName & _
                                                           ";Data Source=" & reg.DbDatabaseName
        sql = pMakeSQLReadData(True)
        '//エラーデータは印刷で出力しない
        sql = sql & " WHERE CIERROR <> " & mRimp.errNormal
        .adoData.Source = sql
        'Call .adoData.Refresh
        Call .Show
    End With
End Sub

Private Sub cmdImport_Click()
    If False = pCheckSubForm Then
        Exit Sub
    End If
    '//ボタンのコントロール
    If -1 <> pAbortButton(cmdImport, cBtnImport) Then
        Exit Sub
    End If
    cmdImport.Caption = cBtnCancel
    '//コマンド・ボタン制御
    Call pLockedControl(False, cmdImport)

    Dim file As New FileClass
    
    dlgFile.DialogTitle = "ファイルを開く(" & mCaption & ")"
    dlgFile.FileName = mReg.InputFileName(mCaption)
    If IsEmpty(file.OpenDialog(dlgFile)) Then
        GoTo cmdImport_ClickAbort:
        Exit Sub
    End If
    '//振込依頼書データをインポート
    Dim Hogosha As tpHogoshaImport
    Dim fp As Integer
    Dim ms As New MouseClass
    Call ms.Start
    
    fp = FreeFile
    Open dlgFile.FileName For Random Access Read As #fp Len = Len(Hogosha)
    fraProgressBar.Visible = True
    pgrProgressBar.Max = LOF(fp) / Len(Hogosha)
    '//ファイルサイズが違う場合の警告メッセージ
    If pgrProgressBar.Max <> Int(pgrProgressBar.Max) Then
        If (LOF(fp) - 1) / Len(Hogosha) <> Int((LOF(fp) - 1) / Len(Hogosha)) Then
            '/処理続行するとＤＢがおかしくなるので中止する
            Close #fp
            Call gdDBS.MsgBox("指定されたファイル(" & dlgFile.FileName & ")が異常です。" & vbCrLf & vbCrLf & "処理を続行出来ません。", vbCritical + vbOKOnly, mCaption)
            GoTo cmdImport_ClickAbort
            Exit Sub
        End If
    End If

    On Error GoTo cmdImport_ClickError
        
    Call gdDBS.AutoLogOut(mCaption, "取込処理が開始されました。")
    
#If ORA_DEBUG = 1 Then
    Dim sql As String, insDate As String, dyn As OraDynaset
#Else
    Dim sql As String, insDate As String, dyn As Object
#End If
    Dim insCnt As Long, updCnt As Long, recCnt As Long
    
    insDate = gdDBS.sysDate()
    
    Call gdDBS.Database.BeginTrans
    '///////////////////////////////////////////////
    '//シーケンスを１番からにリセット
    sql = "declare begin ResetSequence('sqImportSeq',1); end;"
    Call gdDBS.Database.ExecuteSQL(sql)
    
    Do While Loc(fp) < LOF(fp) / Len(Hogosha)
        DoEvents
        If mAbort Then
            GoTo cmdImport_ClickError
        End If
        Get #fp, , Hogosha
        recCnt = Loc(fp)
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = _
            "残り" & Right(String(7, " ") & Format(pgrProgressBar.Max - Loc(fp), "#,##0"), 7) & " 件"
        pgrProgressBar.Value = IIf(Loc(fp) <= pgrProgressBar.Max, Loc(fp), pgrProgressBar.Max)
        
#If DATA_DUPLICATE = True Then  '//振替依頼書は重複をチェック
'''''''''''''        '//2006/03/24 後から持ち込んだデータ(持込日が大きい)が有効にする
'''''''''''''        sql = "SELECT CIMCDT"
'''''''''''''        sql = sql & " FROM " & mRimp.TcHogoshaImport & " a "
'''''''''''''        sql = sql & "WHERE CIINDT = " & "TO_DATE(" & gdDBS.ColumnDataSet(insDate,"D", vEnd:=True) & ",'yyyy-mm-dd hh24:mi:ss')" & vbCrLf  '//取込日
'''''''''''''        sql = sql & "  AND CIKYCD = " & gdDBS.ColumnDataSet(Hogosha.Keiyakusha, vEnd:=True) & vbCrLf    '//契約者番号
'''''''''''''        sql = sql & "  AND CIKSCD = " & gdDBS.ColumnDataSet(Hogosha.Kyoshittsu, vEnd:=True) & vbCrLf    '//教室番号
'''''''''''''        sql = sql & "  AND CIHGCD = " & gdDBS.ColumnDataSet(Hogosha.HogoshaNo, vEnd:=True) & vbCrLf     '//保護者番号
'''''''''''''#If ORA_DEBUG = 1 Then
'''''''''''''        Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
'''''''''''''#Else
'''''''''''''        Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
'''''''''''''#End If
'''''''''''''        If Not dyn.EOF Then
'''''''''''''            '//持込日が後日なら更新
'''''''''''''            If Val(dyn.Fields("CIMCDT")) < Val(Hogosha.MochikomiBi) Then
'''''''''''''                '//更新を試みる：同一テキスト内に同じデータがある？
'''''''''''''                sql = "UPDATE " & mRimp.TcHogoshaImport & " SET " & vbCrLf
'''''''''''''                '//委託者は契約者から割り出すので不要！
'''''''''''''                'sql = sql & "CIITKB = (SELECT ABITKB FROM taItakushaMaster WHERE ABKYTP = '" & Left(Hogosha.Keiyakusha, 1) & "')," & vbcrlf  '//委託者区分
'''''''''''''                sql = sql & "CIKJNM = " & gdDBS.ColumnDataSet(Hogosha.HogoshaKanji) & vbCrLf    '//保護者名_漢字
'''''''''''''                sql = sql & "CIKNNM = " & gdDBS.ColumnDataSet(Hogosha.HogoshaKana) & vbCrLf     '//保護者名_カナ
'''''''''''''                sql = sql & "CISTNM = " & gdDBS.ColumnDataSet(Hogosha.SeitoShimei) & vbCrLf     '//生徒氏名
'''''''''''''                sql = sql & "CISKGK = " & gdDBS.ColumnDataSet(Hogosha.FurikaeGaku, "L") & vbCrLf '//振替金額
'''''''''''''                sql = sql & "CIBKNM = " & gdDBS.ColumnDataSet(Hogosha.BankName) & vbCrLf        '//取込銀行名
'''''''''''''                sql = sql & "CISINM = " & gdDBS.ColumnDataSet(Hogosha.ShitenName) & vbCrLf      '//取込支店名
'''''''''''''                If "" = Trim(Hogosha.TuuchoKigou) _
'''''''''''''                And "" = Trim(Hogosha.TuuchoBango) Then         '//郵便局情報記入なし
'''''''''''''                    sql = sql & "CIKKBN = " & gdDBS.ColumnDataSet(eBankKubun.KinnyuuKikan, "I") & vbCrLf               '//取引金融機関区分
'''''''''''''                Else
'''''''''''''                    sql = sql & "CIKKBN = " & gdDBS.ColumnDataSet(eBankKubun.YuubinKyoku, "I") & vbCrLf               '//取引金融機関区分
'''''''''''''                End If
'''''''''''''                sql = sql & "CIBANK = " & gdDBS.ColumnDataSet(Hogosha.BankCode) & vbCrLf        '//取引銀行
'''''''''''''                sql = sql & "CISITN = " & gdDBS.ColumnDataSet(Hogosha.ShitenCode) & vbCrLf      '//取引支店
'''''''''''''                sql = sql & "CIKZSB = " & gdDBS.ColumnDataSet(Hogosha.YokinShumoku) & vbCrLf    '//口座種別
'''''''''''''                sql = sql & "CIKZNO = " & gdDBS.ColumnDataSet(Hogosha.KouzaBango) & vbCrLf      '//口座番号
'''''''''''''                sql = sql & "CIYBTK = " & gdDBS.ColumnDataSet(Hogosha.TuuchoKigou) & vbCrLf     '//通帳記号
'''''''''''''                sql = sql & "CIYBTN = " & gdDBS.ColumnDataSet(Hogosha.TuuchoBango) & vbCrLf     '//通帳番号
'''''''''''''                sql = sql & "CIKZNM = " & gdDBS.ColumnDataSet(Hogosha.HogoshaKana) & vbCrLf     '//口座名義人_カナ
'''''''''''''                sql = sql & "CIERROR = " & gdDBS.ColumnDataSet(mRimp.errImport) & vbCrLf
'''''''''''''                sql = sql & "CIERSR  = " & gdDBS.ColumnDataSet(mRimp.errImport) & vbCrLf
'''''''''''''                sql = sql & "CIMCDT = " & gdDBS.ColumnDataSet(Hogosha.MochikomiBi, "L") & vbCrLf    '//持込日
'''''''''''''                sql = sql & "CIUPDT = SYSDATE " & vbCrLf                                        '//更新日
'''''''''''''                sql = sql & "WHERE CIINDT = " & "TO_DATE(" & gdDBS.ColumnDataSet(insDate,"D", vEnd:=True) & ",'yyyy-mm-dd hh24:mi:ss')" & vbCrLf  '//取込日
'''''''''''''                sql = sql & "  AND CIKYCD = " & gdDBS.ColumnDataSet(Hogosha.Keiyakusha, vEnd:=True) & vbCrLf    '//契約者番号
'''''''''''''                sql = sql & "  AND CIKSCD = " & gdDBS.ColumnDataSet(Hogosha.Kyoshittsu, vEnd:=True) & vbCrLf    '//教室番号
'''''''''''''                sql = sql & "  AND CIHGCD = " & gdDBS.ColumnDataSet(Hogosha.HogoshaNo, vEnd:=True) & vbCrLf     '//保護者番号
'''''''''''''                Call gdDBS.Database.ExecuteSQL(sql)
'''''''''''''                updCnt = updCnt + 1&
'''''''''''''            End If
'''''''''''''        Else
#End If     '//#If DATA_DUPLICATE = True Then  '//振替依頼書は重複をチェック
            
            insCnt = insCnt + 1&
            '//更新できなかったので挿入を試みる
            '//データをテーブルに挿入
            sql = "INSERT INTO " & mRimp.TcHogoshaImport & "(" & vbCrLf
            sql = sql & "CIINDT,"   '//取込日
            sql = sql & "CISEQN,"   '//取込SEQNO
            sql = sql & "CIITKB,"   '//委託者区分
            sql = sql & "CIKYCD,"   '//契約者番号
            sql = sql & "CIKSCD,"   '//教室番号
            sql = sql & "CIHGCD,"   '//保護者番号
            sql = sql & "CIKJNM,"   '//保護者名_漢字
            sql = sql & "CIKNNM,"   '//保護者名_カナ
            sql = sql & "CISTNM,"   '//生徒氏名
            sql = sql & "CISKGK,"   '//振替金額
            sql = sql & "CIBKNM,"   '//取込銀行名
            sql = sql & "CISINM,"   '//取込支店名
            sql = sql & "CIKKBN,"   '//取引金融機関区分
            sql = sql & "CIBANK,"   '//取引銀行
            sql = sql & "CISITN,"   '//取引支店
            sql = sql & "CIKZSB,"   '//口座種別
            sql = sql & "CIKZNO,"   '//口座番号
            sql = sql & "CIYBTK,"   '//通帳記号
            sql = sql & "CIYBTN,"   '//通帳番号
            sql = sql & "CIKZNM,"   '//口座名義人_カナ
            sql = sql & "CIERROR,"
            sql = sql & "CIERSR,"
            sql = sql & "CIMCDT,"   '//持込日   2006/03/24 ADD
            sql = sql & "CIUSID,"   '//更新者
            sql = sql & "CIUPDT,"   '//更新日
            sql = sql & "CIOKFG " & vbCrLf  '//取込ＯＫフラグ
            sql = sql & ")VALUES(" & vbCrLf
            sql = sql & "TO_DATE(" & gdDBS.ColumnDataSet(insDate, "D", vEnd:=True) & ",'yyyy-mm-dd hh24:mi:ss'),"
            sql = sql & "sqImportSeq.NEXTVAL,"
            sql = sql & " (SELECT ABITKB FROM taItakushaMaster WHERE ABKYTP = '" & Left(Hogosha.Keiyakusha, 1) & "'),"
            sql = sql & gdDBS.ColumnDataSet(Hogosha.Keiyakusha)
            sql = sql & gdDBS.ColumnDataSet(Hogosha.Kyoshittsu)
            sql = sql & gdDBS.ColumnDataSet(Hogosha.HogoshaNo)
            sql = sql & gdDBS.ColumnDataSet(Hogosha.HogoshaKanji)
            sql = sql & gdDBS.ColumnDataSet(Hogosha.HogoshaKana)
            sql = sql & gdDBS.ColumnDataSet(Hogosha.SeitoShimei)
'//2006/04/26 金額なので NULL では無く 「０」を代入する
            sql = sql & gdDBS.ColumnDataSet(gdDBS.Nz(Hogosha.FurikaeGaku, 0), "L") & vbCrLf
            sql = sql & gdDBS.ColumnDataSet(Hogosha.BankName)
            sql = sql & gdDBS.ColumnDataSet(Hogosha.ShitenName)
            If "" <> Trim(Hogosha.BankCode) _
            And "" <> Trim(Hogosha.ShitenCode) Then     '//民間金融機関コード 記入あり
                sql = sql & gdDBS.ColumnDataSet(eBankKubun.KinnyuuKikan, "I")   '//民間金融機関
            ElseIf "" <> Trim(Hogosha.TuuchoKigou) _
                And "" <> Trim(Hogosha.TuuchoBango) Then '//郵便局情報 記入あり
                sql = sql & gdDBS.ColumnDataSet(eBankKubun.YuubinKyoku, "I")   '//郵便局
            Else
                sql = sql & "NULL,"   '//金融機関区分＝NULL
            End If
            sql = sql & gdDBS.ColumnDataSet(Hogosha.BankCode)
            sql = sql & gdDBS.ColumnDataSet(Hogosha.ShitenCode)
            sql = sql & gdDBS.ColumnDataSet(Val(Hogosha.YokinShumoku))  '//預金種目＝０の対応
            sql = sql & gdDBS.ColumnDataSet(Hogosha.KouzaBango)
            sql = sql & gdDBS.ColumnDataSet(Hogosha.TuuchoKigou)
            sql = sql & gdDBS.ColumnDataSet(Hogosha.TuuchoBango)
            sql = sql & gdDBS.ColumnDataSet(Hogosha.HogoshaKana)
            sql = sql & gdDBS.ColumnDataSet(mRimp.errImport) & vbCrLf
            sql = sql & gdDBS.ColumnDataSet(mRimp.errImport) & vbCrLf
            sql = sql & gdDBS.ColumnDataSet(Hogosha.MochikomiBi, "L") & vbCrLf    '//持込日
            sql = sql & gdDBS.ColumnDataSet(gdDBS.LoginUserName)
            sql = sql & "SYSDATE,"
            sql = sql & gdDBS.ColumnDataSet(mRimp.updInvalid, "I", vEnd:=True)
            sql = sql & ")"
            Call gdDBS.Database.ExecuteSQL(sql)
#If DATA_DUPLICATE = True Then  '//振替依頼書は重複をチェック
''''''''''''        End If
''''''''''''        Call dyn.Close
''''''''''''        Set dyn = Nothing
#End If
    Loop
    '//取込結果の最終編集
    '//2006/04/26 保護者番号、口座番号、通帳記号、通帳番号の前ゼロ補間追加
    sql = "UPDATE " & mRimp.TcHogoshaImport & " a SET "
    sql = sql & "CIKSCD = DECODE(CIKSCD,NULL,NULL,LPAD(CIKSCD,3,'0'))," & vbCrLf    '//教室番号：       入力が１桁なので入力が有る場合のみ３桁に編集
    sql = sql & "CIHGCD = DECODE(CIHGCD,NULL,NULL,LPAD(CIHGCD,4,'0'))," & vbCrLf    '//保護者：
    sql = sql & "CIBANK = DECODE(CIBANK,NULL,NULL,LPAD(CIBANK,4,'0'))," & vbCrLf    '//金融機関コード： 入力が４桁だが入力が有る場合のみ４桁に編集
    sql = sql & "CISITN = DECODE(CISITN,NULL,NULL,LPAD(CISITN,3,'0'))," & vbCrLf    '//支店コード：     入力が３桁だが入力が有る場合のみ３桁に編集
    sql = sql & "CIKZNO = DECODE(CIKZNO,NULL,NULL,LPAD(CIKZNO,7,'0'))," & vbCrLf    '//口座番号 ７桁
    sql = sql & "CIYBTK = DECODE(CIYBTK,NULL,NULL,LPAD(CIYBTK," & mRimp.YubinKigouLength & ",'0'))," & vbCrLf     '//通帳記号 ３桁
    sql = sql & "CIYBTN = DECODE(CIYBTN,NULL,NULL,LPAD(CIYBTN," & mRimp.YubinBangoLength & ",'0')) " & vbCrLf     '//通帳番号 ８桁
    sql = sql & " WHERE CIINDT = TO_DATE(" & gdDBS.ColumnDataSet(insDate, "D", vEnd:=True) & ",'yyyy-mm-dd hh24:mi:ss')"
    Call gdDBS.Database.ExecuteSQL(sql)
    Close #fp
    '//ステータス行の整列・調整
    stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "取込完了(" & recCnt & "件)"
    pgrProgressBar.Value = pgrProgressBar.Max
    '//振込依頼書データの位置をレジストリに保管
    mReg.InputFileName(mCaption) = dlgFile.FileName
    '//取込データのバックアップ
    Call gBackupTextData(dlgFile.FileName)

    Call gdDBS.Database.CommitTrans
    
    Call gdDBS.AutoLogOut(mCaption, "取込日時=[" & insDate & "]で " & recCnt & " 件（追加=" & insCnt & " / 重複=" & updCnt & "）のデータが取り込まれました。")
    
    '//取込結果をコンボボックスにセット
    Call pMakeComboBox

cmdImport_ClickAbort:
    '//すべての定義をリセット
    Set file = Nothing
    Set ms = Nothing
    cmdImport.Caption = cBtnImport
    fraProgressBar.Visible = False
    Call pLockedControl(True)
    Exit Sub
cmdImport_ClickError:
    '//ステータス行の整列・調整
    'cmdImport.Caption = cBtnImport
    Call gdDBS.Database.Rollback
    Call gdDBS.ErrorCheck       '//エラートラップ
    If err Then
        Dim errCode As Integer, errMsg As String
        If gdDBS.Database.LastServerErr Then
            errCode = gdDBS.Database.LastServerErr
            errMsg = gdDBS.Database.LastServerErrText
        Else
            errCode = err
            errMsg = Error
        End If
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "取込エラー(" & errCode & ")"
        Call gdDBS.AutoLogOut(mCaption, recCnt & "件目でエラーが発生したため取込処理は中止されました。(Error=" & errMsg & ")")
    Else
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "取込中断"
        Call gdDBS.AutoLogOut(mCaption, "取込処理は中止されました。")
    End If
    Call pLockedControl(True)
    GoTo cmdImport_ClickAbort:
End Sub

'//金融機関＆支店名のマッチング用
Private Function pCompare(vElm1 As Variant, vElm2 As Variant, Optional vCutString As Variant = "") As Boolean
    '// vElm1 と vElm2 が同じであれば True
    '//Replace()以外でしようとするとややこしいので！！！止め。
    pCompare = Replace(vElm1, vCutString, "") = Replace(vElm2, vCutString, "")
End Function

Private Function pErrorCount() As Integer
    On Error GoTo pErrorCountError
    pErrorCount = UBound(mErrStts)
    Exit Function
pErrorCountError:
    pErrorCount = -1
End Function

Private Sub pSetErrorStatus(vField As Variant, vError As Integer, Optional vMsg As String = "")
    On Error GoTo SetErrorStatusError:
    Dim ix As Integer
    For ix = LBound(mErrStts) To UBound(mErrStts)
        If UCase(vField) = UCase(mErrStts(ix).Field) Then
            If vError < mErrStts(ix).Error Then
                GoTo SetErrorStatusSet:
            End If
            Exit Sub
        End If
    Next ix
    ix = UBound(mErrStts) + 1
    ReDim Preserve mErrStts(ix) As tpErrorStatus
SetErrorStatusSet:
    mErrStts(ix).Field = UCase(vField)
    mErrStts(ix).Error = vError
    If "" <> vMsg Then
        mErrStts(ix).Message = vMsg
    End If
    Exit Sub
SetErrorStatusError:
    ix = 0
    ReDim Preserve mErrStts(0 To 0) As tpErrorStatus
    GoTo SetErrorStatusSet:
End Sub

Private Function pProgressBarSet(ByRef rBlockStep As Integer, Optional ByRef rStepCnt As Long = -1) As Boolean
    DoEvents    '//イベント受付
    If mAbort Then
        pProgressBarSet = False     '//処理中断！
        Exit Function
    End If
    '//ステータス行の整列・調整
    If 0 <= rStepCnt Then
        If 0 = rStepCnt Then
            rBlockStep = rBlockStep - 1
        End If
        rStepCnt = rStepCnt + 1
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "残り(" & rBlockStep & ") - " & pgrProgressBar.Max - rStepCnt
        pgrProgressBar.Value = IIf(rStepCnt < pgrProgressBar.Max, rStepCnt, pgrProgressBar.Max)
    Else
        rBlockStep = rBlockStep - 1
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "残り(" & rBlockStep & ")"
        pgrProgressBar.Value = IIf(0 <= pgrProgressBar.Max - rBlockStep, pgrProgressBar.Max - rBlockStep, pgrProgressBar.Max)
    End If
    pProgressBarSet = True
#If BLOCK_CHECK = True Then           '//チェック時のブロックがいくつあるか？を表示：デバック時のみ
    If rStepCnt <= 1 Then
        mCheckBlocks = mCheckBlocks + 1
    End If
#End If
End Function

'/////////////////////////////////////////////////////////////////////////
'//個別に１件ずつ処理
Public Function gDataCheck(vImpDate As Variant, Optional vSeqNo As Long = -1) As Boolean
    Dim Block As Integer, sqlStep As Long
    Const cMaxBlock As Integer = 5
    Block = cMaxBlock
#If BLOCK_CHECK = True Then           '//チェック時のブロックがいくつあるか？を表示：デバック時のみ
    mCheckBlocks = 0
#End If
    
    '// WHERE 句には必ず付加
    Dim SameConditions As String
    SameConditions = " AND CIINDT = TO_DATE('" & vImpDate & "','yyyy/mm/dd hh24:mi:ss')" & vbCrLf
    'SameConditions = " AND CIOKFG NOT IN(" & mRimp.updInvalid & "," & mRimp.updWarnErr & ")" & vbCrLf
    'SameConditions = " AND CIERROR = " & mRimp.errNormal
    If -1 <> vSeqNo Then
        SameConditions = SameConditions & vbCrLf & " AND CISEQN = " & vSeqNo
    End If
    
    On Error GoTo gDataCheckError:
    
    Dim ms As New MouseClass
    Call ms.Start
    fraProgressBar.Visible = True
    
    Call gdDBS.AutoLogOut(mCaption, "[" & vImpDate & ":" & vSeqNo & "] のチェック処理が開始されました。")
    
    Call gdDBS.Database.BeginTrans          '//トランザクション開始

    '////////////////////////////////////////
    '//削除してチェックする文字を定義
    Dim BankCutName As Variant, ShitenCutName As Variant
    Dim updFlag As Integer, impName As String, mstName As String
    '//銀行名称
    BankCutName = Array("", "銀行", "信用金庫", "信用組合", _
                            "労働金庫", "協同組合", "農業協同組合", _
                            "漁業協同組合連合会")
    '//支店名称
    ShitenCutName = Array("", "支店", "出張所", "営業部", "支所")
    Dim sql As String, recCnt As Long, sysDate As String
    Dim ix As Integer, msg As String
#If ORA_DEBUG = 1 Then
    Dim dynM As OraDynaset, dynS As OraDynaset
#Else
    Dim dynM As Object, dynS As Object
#End If
    sysDate = gdDBS.sysDate("YYYYMMDD")
    '//////////////////////////////////////////////////
    '//エラー項目リセット
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
    sql = "UPDATE " & mRimp.TcHogoshaImport & " a SET " & vbCrLf
    sql = sql & mRimp.StatusColumns(" = " & mRimp.errNormal & "," & vbCrLf)
    '//手作業で警告データを「マスタ反映する」としているデータがあるので初期化しない
    '//2006/03/14 手修正した分はそのままにして「０」に置換え
    sql = sql & " CIOKFG = CASE WHEN CIOKFG >= " & mRimp.updWarnUpd & " THEN CIOKFG" & vbCrLf
    sql = sql & "               ELSE " & mRimp.updNormal & vbCrLf
    sql = sql & "          END,"
    sql = sql & " CIWMSG = NULL,"   '//ワーニングメッセージ
    sql = sql & " CIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
    sql = sql & " CIUPDT = SYSDATE" & vbCrLf
    sql = sql & " WHERE 1 = 1" & vbCrLf '//おまじない
    sql = sql & SameConditions & vbCrLf
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    '////////////////////////////////////////////
    '//振替依頼書を１件ずつ処理する
    sql = "SELECT a.* " & vbCrLf
    sql = sql & " FROM " & mRimp.TcHogoshaImport & " a " & vbCrLf
    sql = sql & " WHERE 1 = 1" & vbCrLf
    sql = sql & SameConditions & vbCrLf
    sql = sql & " ORDER BY CIKYCD,CIHGCD,CIKSCD"
    Set dynM = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
    If Not dynM.EOF Then
        pgrProgressBar.Max = dynM.RecordCount
    End If
    Do Until dynM.EOF
        '//////////////////////////////////////////////////
        '// DoEvents は pProgressBarSet() の中で実行されている
        If False = pProgressBarSet(Block, dynM.RowPosition - 1) Then
            GoTo gDataCheckError:
        End If
        '//結果を初期化
        Erase mErrStts
        '//////////////////////////////////////////
        '//委託者コードチェック:先頭１文字 ２＝公文、７＝公文エルアイエル
        sql = "SELECT ABITKB " & vbCrLf
        sql = sql & " FROM taItakushaMaster   a " & vbCrLf
        sql = sql & " WHERE ABKYTP = " & gdDBS.ColumnDataSet(Left(dynM.Fields("CIKYCD"), 1), vEnd:=True) & vbCrLf
        Set dynS = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
        If dynS.EOF Then
            Call pSetErrorStatus("CIITKBE", mRimp.errInvalid, "委託者が間違っています.")
        End If
        Call dynS.Close
        Set dynS = Nothing
        '//////////////////////////////////////////
        '//契約者コードチェック
        sql = "SELECT BAKYED,BAKYFG " & vbCrLf
        sql = sql & " FROM tbKeiyakushaMaster a " & vbCrLf
        sql = sql & " WHERE (BAITKB,BAKYCD,BASQNO) IN(" & vbCrLf
        sql = sql & "       SELECT BAITKB,BAKYCD,MAX(BASQNO) " & vbCrLf
        sql = sql & "       FROM tbKeiyakushaMaster a" & vbCrLf
        sql = sql & "       WHERE BAITKB = " & gdDBS.ColumnDataSet(dynM.Fields("CIITKB"), vEnd:=True) & vbCrLf
        sql = sql & "         AND BAKYCD = " & gdDBS.ColumnDataSet(dynM.Fields("CIKYCD"), vEnd:=True) & vbCrLf
        sql = sql & "       GROUP BY BAITKB,BAKYCD" & vbCrLf
        sql = sql & "     )" & vbCrLf
        Set dynS = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
        If dynS.EOF Then
            Call pSetErrorStatus("CIKYCDE", mRimp.errInvalid, "契約者が存在しません.")
        ElseIf dynS.Fields("BAKYED") < sysDate Or 0 <> Val(gdDBS.Nz(dynS.Fields("BAKYFG"))) Then
            Call pSetErrorStatus("CIKYCDE", mRimp.errInvalid, "契約者が解約状態です.")
        End If
        Call dynS.Close
        Set dynS = Nothing
        '//################################################
        '//2008/10/14 中途ブランク検知
        If 0 <> InStr(dynM.Fields("CIHGCD"), " ") Then
            Call pSetErrorStatus("CIHGCDE", mRimp.errInvalid, "保護者番号にブランクがあります.")
        End If
        If dynM.Fields("CIKKBN") = eBankKubun.KinnyuuKikan Then
            If 0 <> InStr(dynM.Fields("CIBANK"), " ") Then
                Call pSetErrorStatus("CIBANKE", mRimp.errInvalid, "金融機関にブランクがあります.")
            End If
            If 0 <> InStr(dynM.Fields("CISITN"), " ") Then
                Call pSetErrorStatus("CISITNE", mRimp.errInvalid, "支店にブランクがあります.")
            End If
            If 0 <> InStr(dynM.Fields("CIKZNO"), " ") Then
                Call pSetErrorStatus("CIKZNOE", mRimp.errInvalid, "口座番号にブランクがあります.")
            End If
        ElseIf dynM.Fields("CIKKBN") = eBankKubun.YuubinKyoku Then
            If 0 <> InStr(dynM.Fields("CIYBTK"), " ") Then
                Call pSetErrorStatus("CIYBTKE", mRimp.errInvalid, "通帳記号にブランクがあります.")
            End If
            If 0 <> InStr(dynM.Fields("CIYBTN"), " ") Then
                Call pSetErrorStatus("CIYBTNE", mRimp.errInvalid, "通帳番号にブランクがあります.")
            End If
        End If
        '//2008/10/14 中途ブランク検知
        '//################################################
        '//教室番号チェック
        If IsNull(dynM.Fields("CIKSCD")) Then
            Call pSetErrorStatus("CIKSCDE", mRimp.errInvalid, "教室番号が未入力です.")
        End If
        '//////////////////////////////////////////
        '//保護者番号チェック
        If IsNull(dynM.Fields("CIHGCD")) Then
            Call pSetErrorStatus("CIHGCDE", mRimp.errInvalid, "保護者番号が未入力です.")
        End If
        '//////////////////////////////////////////
        '//保護者名(漢字)チェック
        If IsNull(dynM.Fields("CIKJNM")) Then
            Call pSetErrorStatus("CIKJNME", mRimp.errInvalid, "保護者名(漢字)が未入力です.")
        End If
        '//////////////////////////////////////////
        '//保護者名(カナ)チェック
        If IsNull(dynM.Fields("CIKNNM")) Then
            Call pSetErrorStatus("CIKNNME", mRimp.errInvalid, "保護者名(カナ)が未入力です.")
        End If
        '//////////////////////////////////////////
        '//過去/今回 振替依頼書・取込データとのチェック
        sql = "SELECT MAX(DupCode) DUPCODE FROM(" & vbCrLf
        sql = sql & " SELECT " & gdDBS.ColumnDataSet("過去", vEnd:=True) & " DupCode " & vbCrLf
        sql = sql & " FROM " & mRimp.TcHogoshaImport & " a " & vbCrLf
        sql = sql & " WHERE CIINDT <>TO_DATE('" & vImpDate & "','yyyy/mm/dd hh24:mi:ss')" & vbCrLf
        sql = sql & "   AND CIITKB = " & gdDBS.ColumnDataSet(dynM.Fields("CIITKB"), vEnd:=True) & vbCrLf
        sql = sql & "   AND CIKYCD = " & gdDBS.ColumnDataSet(dynM.Fields("CIKYCD"), vEnd:=True) & vbCrLf
        sql = sql & "   AND CIKSCD = " & gdDBS.ColumnDataSet(dynM.Fields("CIKSCD"), vEnd:=True) & vbCrLf
        sql = sql & "   AND CIHGCD = " & gdDBS.ColumnDataSet(dynM.Fields("CIHGCD"), vEnd:=True) & vbCrLf
        sql = sql & " UNION " & vbCrLf
        sql = sql & " SELECT " & gdDBS.ColumnDataSet("今回", vEnd:=True) & " DupCode " & vbCrLf
        sql = sql & " FROM " & mRimp.TcHogoshaImport & " a " & vbCrLf
        sql = sql & " WHERE CIINDT = TO_DATE('" & vImpDate & "','yyyy/mm/dd hh24:mi:ss')" & vbCrLf
        '//自分自身以外
        sql = sql & "   AND CISEQN <>" & gdDBS.ColumnDataSet(dynM.Fields("CISEQN"), "I", vEnd:=True) & vbCrLf
        sql = sql & "   AND CIITKB = " & gdDBS.ColumnDataSet(dynM.Fields("CIITKB"), vEnd:=True) & vbCrLf
        sql = sql & "   AND CIKYCD = " & gdDBS.ColumnDataSet(dynM.Fields("CIKYCD"), vEnd:=True) & vbCrLf
        sql = sql & "   AND CIKSCD = " & gdDBS.ColumnDataSet(dynM.Fields("CIKSCD"), vEnd:=True) & vbCrLf
        sql = sql & "   AND CIHGCD = " & gdDBS.ColumnDataSet(dynM.Fields("CIHGCD"), vEnd:=True) & vbCrLf
        sql = sql & ")"
        Set dynS = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
        '//MAX() でしているので必ず存在する
        If Not IsNull(dynS.Fields("DupCode")) Then
            Call pSetErrorStatus("CIHGCDE", mRimp.errWarning, dynS.Fields("DupCode") & "の取込データに存在します.")
        End If
        Call dynS.Close
        Set dynS = Nothing
        '//////////////////////////////////////////
        '//保護者マスタとのチェック
        sql = "SELECT a.* " & vbCrLf
        sql = sql & " FROM tcHogoshaMaster a " & vbCrLf
        sql = sql & " WHERE (CAITKB,CAKYCD,CAKSCD,CAHGCD,CASQNO) IN(" & vbCrLf
        sql = sql & "       SELECT CAITKB,CAKYCD,CAKSCD,CAHGCD,MAX(CASQNO) " & vbCrLf
        sql = sql & "       FROM tcHogoshaMaster a" & vbCrLf
        sql = sql & "       WHERE CAITKB = " & gdDBS.ColumnDataSet(dynM.Fields("CIITKB"), vEnd:=True) & vbCrLf
        sql = sql & "         AND CAKYCD = " & gdDBS.ColumnDataSet(dynM.Fields("CIKYCD"), vEnd:=True) & vbCrLf
        sql = sql & "         AND CAKSCD = " & gdDBS.ColumnDataSet(dynM.Fields("CIKSCD"), vEnd:=True) & vbCrLf
        sql = sql & "         AND CAHGCD = " & gdDBS.ColumnDataSet(dynM.Fields("CIHGCD"), vEnd:=True) & vbCrLf
        sql = sql & "       GROUP BY CAITKB,CAKYCD,CAKSCD,CAHGCD" & vbCrLf
        sql = sql & "     )" & vbCrLf
        Set dynS = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
        '//////////////////////////////////////////
        '//データがある場合のみで可：無い筈は無い：ＳＴＡＲＴ
        If Not dynS.EOF Then
            If dynS.Fields("CAKYED") < sysDate Or 0 <> Val(gdDBS.Nz(dynS.Fields("CAKYFG"))) Then
                Call pSetErrorStatus("CIHGCDE", mRimp.errWarning, "保護者マスタは解約状態です.")
            Else
                Call pSetErrorStatus("CIHGCDE", mRimp.errWarning, "保護者マスタに既に存在します.")
            End If
            '//////////////////////////////////////////
            '//保護者名(漢字)チェック
            If Replace(Replace(dynM.Fields("CIKJNM"), "　", ""), " ", "") _
            <> Replace(Replace(dynS.Fields("CAKJNM"), "　", ""), " ", "") Then
                Call pSetErrorStatus("CIKJNME", mRimp.errWarning, "保護者名(漢字)に相違があります.")
            End If
            '//////////////////////////////////////////
            '//保護者名(カナ)チェック
            '//2007/04/20 パンチに保護者カナ NULL 有りの為エラー
            If Not IsNull(dynM.Fields("CIKNNM")) Then
                If Replace(Replace(dynM.Fields("CIKNNM"), "　", ""), " ", "") _
                <> Replace(Replace(dynS.Fields("CAKNNM"), "　", ""), " ", "") Then
                    Call pSetErrorStatus("CIKNNME", mRimp.errWarning, "保護者名(カナ)に相違があります.")
                End If
            End If
            If dynM.Fields("CIKKBN") = eBankKubun.KinnyuuKikan Then
                '//////////////////////////////////////////
                '//金融機関チェック
                If dynM.Fields("CIBANK") <> dynS.Fields("CABANK") Then
                    Call pSetErrorStatus("CIBANKE", mRimp.errWarning, "金融機関に相違があります.")
                End If
                '//////////////////////////////////////////
                '//支店チェック
                If dynM.Fields("CISITN") <> dynS.Fields("CASITN") Then
                    Call pSetErrorStatus("CISITNE", mRimp.errWarning, "支店に相違があります.")
                End If
                '//////////////////////////////////////////
                '//預金種目チェック
                If dynM.Fields("CIKZSB") <> dynS.Fields("CAKZSB") Then
                    Call pSetErrorStatus("CIKZSBE", mRimp.errWarning, "預金種目に相違があります.")
                End If
                '//////////////////////////////////////////
                '//口座番号チェック
                If dynM.Fields("CIKZNO") <> dynS.Fields("CAKZNO") Then
                    Call pSetErrorStatus("CIKZNOE", mRimp.errWarning, "口座番号に相違があります.")
                End If
            ElseIf dynM.Fields("CIKKBN") = eBankKubun.YuubinKyoku Then
                '//////////////////////////////////////////
                '//通帳記号チェック
                If dynM.Fields("CIYBTK") <> dynS.Fields("CAYBTK") Then
                    Call pSetErrorStatus("CIYBTKE", mRimp.errWarning, "通帳記号に相違があります.")
                End If
                '//////////////////////////////////////////
                '//通帳番号チェック
                If dynM.Fields("CIYBTN") <> dynS.Fields("CAYBTN") Then
                    Call pSetErrorStatus("CIYBTNE", mRimp.errWarning, "通帳番号に相違があります.")
                End If
            Else
                Call pSetErrorStatus("CIKKBNE", mRimp.errWarning, "金融機関区分が間違っています.")
            End If
            '//////////////////////////////////////////
            '//口座名義人名チェック
            If dynM.Fields("CIKZNM") <> dynS.Fields("CAKZNM") Then
                Call pSetErrorStatus("CIKZNME", mRimp.errWarning, "口座名義人名に相違があります.")
            End If
        End If
        '//データがある場合のみで可：ＥＮＤ
        '//////////////////////////////////////////
        Call dynS.Close
        Set dynS = Nothing
        '//////////////////////////////////////
        '//金融機関チェック
        If dynM.Fields("CIKKBN") = eBankKubun.KinnyuuKikan Then
            If IsNull(dynM.Fields("CIBANK")) Then
                Call pSetErrorStatus("CIBANKE", mRimp.errWarning, "金融機関コードが未入力です.")
            Else
                sql = "SELECT * FROM tdBankMaster " & vbCrLf
                sql = sql & " WHERE DARKBN = " & gdDBS.ColumnDataSet(eBankRecordKubun.Bank, vEnd:=True) & vbCrLf
                sql = sql & "   AND DABANK = " & gdDBS.ColumnDataSet(dynM.Fields("CIBANK"), vEnd:=True) & vbCrLf
                sql = sql & "   AND DASITN = '000'" & vbCrLf
                sql = sql & " ORDER BY DECODE(DASQNO,':',0,'#',1,'@',2,'''',3,'=',4,9)" & vbCrLf
                Set dynS = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
                If dynS.EOF Then
                    Call pSetErrorStatus("CIBANKE", mRimp.errWarning, "金融機関が存在しません.")
                Else
                    '//毎回、取得はレスポンスが悪いだろう！から変数に代入してチェック
                    impName = gdDBS.Nz(dynM.Fields("CIBKNM"))
                    updFlag = mRimp.errNormal
                    Do Until dynS.EOF
                        mstName = dynS.Fields("DAKJNM")
                        For ix = LBound(BankCutName) To UBound(BankCutName)
                            If True = pCompare(impName, mstName, BankCutName(ix)) Then '//「？？？？」を取ってチェック
                                updFlag = mRimp.errNormal
                                Exit Do    '//チェックＯＫ
                            Else
                                updFlag = mRimp.errWarning
                            End If
                        Next ix
                        Call dynS.MoveNext
                    Loop
                    If updFlag <> mRimp.errNormal Then
                        Call pSetErrorStatus("CIBKNME", mRimp.errWarning, "金融機関名称が合致しません.")
                        Call pSetErrorStatus("CIBANKE", mRimp.errWarning)
                    End If
                End If
                Call dynS.Close
                Set dynS = Nothing
            End If
            '//////////////////////////////////////
            '//支店チェック
            If IsNull(dynM.Fields("CISITN")) Then
                Call pSetErrorStatus("CISITNE", mRimp.errWarning, "支店コードが未入力です.")
'//2006/07/25 支店名チェックに行ってない？ので Not 付与
'//2007/05/23 支店名称のチェックのデバッグ
            ElseIf Not IsNull(dynM.Fields("CIBANK")) Then
                sql = "SELECT * FROM tdBankMaster"
                sql = sql & " WHERE DARKBN = " & gdDBS.ColumnDataSet(eBankRecordKubun.Shiten, vEnd:=True) & vbCrLf
                sql = sql & "   AND DABANK = " & gdDBS.ColumnDataSet(dynM.Fields("CIBANK"), vEnd:=True)
                sql = sql & "   AND DASITN = " & gdDBS.ColumnDataSet(dynM.Fields("CISITN"), vEnd:=True)
                sql = sql & " ORDER BY DASQNO"
                Set dynS = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
                If dynS.EOF Then
                    Call pSetErrorStatus("CISITNE", mRimp.errWarning, "支店が存在しません.")
                Else
                    '//毎回、取得はレスポンスが悪いだろう！から変数に代入してチェック
                    impName = gdDBS.Nz(dynM.Fields("CISINM"))
                    updFlag = mRimp.errNormal
                    Do Until dynS.EOF
                        mstName = dynS.Fields("DAKJNM")
                        For ix = LBound(ShitenCutName) To UBound(ShitenCutName)
                            If True = pCompare(impName, mstName, ShitenCutName(ix)) Then '//「？？？？」を取ってチェック
                                updFlag = mRimp.errNormal
                                Exit Do    '//チェックＯＫ
                            Else
                                updFlag = mRimp.errWarning
                            End If
                        Next ix
                        Call dynS.MoveNext
                    Loop
                    If updFlag <> mRimp.errNormal Then
                        Call pSetErrorStatus("CISINME", mRimp.errWarning, "支店名称が合致しません.")
                        Call pSetErrorStatus("CISITNE", mRimp.errWarning)
                    End If
                End If
                Call dynS.Close
                Set dynS = Nothing
            End If
            '//////////////////////////////////////////
            '//預金種目チェック
            If dynM.Fields("CIKZSB") = eBankYokinShubetsu.Futsuu _
            Or dynM.Fields("CIKZSB") = eBankYokinShubetsu.Touza Then
            Else
                Call pSetErrorStatus("CIKZSBE", mRimp.errWarning, "預金種目に誤りがあります.")
            End If
            '//////////////////////////////////////////
            '//口座番号チェック
            If "" = gdDBS.Nz(dynM.Fields("CIKZNO")) Then
                Call pSetErrorStatus("CIKZNOE", mRimp.errWarning, "口座番号に誤りがあります.")
            End If
        ElseIf dynM.Fields("CIKKBN") = eBankKubun.YuubinKyoku Then
            '//////////////////////////////////////////
            '//通帳記号チェック
            If IsNull(dynM.Fields("CIYBTK")) Or Len(dynM.Fields("CIYBTK")) < mRimp.YubinKigouLength Then
                Call pSetErrorStatus("CIYBTKE", mRimp.errWarning, "通帳記号に誤りがあります.")
            End If
            '//////////////////////////////////////////
            '//通帳番号チェック
            If IsNull(dynM.Fields("CIYBTN")) Or Len(dynM.Fields("CIYBTN")) < mRimp.YubinBangoLength Then
                Call pSetErrorStatus("CIYBTNE", mRimp.errWarning, "通帳番号に誤りがあります.")
            ElseIf "1" <> Right(dynM.Fields("CIYBTN"), 1) Then
                Call pSetErrorStatus("CIYBTNE", mRimp.errWarning, "通帳番号に誤りがあります(末尾が１以外).")
            End If
        Else
            Call pSetErrorStatus("CIKKBNE", mRimp.errWarning, "金融機関区分に誤りがあります.")
        End If
'//2006/04/26 金融機関・郵便局の両方入力がある
'//2007/06/12 両方あっても入力が正常であれば？良いだろう。と思ったが？？？
        If "" <> gdDBS.Nz(dynM.Fields("CIYBTK")) & gdDBS.Nz(dynM.Fields("CIYBTK")) _
        And "" <> gdDBS.Nz(dynM.Fields("CIBANK")) & gdDBS.Nz(dynM.Fields("CISITN")) & gdDBS.Nz(dynM.Fields("CIKZNO")) Then
            Call pSetErrorStatus("CIKKBNE", mRimp.errWarning, "金融機関/郵便局の両方に入力があります.")
        End If
        
        '////////////////////////////////////////////////
        '//エラーの配列が存在すれば UPDATE 文を生成
        If 0 <= pErrorCount() Then
            sql = "UPDATE " & mRimp.TcHogoshaImport & " SET " & vbCrLf
            msg = ""
            For ix = LBound(mErrStts) To UBound(mErrStts)
                msg = msg & mErrStts(ix).Message & vbCrLf
                sql = sql & mErrStts(ix).Field & " = " & mErrStts(ix).Error & "," & vbCrLf
            Next ix
            sql = sql & " CIWMSG = '" & msg & "'," & vbCrLf
            sql = sql & " CIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
            sql = sql & " CIUPDT = SYSDATE" & vbCrLf
            sql = sql & " WHERE CISEQN = " & dynM.Fields("CISEQN") & vbCrLf
            sql = sql & SameConditions & vbCrLf
            recCnt = gdDBS.Database.ExecuteSQL(sql)
        End If
        Call dynM.MoveNext
    Loop
    Call dynM.Close
    Set dynM = Nothing
    pgrProgressBar.Max = cMaxBlock
    
    '//////////////////////////////////////////////////
    '//全体エラー項目セット：最初に正常にしているので「正常」フラグは不要
    '//異常データ
    '//////////////////////////////////////////////////
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
    sql = "UPDATE " & mRimp.TcHogoshaImport & " a SET " & vbCrLf
    sql = sql & " CIOKFG =  " & mRimp.updInvalid & "," & vbCrLf    '//マスタ反映不可
    sql = sql & " CIERROR = " & mRimp.errInvalid & "," & vbCrLf
    sql = sql & " CIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
    sql = sql & " CIUPDT = SYSDATE" & vbCrLf
    sql = sql & " WHERE(" & vbCrLf
    sql = sql & mRimp.StatusColumns(" = " & mRimp.errInvalid & vbCrLf & " OR ", Len(vbCrLf & " OR ")) & vbCrLf & ")" & vbCrLf
    sql = sql & SameConditions & vbCrLf
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    '//////////////////////////////////////////////////
    '//全体エラー項目セット：最初に正常にしているので「正常」フラグは不要
    '//警告データ：マスタ反映しないデータ
    '//////////////////////////////////////////////////
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
    sql = "UPDATE " & mRimp.TcHogoshaImport & " a SET " & vbCrLf
    sql = sql & " CIOKFG =  " & mRimp.updWarnErr & "," & vbCrLf   '//マスタ反映しないフラグ
    sql = sql & " CIERROR = " & mRimp.errWarning & "," & vbCrLf
    sql = sql & " CIUSID = '" & gdDBS.LoginUserName & "'," & vbCrLf
    sql = sql & " CIUPDT = SYSDATE" & vbCrLf
    sql = sql & " WHERE CIERROR = " & mRimp.errNormal & vbCrLf    '//異常で無い
    sql = sql & "   AND CIOKFG <= " & mRimp.updNormal & vbCrLf
    sql = sql & "   AND(" & vbCrLf
    sql = sql & mRimp.StatusColumns(" >= " & mRimp.errWarning & vbCrLf & " OR ", Len(vbCrLf & " OR ")) & vbCrLf & ")" & vbCrLf
    sql = sql & SameConditions & vbCrLf
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    '//////////////////////////////////////////////////
    '//ソート用に CIERROR=>CIERSR にコピー
    '//Spreadで仮想モードにするとリアルに変わる為、修正部=CIEROR、固定部=CIERSR とする
    '//////////////////////////////////////////////////
    If False = pProgressBarSet(Block) Then
        GoTo gDataCheckError:
    End If
    sql = "UPDATE " & mRimp.TcHogoshaImport & " a SET " & vbCrLf
    sql = sql & " CIERSR = CIERROR "
    sql = sql & " WHERE 1 = 1"  '//おまじない
    sql = sql & SameConditions & vbCrLf
    If -1 <> vSeqNo Then        '//行指定時には更新しないおまじない：仮想モードでリアルに変わる為
        sql = sql & " AND 1 = -1"
    End If
    recCnt = gdDBS.Database.ExecuteSQL(sql)
    
    Call gdDBS.Database.CommitTrans         '//トランザクション正常終了
    fraProgressBar.Visible = False
    Call gdDBS.AutoLogOut(mCaption, "[" & vImpDate & ":" & vSeqNo & "] のチェック処理が完了しました。")
    '//ステータス行の整列・調整
    stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "チェック完了"
    gDataCheck = True
    
#If BLOCK_CHECK = True Then           '//チェック時のブロックがいくつあるか？を表示：デバック時のみ
     Call MsgBox("チェックしたブロックは " & mCheckBlocks & " 箇所でした。")
#End If
    
    Exit Function
gDataCheckError:
    fraProgressBar.Visible = False
    Call gdDBS.Database.Rollback            '//トランザクション異常終了
    If err Then
        Dim errCode As Integer, errMsg As String
        If gdDBS.Database.LastServerErr Then
            errCode = gdDBS.Database.LastServerErr
            errMsg = gdDBS.Database.LastServerErrText
        Else
            errCode = err
            errMsg = Error
        End If
        fraProgressBar.Visible = False
        '//ステータス行の整列・調整
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "チェックエラー(" & errCode & ")"
        Call gdDBS.AutoLogOut(mCaption, "[" & vImpDate & ":" & vSeqNo & "] のチェック処理中にエラーが発生しました。(Error=" & errCode & ")")
        Call MsgBox("チェック対象 = [" & cboImpDate.Text & "]" & vbCrLf & _
                    "はエラーが発生したためチェックは中止されました。" & vbCrLf & errMsg, _
                vbOKOnly + vbCritical, mCaption)
    Else
        Call gdDBS.AutoLogOut(mCaption, "[" & vImpDate & ":" & vSeqNo & "] のチェック処理が中断されました。")
        '//ステータス行の整列・調整
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "チェック中断"
    End If
End Function

Private Sub cmdCheck_Click()
    If False = pCheckSubForm Then
        Exit Sub
    End If
    If -1 <> pAbortButton(cmdCheck, cBtnCheck) Then
        Exit Sub
    End If
    cmdCheck.Caption = cBtnCancel
    '//コマンド・ボタン制御
    Call pLockedControl(False, cmdCheck)
    '//チェック処理
    If True = gDataCheck(cboImpDate.Text) Then
        '//データ読み込み＆ Spread に設定反映
        Call pReadDataAndSetting
    End If
    '//ボタンを戻す
    cmdCheck.Caption = cBtnCheck
    '//コマンド・ボタン制御
    Call pLockedControl(True)
End Sub

#If ORA_DEBUG = 1 Then
Private Function pHogoshaInsert(vInDyn As OraDynaset) As Boolean
#Else
Private Function pHogoshaInsert(vInDyn As Object) As Boolean
#End If
    Dim sql As String
    sql = "INSERT INTO tcHogoshaMaster ( " & vbCrLf
    sql = sql & "CAITKB," & vbCrLf  '//委託者区分
    sql = sql & "CAKYCD," & vbCrLf  '//契約者番号
    sql = sql & "CAKSCD," & vbCrLf  '//教室番号
    sql = sql & "CAHGCD," & vbCrLf  '//保護者番号
    sql = sql & "CASQNO," & vbCrLf  '//保護者ＳＥＱ
    sql = sql & "CAKJNM," & vbCrLf  '//保護者名_漢字
    sql = sql & "CAKNNM," & vbCrLf  '//保護者名_カナ
    sql = sql & "CASTNM," & vbCrLf  '//生徒氏名
    sql = sql & "CAKKBN," & vbCrLf  '//取引金融機関区分
    sql = sql & "CABANK," & vbCrLf  '//取引銀行
    sql = sql & "CASITN," & vbCrLf  '//取引支店
    sql = sql & "CAKZSB," & vbCrLf  '//口座種別
    sql = sql & "CAKZNO," & vbCrLf  '//口座番号
    sql = sql & "CAYBTK," & vbCrLf  '//通帳記号
    sql = sql & "CAYBTN," & vbCrLf  '//通帳番号
    sql = sql & "CAKZNM," & vbCrLf  '//口座名義人_カナ
    sql = sql & "CAKYST," & vbCrLf  '//契約開始日
    sql = sql & "CAKYED," & vbCrLf  '//契約終了日
    sql = sql & "CAFKST," & vbCrLf  '//振替開始日
    sql = sql & "CAFKED," & vbCrLf  '//振替終了日
    sql = sql & "CASKGK," & vbCrLf  '//請求予定額
    sql = sql & "CAHKGK," & vbCrLf  '//変更後金額
    sql = sql & "CAKYDT," & vbCrLf  '//解約日
    sql = sql & "CAKYFG," & vbCrLf  '//解約フラグ
    sql = sql & "CATRFG," & vbCrLf  '//伝送更新フラグ
    sql = sql & "CAUSID," & vbCrLf  '//作成日
    sql = sql & "CAADDT," & vbCrLf  '//更新日
    sql = sql & "CANWDT " & vbCrLf  '//新規データ扱い日
    sql = sql & ") SELECT " & vbCrLf
    sql = sql & "CiITKB," & vbCrLf  '//委託者区分
    sql = sql & "CiKYCD," & vbCrLf  '//契約者番号
    sql = sql & "CiKSCD," & vbCrLf  '//教室番号
    sql = sql & "CiHGCD," & vbCrLf  '//保護者番号
    sql = sql & "TO_CHAR(SYSDATE,'yyyymmdd')," & vbCrLf  '//保護者ＳＥＱ
    sql = sql & "CiKJNM," & vbCrLf  '//保護者名_漢字
    sql = sql & "CiKNNM," & vbCrLf  '//保護者名_カナ
    sql = sql & "CiSTNM," & vbCrLf  '//生徒氏名
    sql = sql & "CiKKBN," & vbCrLf  '//取引金融機関区分
    sql = sql & "CiBANK," & vbCrLf  '//取引銀行
    sql = sql & "CiSITN," & vbCrLf  '//取引支店
    sql = sql & "CiKZSB," & vbCrLf  '//口座種別
    sql = sql & "CiKZNO," & vbCrLf  '//口座番号
    sql = sql & "CiYBTK," & vbCrLf  '//通帳記号
    sql = sql & "CiYBTN," & vbCrLf  '//通帳番号
    sql = sql & "CiKZNM," & vbCrLf  '//口座名義人_カナ
    sql = sql & "     0," & vbCrLf  '//契約開始日
    sql = sql & "20991231," & vbCrLf  '//契約終了日
    sql = sql & "     0," & vbCrLf  '//振替開始日
    sql = sql & "20991231," & vbCrLf  '//振替終了日
    sql = sql & "CiSKGK," & vbCrLf  '//請求予定額
    sql = sql & "  NULL," & vbCrLf  '//変更後金額
    sql = sql & "  NULL," & vbCrLf  '//解約日
    sql = sql & "     0," & vbCrLf  '//解約フラグ
    sql = sql & "  NULL," & vbCrLf  '//伝送更新フラグ
    sql = sql & gdDBS.ColumnDataSet(MainModule.gcImportHogoshaUser) & vbCrLf    '//更新者ＩＤ
    sql = sql & "SYSDATE," & vbCrLf  '//更新日
    sql = sql & "   NULL " & vbCrLf  '//新規データ扱い日
    sql = sql & " FROM " & mRimp.TcHogoshaImport
    sql = sql & " WHERE CIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss') " & vbCrLf
    sql = sql & "   AND CIKYCD = " & gdDBS.ColumnDataSet(vInDyn.Fields("CIKYCD"), vEnd:=True)
    sql = sql & "   AND CIHGCD = " & gdDBS.ColumnDataSet(vInDyn.Fields("CIHGCD"), vEnd:=True)
    sql = sql & "   AND CISEQN = " & gdDBS.ColumnDataSet(vInDyn.Fields("CISEQN"), vEnd:=True)
    Call gdDBS.Database.ExecuteSQL(sql)
    pHogoshaInsert = True
End Function

#If ORA_DEBUG = 1 Then
Private Function pHogoshaUpdate(vOutDyn As OraDynaset, vInDyn As OraDynaset) As Boolean
#Else
Private Function pHogoshaUpdate(vOutDyn As Object, vInDyn As Object) As Boolean
#End If
    Dim Fields As Variant, ix As Integer, chg As Boolean
    Dim sql As String
    Fields = Array("CaITKB", "CaKYCD", "CaKSCD", "CaHGCD", "CaKJNM", "CaKNNM", "CaSTNM", "CaKKBN", _
                   "CaBANK", "CaSITN", "CaKZSB", "CaKZNO", "CaYBTK", "CaYBTN", "CaKZNM", "CaSKGK")
    '//入力の相違分のみ更新する
    For ix = LBound(Fields) To UBound(Fields)
        chg = False
        '//2007/04/24 相手が NULL であると違うと判断されて更新する項目ではなくなるバグ修正
        If IsNull(vOutDyn.Fields(Fields(ix))) And Not IsNull(vInDyn.Fields("Ci" & Mid(Fields(ix), 3))) Then
            '//出力先が片方 NULL
            chg = True
        ElseIf Not IsNull(vOutDyn.Fields(Fields(ix))) And IsNull(vInDyn.Fields("Ci" & Mid(Fields(ix), 3))) Then
            '//入力先が片方 NULL
            chg = True
        ElseIf vOutDyn.Fields(Fields(ix)) <> vInDyn.Fields("Ci" & Mid(Fields(ix), 3)) Then
            '//出力先と入力先に相違が有る
            chg = True
        End If
        If True = chg Then
            sql = sql & Fields(ix) & " = " & gdDBS.ColumnDataSet(vInDyn.Fields("Ci" & Mid(Fields(ix), 3)), "S") & vbCrLf
        End If
    Next ix
'//パンチデータとの件数が合わなくなるのでやめた：常に何らかは更新する
#If 0 Then
    '//解約解除でなく、すべての列に変更が無ければ更新しない
    If mRimp.updResetCancel <> vInDyn.Fields("CiOKFG") And "" = sql Then
        pHogoshaUpdate = True
        Exit Function
    End If
#End If
    sql = "UPDATE tcHogoshaMaster SET " & sql   '//上で定義した構文を「最後に」に付加
    If mRimp.updResetCancel = vInDyn.Fields("CiOKFG") Then
        sql = sql & " CAKYED = CASE WHEN CAKYED < 20991231 THEN 20991231 END," & vbCrLf
        sql = sql & " CAFKED = CASE WHEN CAFKED < 20991231 THEN 20991231 END," & vbCrLf
        sql = sql & " CAKYDT = NULL," & vbCrLf
        sql = sql & " CAKYFG = 0," & vbCrLf
    End If
    sql = sql & " CAUSID = " & gdDBS.ColumnDataSet(MainModule.gcImportHogoshaUser) & vbCrLf
    sql = sql & " CAUPDT = SYSDATE" & vbCrLf
    '//既に更新するべき該当レコードは読み出し済み
    sql = sql & " WHERE CAKYCD = " & gdDBS.ColumnDataSet(vOutDyn.Fields("CAKYCD"), vEnd:=True) & vbCrLf
    sql = sql & "   AND CAKSCD = " & gdDBS.ColumnDataSet(vOutDyn.Fields("CAKSCD"), vEnd:=True) & vbCrLf
    sql = sql & "   AND CAHGCD = " & gdDBS.ColumnDataSet(vOutDyn.Fields("CAHGCD"), vEnd:=True) & vbCrLf
    sql = sql & "   AND CASQNO = " & gdDBS.ColumnDataSet(vOutDyn.Fields("CASQNO"), "L", vEnd:=True) & vbCrLf
    Call gdDBS.Database.ExecuteSQL(sql)
    pHogoshaUpdate = True
End Function

Private Sub cmdUpdate_Click()
    If False = pCheckSubForm Then
        Exit Sub
    End If
    If -1 <> pAbortButton(cmdUpdate, cBtnUpdate) Then
        Exit Sub
    End If
    If vbOK <> MsgBox("マスタの反映を開始します。" & vbCrLf & vbCrLf & "よろしいですか？", vbOKCancel + vbInformation, Me.Caption) Then
        Exit Sub
    End If
    cmdUpdate.Caption = cBtnCancel
    '//コマンド・ボタン制御
    Call pLockedControl(False, cmdUpdate)
    
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset, msg As String
#Else
    Dim sql As String, dyn As Object, msg As String
#End If
    '//////////////////////////////////////////////////////////
    '//ここで使用する共通の WHERE 条件
    Dim Condition As String
    Condition = Condition & " AND CIINDT = TO_DATE('" & cboImpDate.Text & "','yyyy/mm/dd hh24:mi:ss') " & vbCrLf
    '// CIERROR >= 0 AND CIOKFG >= 0 であること
    Condition = Condition & " AND CIERROR >= 0" & vbCrLf
    Condition = Condition & " AND CIOKFG  >= 0"
    Condition = Condition & " AND CIMUPD   = 0" '//2006/04/04 マスタ反映ＯＫフラグ項目追加
    '///////////////////////////////////////
    '// 取込日時単位で TcHogoshaImport 内に同じ保護者が存在しないこと
    '//2006/03/17 重複データは後勝ちで更新するように変更にしたのでありえないだろう？
    '//2006/04/24 教室番号を追加
    sql = " SELECT CIKYCD,CIKSCD,CIHGCD"
    sql = sql & " FROM " & mRimp.TcHogoshaImport
    sql = sql & " WHERE 1 = 1"  '//おまじない
    sql = sql & Condition
    sql = sql & " GROUP BY CIKYCD,CIKSCD,CIHGCD"
    sql = sql & " HAVING COUNT(*) > 1 "     '//同一の保護者が存在するか？
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If Not dyn.EOF Then
        msg = "取込日時 [ " & cboImpDate.Text & " ] 内に" & vbCrLf & _
              "　 保護者 [ " & dyn.Fields("CIKYCD") & " - " & dyn.Fields("CIHGCD") & " ] が複数存在する為     " & vbCrLf & _
              "マスタ反映は処理続行が出来ません。"
    End If
    Call dyn.Close
    Set dyn = Nothing
    If "" <> msg Then
        Call MsgBox(msg, vbOKOnly + vbCritical, mCaption)
        '//ボタンを戻す
        cmdUpdate.Caption = cBtnUpdate
        '//コマンド・ボタン制御
        Call pLockedControl(True)
        Exit Sub
    End If
    
    Call gdDBS.AutoLogOut(mCaption, "[" & cboImpDate.Text & "] のマスタ反映が開始されました。")
    
    On Error GoTo cmdUpdate_ClickError:
    Call gdDBS.Database.BeginTrans
    
#If ORA_DEBUG = 1 Then
    Dim updDyn As OraDynaset, recCnt As Long
#Else
    Dim updDyn As Object, recCnt As Long
#End If
    Dim ms As New MouseClass
    Call ms.Start
    
    sql = "SELECT a.*" & vbCrLf
    sql = sql & " FROM " & mRimp.TcHogoshaImport & " a " & vbCrLf
    sql = sql & " WHERE 1 = 1" & vbCrLf
    sql = sql & Condition & vbCrLf
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    recCnt = dyn.RecordCount
    If dyn.EOF Then
        Call dyn.Close
        Set dyn = Nothing
        Call MsgBox("取込日時 [ " & cboImpDate.Text & " ]" & vbCrLf & "にマスタ反映すべきデータはありません。", vbOKOnly + vbInformation, mCaption)
        '//ボタンを戻す
        cmdUpdate.Caption = cBtnUpdate
        '//コマンド・ボタン制御
        Call pLockedControl(True)
        Exit Sub
    End If
'//2007/07/19 口座戻り件数を表示
    Dim modoriCnt As Long
'//2007/06/11 大量に AutoLog にかかれるのでトリガを停止
    Call gdDBS.TriggerControl("tcHogoshaMaster", False)
    fraProgressBar.Visible = True
    pgrProgressBar.Max = recCnt
    Do Until dyn.EOF
        DoEvents
        If mAbort Then
            GoTo cmdUpdate_ClickError
        End If
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = _
            "残り" & Right(String(7, " ") & Format(recCnt - dyn.RowPosition, "#,##0"), 7) & " 件"
        pgrProgressBar.Value = dyn.RowPosition
        sql = "SELECT b.* "
        sql = sql & " FROM tcHogoshaMaster b "
        sql = sql & " WHERE CAKYCD = " & gdDBS.ColumnDataSet(dyn.Fields("CIKYCD"), vEnd:=True)
        sql = sql & "   AND CAKSCD = " & gdDBS.ColumnDataSet(dyn.Fields("CIKSCD"), vEnd:=True)
        sql = sql & "   AND CAHGCD = " & gdDBS.ColumnDataSet(dyn.Fields("CIHGCD"), vEnd:=True)
        sql = sql & " ORDER BY CASQNO DESC"     '//最終レコードのみが更新対象
#If ORA_DEBUG = 1 Then
        Set updDyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
        Set updDyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
        If updDyn.EOF Then
            If False = pHogoshaInsert(dyn) Then
                GoTo cmdUpdate_ClickError:
            End If
        Else
            If False = pHogoshaUpdate(updDyn, dyn) Then
                GoTo cmdUpdate_ClickError:
            End If
            modoriCnt = modoriCnt + 1
        End If
        Call updDyn.Close
        Set updDyn = Nothing
        Call dyn.MoveNext
    Loop
    Call dyn.Close
    Set dyn = Nothing
    '//マスタ反映時にも同じ事をするので共通化
    If pMoveTempRecords(Condition, gcFurikaeImportToMaster) < 0 Then
        GoTo cmdUpdate_ClickError:
    End If
    Call gdDBS.Database.CommitTrans
'//2007/06/11 先頭で停止しているのでトリガを再開
    Call gdDBS.TriggerControl("tcHogoshaMaster")
    
    pgrProgressBar.Max = pgrProgressBar.Max
    fraProgressBar.Visible = False
    
    '//ステータス行の整列・調整
    stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "反映完了"
    Call MsgBox("マスタ反映対象 = [" & cboImpDate.Text & "]" & vbCrLf & vbCrLf & _
                recCnt & " 件がマスタ反映されました." & vbCrLf & vbCrLf & _
                "内、口座戻りの件数は " & modoriCnt & " 件です。", vbOKOnly + vbInformation, mCaption)
    Call gdDBS.AutoLogOut(mCaption, "[" & cboImpDate.Text & "] の " & recCnt & " 件の反映が完了しました。内、口座戻りの件数は " & modoriCnt & " 件です。")
    '//リストを再設定
    Call pMakeComboBox
    '//ボタンを戻す
    cmdUpdate.Caption = cBtnUpdate
    '//コマンド・ボタン制御
    Call pLockedControl(True)
    Exit Sub
cmdUpdate_ClickError:
    Call gdDBS.Database.Rollback
'//2007/06/11 先頭で停止しているのでトリガを再開
    Call gdDBS.TriggerControl("tcHogoshaMaster")
    If err Then
        Dim errCode As Integer, errMsg As String
        If gdDBS.Database.LastServerErr Then
            errCode = gdDBS.Database.LastServerErr
            errMsg = gdDBS.Database.LastServerErrText
        Else
            errCode = err
            errMsg = Error
        End If
        fraProgressBar.Visible = False
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "反映エラー(" & errCode & ")"
        Call gdDBS.AutoLogOut(mCaption, "マスタ反映対象 = [" & cboImpDate.Text & "] はエラーが発生したためマスタ反映は中止されました。(Error=" & errMsg & ")")
        Call MsgBox("マスタ反映対象 = [" & cboImpDate.Text & "]" & vbCrLf & _
                    "はエラーが発生したためマスタ反映は中止されました。" & vbCrLf & errMsg, _
                vbOKOnly + vbCritical, mCaption)
    Else
        stbStatus.Panels.Item(stbStatus.Panels.Count).Text = "反映中断"
        Call gdDBS.AutoLogOut(mCaption, "マスタ反映対象 = [" & cboImpDate.Text & "]" & vbCrLf & "のマスタ反映は中止されました。")
    End If
    '//ボタンを戻す
    cmdUpdate.Caption = cBtnUpdate
    '//コマンド・ボタン制御
    Call pLockedControl(True)
End Sub

Private Sub pMakeComboBox()
    Dim ms As New MouseClass
    Call ms.Start
    '//コマンド・ボタン制御
    Call pLockedControl(False)
'    Dim sql As String, dyn As OraDynaset, MaxDay As Variant
    Dim sql As String, dyn As Object, MaxDay As Variant
    sql = "SELECT DISTINCT TO_CHAR(CIINDT,'yyyy/mm/dd hh24:mi:ss') CIINDT_A"
    sql = sql & " FROM " & mRimp.TcHogoshaImport
    sql = sql & " ORDER BY CIINDT_A"
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
    Call cboImpDate.Clear
    Do Until dyn.EOF()
        Call cboImpDate.AddItem(dyn.Fields("CIINDT_A"))
        'cboImpDate.ItemData(cboImpDate.NewIndex) = dyn.Fields("CIINDT_B")
        Call dyn.MoveNext
    Loop
    Call dyn.Close
    If cboImpDate.ListCount Then
        cboImpDate.ListIndex = cboImpDate.ListCount - 1
    Else
        sprMeisai.MaxRows = 0
    End If
    '//コマンド・ボタン制御
    Call pLockedControl(True)
End Sub

Private Sub Form_Activate()
'    If sprMeisai.ColWidth(eSprCol.eMaxCols) Then
End Sub

Private Sub Form_Load()
    Me.Show
    mCaption = Me.Caption
    Call mForm.Init(Me, gdDBS)
    Call mSpread.Init(sprMeisai)
    lblModoriCount.Caption = "【 口座戻り件数： " & Format(0, "#,0") & " 件 】"
    lblModoriCount.Refresh
    '//Spreadの列調整
    Dim ix As Long
    With sprMeisai
        Call sprMeisai_LostFocus    '//ToolTip を設定
        .MaxCols = eSprCol.eMaxCols
        '//エラー列もあるので表示列(eUseCol)以降は非表示にする
        For ix = eSprCol.eUseCols + 1 To eSprCol.eMaxCols
            .ColWidth(ix) = 0
        Next ix
        '.ColWidth(eSprCol.eImpDate) = 0
        '.ColWidth(eSprCol.eImpSEQ) = 0
    End With
    '//ステータス行の整列・調整
    stbStatus.Panels.Item(stbStatus.Panels.Count).Text = ""
    pgrProgressBar.Left = 15
    pgrProgressBar.Top = 15
    pgrProgressBar.Height = 255
    pgrProgressBar.Width = 7035
    fraProgressBar.Height = pgrProgressBar.Height + 30
    fraProgressBar.Width = pgrProgressBar.Width + 30
    fraProgressBar.Visible = False
    cboSort.ListIndex = 0
    Call fraProgressBar.ZOrder(0)   '//最前面に
    Call pMakeComboBox
    'Call cmdEnd.SetFocus
End Sub

Private Sub Form_Resize()
    '//これ以上小さくするとコントロールが隠れるので制御する
    If Me.Height < 8500 Then
        Me.Height = 8500
    End If
    If Me.Width < 11220 Then
        Me.Width = 11220
    End If
    Call mForm.Resize
    fraProgressBar.Left = 1860
    fraProgressBar.Top = Me.Height - 970
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mAbort = True
    Set mForm = Nothing
    Set mReg = Nothing
    If Not gdFormSub Is Nothing Then
        Unload gdFormSub
    End If
    Set gdFormSub = Nothing
    '//最後にしないとこのフォームの他からの参照により再ロードされる
    Set frmFurikaeReqImport = Nothing
    Call gdForm.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Public Sub gEditToSpreadSheet(vMove As Integer)
'// vMove => -1:前方移動 / 0:移動無し / 1:後方移動
'CIITKB eItakuName           '委託者名
'CIKYCD eKeiyakuCode         '契約者コード
'       eKeiyakuName         '契約者名
'CIKSCD eKyoshitsuNo         '教室番号
'CIHGCD eHogoshaCode         '保護者コード
'CIKJNM eHogoshaName         '保護者名(漢字)
'CIKNNM eHogoshaKana         '保護者名(カナ)=>口座名義人名
'CISTNM eSeitoName           '生徒氏名
'CISKGK eFurikaeGaku         '振替金額
'CIKKBN eKinyuuKubun         '金融機関区分
'CIBANK eBankCode            '銀行コード
'       eBankName_m          '銀行名(マスター)
'CIBKNM eBankName_i          '銀行名(取込)
'CISITN eShitenCode          '支店コード
'       eShitenName_m        '支店名(マスター)
'CISINM eShitenName_i        '支店名(取込)
'CIKZSB eYokinShumoku        '預金種目
'CIKZNO eKouzaBango          '口座番号
'CIYBTK eYubinKigou          '郵便局:通帳記号
'CIYBTN eYubinBango          '郵便局:通帳番号
'CIKZNM eKouzaName           '口座名義人=>保護者名(カナ)
'CIINDT eImpDate             '取込日
'CISEQN eImpSEQ              'ＳＥＱ

    '//行のデータが一致していなければ置換えしない
    If Not (Format(gdFormSub.lblCIINDT.Caption, "yyyy/MM/dd hh:nn:ss") = Format(mSpread.Value(eSprCol.eImpDate, mEditRow), "yyyy/MM/dd hh:nn:ss") _
        And gdFormSub.lblCISEQN.Caption = mSpread.Value(eSprCol.eImpSEQ, mEditRow) _
      ) Then
        Call MsgBox("行データが異常な為" & vbCrLf & "更新出来ませんでした.", vbOKOnly + vbCritical, mCaption)
        Exit Sub
    End If
    Dim obj As Object
    mSpread.Value(eSprCol.eErrorStts, mEditRow) = cEditDataMsg
    mSpread.BackColor(eSprCol.eErrorStts, mEditRow) = mRimp.ErrorStatus(mRimp.errEditData)
    For Each obj In gdFormSub.Controls
        If TypeOf obj Is imText _
        Or TypeOf obj Is imNumber _
        Or TypeOf obj Is imDate _
        Or TypeOf obj Is Label Then
            '//コントロールの DataChanged プロパティを検査して更新を必要とするか判断
            If "" <> obj.DataField And True = obj.DataChanged Then
                Select Case UCase(Right(obj.Name, 6))
                Case "CIITKB" '//eItakuName           '委託者名
                    mSpread.Value(eSprCol.eItakuName, mEditRow) = gdFormSub.cboABKJNM.Text
                Case "CIKYCD" '//eKeiyakuCode         '契約者コード
                              '//eKeiyakuName         '契約者名
                    mSpread.Value(eSprCol.eKeiyakuCode, mEditRow) = obj.Text
                    mSpread.Value(eSprCol.eKeiyakuName, mEditRow) = gdFormSub.lblBAKJNM.Caption
                Case "CIKSCD" '//eKyoshitsuNo         '教室番号
                    mSpread.Value(eSprCol.eKyoshitsuNo, mEditRow) = obj.Text
                Case "CIHGCD" '//eHogoshaCode         '保護者コード
                    mSpread.Value(eSprCol.eHogoshaCode, mEditRow) = obj.Text
                Case "CIKJNM" '//eHogoshaName         '保護者名(漢字)
                    mSpread.Value(eSprCol.eHogoshaName, mEditRow) = obj.Text
                Case "CIKNNM" '//eHogoshaKana         '保護者名(カナ)=>口座名義人名
                    mSpread.Value(eSprCol.eHogoshaKana, mEditRow) = obj.Text
                Case "CISTNM" '//eSeitoName           '生徒氏名
                    mSpread.Value(eSprCol.eSeitoName, mEditRow) = obj.Text
                Case "CISKGK" '//eFurikaeGaku         '振替金額
                    mSpread.Value(eSprCol.eFurikaeGaku, mEditRow) = obj.Text
                Case "CIKKBN" '//eKinyuuKubun         '金融機関区分
                    If 0 = gdFormSub.lblCIKKBN.Caption Or 1 = gdFormSub.lblCIKKBN.Caption Then
                        mSpread.Value(eSprCol.eKinyuuKubun, mEditRow) = gdFormSub.optCIKKBN(gdFormSub.lblCIKKBN.Caption).Caption
                    End If
                Case "CIBANK" '//eBankCode            '銀行コード
                              '//eBankName_m          '銀行名(マスター)
                    mSpread.Value(eSprCol.eBankCode, mEditRow) = obj.Text
                    mSpread.Value(eSprCol.eBankName_m, mEditRow) = gdFormSub.lblBankName.Caption
                Case "CIBKNM" '//eBankName_i          '銀行名(取込)
                    mSpread.Value(eSprCol.eBankName_i, mEditRow) = obj.Text
                Case "CISITN" '//eShitenCode          '支店コード
                              '//eShitenName_m        '支店名(マスター)
                    mSpread.Value(eSprCol.eShitenCode, mEditRow) = obj.Text
                    mSpread.Value(eSprCol.eShitenName_m, mEditRow) = gdFormSub.lblShitenName.Caption
                Case "CISINM" '//eShitenName_i        '支店名(取込)
                    mSpread.Value(eSprCol.eShitenName_i, mEditRow) = obj.Text
                Case "CIKZSB" '//eYokinShumoku        '預金種目
                    If 1 = gdFormSub.lblCIKZSB.Caption Or 2 = gdFormSub.lblCIKZSB.Caption Then
                        mSpread.Value(eSprCol.eYokinShumoku, mEditRow) = gdFormSub.optCIKZSB(gdFormSub.lblCIKZSB.Caption).Caption
                    End If
                Case "CIKZNO" '//eKouzaBango          '口座番号
                    mSpread.Value(eSprCol.eKouzaBango, mEditRow) = obj.Text
                Case "CIYBTK" '//eYubinKigou          '郵便局:通帳記号
                    mSpread.Value(eSprCol.eYubinKigou, mEditRow) = obj.Text
                Case "CIYBTN" '//eYubinBango          '郵便局:通帳番号
                    mSpread.Value(eSprCol.eYubinBango, mEditRow) = obj.Text
                Case "CIKZNM" '//eKouzaName           '口座名義人=>保護者名(カナ)
                    mSpread.Value(eSprCol.eKouzaName, mEditRow) = obj.Text
                End Select
            End If
        End If
    Next obj
    mEditRow = mEditRow + vMove   '//-1:前方移動 / 0:移動無し / 1:後方移動
End Sub

Private Sub sprMeisai_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Not gdFormSub Is Nothing Then
        '//効かない？
        'If gdFormSub.dbcImport.EditMode <> OracleConstantModule.ORADATA_EDITNONE Then
            If vbOK <> MsgBox("現在編集中のデータは破棄されます.", vbOKCancel + vbInformation, mCaption) Then
                Exit Sub
            End If
            Call gdFormSub.dbcImport.UpdateControls   '//キャンセル
        'End If
        'Unload gdFormSub
    End If
    If Row <= 0 Then
        Exit Sub
    End If
    '//修正画面へ渡す
    mEditRow = Row
    Set gdFormSub = frmFurikaeReqImportEdit
    Call gdFormSub.Show
    gdFormSub.dbcImportEdit.RecordSource = "SELECT * " & mMainSQL
    Call gdFormSub.dbcImportEdit.Refresh
    Call gdFormSub.dbcImportEdit.Recordset.FindFirst( _
            "     CIINDT = TO_DATE('" & Format(mSpread.Value(eSprCol.eImpDate, Row), "yyyy/MM/dd hh:nn:ss") & "','yyyy/mm/dd hh24:mi:ss') " & _
            " AND CISEQN = " & mSpread.Value(eSprCol.eImpSEQ, Row))
    Call gdFormSub.txtCIKYCD_KeyDown(vbKeyReturn, 0)    '//契約者名を強制表示
End Sub

Private Sub sprMeisai_LostFocus()
    With sprMeisai
        .TextTipDelay = 1
        .TextTip = TextTipFixedFocusOnly
        .ToolTipText = "クリックすると" & vbCrLf & "「取込＆チェックの処理結果」の" & vbCrLf & "詳細が表示されます."
    End With
End Sub

Private Sub sprMeisai_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    If 0 < Row Then
        sprMeisai.ToolTipText = mSpread.Value(eSprCol.eErrorStts, Row)
        '//機能しない！
        'sprMeisai.SetTextTipAppearance "ＭＳ ゴシック", 15, True, True, vbBlue, vbWhite
    End If
End Sub

Private Sub sprMeisai_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    '// OldTop = 1 の時はイベントが起きない
#If True = VIRTUAL_MODE Then
    Call pSpreadSetErrorStatus
#Else
    If OldTop <> NewTop Then     '//すべてバッファにあるので前行に戻る時はしないように
        Call pSpreadSetErrorStatus
    End If
#End If
End Sub

'//セル単位にエラー箇所をカラー表示
Private Sub pSpreadSetErrorStatus(Optional vReset As Boolean = False)
#If ORA_DEBUG = 1 Then
    Dim sql As String, dyn As OraDynaset
#Else
    Dim sql As String, dyn As Object
#End If
    Dim ErrStts() As Variant, ix As Integer, cnt As Long
    Dim ms As New MouseClass
    Call ms.Start
'    eErrorStts = 1  'エラー内容：異常、正常、警告
'    eItakuName      '委託者名
'    eKeiyakuCode    '契約者コード
'    eKeiyakuName    '契約者名
'    eKyoshitsuNo    '教室番号
'    eHogoshaCode    '保護者コード
'    eHogoshaName    '保護者名(漢字)
'    eHogoshaKana    '保護者名(カナ)=>口座名義人名
'    eSeitoName      '生徒氏名
'    eFurikaeGaku    '振替金額
'    eKinyuuKubun    '金融機関区分
'    eBankCode       '銀行コード
'    eBankName_m     '銀行名(マスター)
'    eBankName_i     '銀行名(取込)
'    eShitenCode     '支店コード
'    eShitenName_m   '支店名(マスター)
'    eShitenName_i   '支店名(取込)
'    eYokinShumoku     '口座種別
'    eKouzaBango     '口座番号
'    eYubinKigou     '郵便局:通帳記号
'    eYubinBango     '郵便局:通帳番号
'    eKouzaName      '口座名義人=>保護者名(カナ)
    
    If sprMeisai.MaxRows = 0 Then
        Exit Sub
    End If
    '//コマンド・ボタン制御
    Call pLockedControl(False)
    '//エラー列を設定
    ErrStts = Array("CIERROr", "CIITKBe", _
                    "CIKYCDe", "cikycde", "CIKSCDe", "CIHGCDe", "CIKJNMe", "CIKNNMe", "CISTNMe", "CISKGKe", _
                    "CIKKBNe", "CIBANKe", "cibanke", "CIBKNMe", "CISITNe", "cisitne", "CISINMe", "CIKZSBe", "CIKZNOe", _
                    "CIYBTKe", "CIYBTNe", _
                    "CIKZNMe" _
                )
    sql = "SELECT ROWNUM,a.* FROM(" & vbCrLf
    sql = sql & "SELECT CIINDT,CISEQN,CIMUPD," & mRimp.StatusColumns("," & vbCrLf, Len("," & vbCrLf))
    sql = sql & mMainSQL
    sql = sql & ") a"
#If ORA_DEBUG = 1 Then
    Set dyn = gdDBS.OpenRecordset(sql, dynOption.ORADYN_READONLY)
#Else
    Set dyn = gdDBS.OpenRecordset(sql, OracleConstantModule.ORADYN_READONLY)
#End If
    If False = vReset Then
        'SPread のスクロールバー押下時のみ開始行に移動
        Call dyn.FindFirst("ROWNUM >= " & sprMeisai.TopRow)
    End If
    mSpread.Redraw = False
    cnt = 0
    Do Until dyn.EOF
        '//処理が複雑になるのでとにかく２５行は読んでしまえ！
        'If 0 = dyn.Fields(ErrStts(0)) And "正常" = mSpread.Value(eSprCol.eErrorStts, dyn.RowPosition) Then
        '    Exit Do     '//異常、警告、正常のデータ順に並んでいるはずなので正常データが来たなら終了しても可！
        'End If
        cnt = cnt + 1
        If cnt > cVisibleRows Then    '//仮想モードなので２５行設定した時点で終了
            Exit Do
        End If
        For ix = LBound(ErrStts) To UBound(ErrStts)
            '//各列の表示色変更
            mSpread.BackColor(ix + 1, dyn.RowPosition) = mRimp.ErrorStatus(dyn.Fields(ErrStts(ix)))
        Next ix
        '//処理結果列の表示色
        '//2006/04/04 マスタ反映ＯＫフラグ判断
        If 0 <> Val(dyn.Fields("CIMUPD")) Then
            mSpread.BackColor(eSprCol.eErrorStts, dyn.RowPosition) = vbYellow
        ElseIf mRimp.ErrorStatus(0) = mSpread.BackColor(eSprCol.eErrorStts, dyn.RowPosition) Then
            mSpread.BackColor(eSprCol.eErrorStts, dyn.RowPosition) = vbCyan
        End If
        Call dyn.MoveNext
    Loop
    Call dyn.Close
    Set dyn = Nothing
    mSpread.Redraw = True
    '//コマンド・ボタン制御
    Call pLockedControl(True)
End Sub



