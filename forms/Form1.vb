
Imports System.Data.Entity
Imports System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder
Imports System.Data.SqlClient
Imports System.Data.SQLite
Imports System.Dynamic
Imports System.IO
Imports System.Net
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports Microsoft.VisualBasic.ApplicationServices
Imports Microsoft.Web.WebView2.Core
'Imports System.Net.Mime.MediaTypeNames
Imports Microsoft.Web.WebView2.WinForms
Imports stdole
Imports WindowsApplication1.Form1

Public Class Form1

    Public Sub New()
        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()
        ' InitializeComponent() 呼び出しの後で初期化を追加します。
    End Sub
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '(とりあえず動けばよいという事で、性能は考慮してない)不足データは追々足す
    '-----------------
    '研究所

    Public Structure Lab_Type
        Public Db_Type As String '利用するDB名　’MSSQL　SQLite

        Public Owner As String 'OWNER
        Public OwnerID As Long  'OWNER
        Public OwnerList As DataTable

        Public XOwner As String 'カレントOWNER
        Public XOwnerID As Long  'カレントOWNER

        Public ProjList As DataTable
        Public PID As Long  'カレントのRepoID

        Public ProjID As Long  'カレントのRepoID
        Public ProjName As String  'カレントのRepo名


        Public ProjKey As String
        Public ProjPath As String

        Public ProjOwnerID As String

        Public WorldX As Long  'デスクのX初期位置
        Public WorldY As Long  'デスクのY初期位置
        Public WorldW As Long  'デスクの幅
        Public WorldH As Long  'デスクの高さ

        Public Desk() As Desk_Type
        Public DeskList As DataTable
        Public DeskName As String  'カレントのデスク名
        Public DeskID As Long  'カレントのデスクID
        Public DeskIDShift As Integer  'デスクID表示のShift値
        Public DesKTop() As Panel
        Public DeskTopID As Integer  'カレントのDESKID
        Public DeskViewMax As Integer '   Public TabP As Panel 'カレントのパネル
        Public DeskXX As Integer 'カレントベクター
        Public DeskYY As Integer 'カレントベクター
        Public MvX As Integer 'カレントベクター
        Public MvY As Integer 'カレントベクター

        Public BoxList As DataTable
        Public BoxID As Long  'カレントのBOXID
        ' Public Box As Panel  'カレントのBASE
        Public Base As Panel  'カレントのBASE
        Public BluePrint As String  'カレントのBluePrint
        Public Line As Panel  'カレントのラインパネル
        Public LineID As Long  'カレントのラインID
        Public LineMax As Integer  'カレントの最大要素数
        Public LineNo As Integer  'カレントのクローズ要素数

        Public DbPath As String

        Public DBACS_Type As Dictionary(Of String, String)
        Public DBACS_Field As Dictionary(Of String, String)

        Public ProjAT As ProjAT_Type
        Public Proj() As Proj_Type
        Public Person() As Person_Type
        ' Public DeskAT As DeskAT_Type 'Desk


        '  Public Pbox() As Pbox_Type
        '  Public Tbox() As Tbox_Type

        Public StackAT As ElementAT_Type 'ノードの格納域RepoEleTbl_Type 

        Public ElementAT As ElementAT_Type 'ノードの格納域RepoEleTbl_Type 
        Public Element() As Element_Type

        'Public RoleAT As RoleAT_Type
        Public Role() As Role_Type

        Public Image_List As Image_List_Type
        Public ToolTip As ToolTip
        Public Flg_BoxMouse As String 'BOX上のマウスの状態フラグBOXに入ると１　マウスが動くと０
        Public Flg_DeskMove As Integer 'デスクを動かすと１、動いてなければ０
        Public Flg_DeskRename As Integer 'True ユーザによる変更
    End Structure
    Public Structure ProjAT_Type
        Public View As ListView_Type 'リスト表示
        Public Menu As ContextMenuStrip
    End Structure

    Public Structure Proj_Type
        Public ProjID As Long
        Public Title As String
        Public OwnerID As Long
        Public State As String
        Public Summary As String
        Public Hypothesis As String
        Public Purpose As String
        Public Problem As String


    End Structure
    '人事情報

    Public Structure PersonAT_Type
        Public DBACS As DBACS_Type     'データDBアクセスタイプ
        ' Public DBACS_ProjPerson As DBACS_Type


    End Structure
    Public Structure Person_Type
        Public ID As Long
        Public Name As String
        Public Prof As String
        Public Mail As String
        Public Tel As String
        Public Address As String
        Public Image As String 'URLファイルパスが本来だがこうしておく

    End Structure
    Public Structure RoleAT_Type
        Public DBACS As DBACS_Type     'データDBアクセスタイプ

    End Structure
    Public Structure Role_Type
        Public ID As Long
        Public LID As Integer
        Public PID As Integer
        Public Role As String

    End Structure
    'Role
    '筆頭著者	First Author	
    '第二執筆者	Second Author	
    '第三執筆者	Third Author	
    '執筆者	Author	
    '最終著者	Last Author	
    '責任著者	Corresponding Author	
    '研究責任者	Principal Investigator	
    '協力者	Collaborator	
    '研究責任者	Research Director
    '研究所長 Director Of Laboratory
    '研究所長 lab chief


    'Public Structure RepoDB_Type
    '    Public Db As DBACS_Type     'データDBアクセスタイプ


    'End Structure
    ''Base情報
    '--------------------------------
    Public Structure BaseAT_Type
        Public DBACS As DBACS_Type     'データDBアクセスタイプ

    End Structure

    ''エレメント情報
    '--------------------------------
    Public Structure ElementAT_Type

        Public View As ListView_Type

        Public Image As ImageList
        Public ImageID As Integer

    End Structure

    ' 'エレメント情報
    '--------------------------------
    Public Structure Element_Type
        Public ID As Long
        ' Public PageID As Integer
        '  Public Pagename As String
        Public Name As String
        Public Type As String
        Public Path As String

        Public Note As String
    End Structure


    Public Structure Image_List_Type
        Public Key() As String
        Public Image As ImageList
        Public Sub Int(ByVal Hmax)
            ReDim Preserve Key(Hmax)
            Image = New ImageList
        End Sub

    End Structure


    ''作業部屋情報
    ''--------------------------------
    'Public Structure DeskAT_Type
    '    'Public DBACS As DBACS_Type     'データDBアクセスタイプ

    '    Public View As TabControl '表示タイプ

    'End Structure

    '作業机情報
    '--------------------------------
    Public Structure Desk_Type

        Public DeskID As Long
        Public ProjID As Long
        Public ProjKey As String
        Public OwnerID As Long
        Public State As String
        Public Type As String 'Desk名　
        Public Name As String 'Desk名　
        Public Color As String '色　

        Public ShiftX As Integer 'DeskSIFT
        Public ShiftY As Integer 'DeskSIFT
        Public ZM As Integer 'Zoom
        Public DTop As Panel


        Public Box() As Box_Type
        '   Public BMax As Integer
    End Structure
    Public Structure Box_Type
        Public ID As Integer
        Public BoxID As Integer
        Public DeskID As Integer
        Public PageID As Integer
        Public ProjID As Long
        Public OwnerID As Long
        Public State As String
        Public ParentID As Integer
        Public LineID As Integer
        Public LineNo As Integer
        Public BluePrint As String
        Public Base As Panel

        Public Name As String
        Public Type As String

        '位置サイズ情報
        Public X As Integer
        Public Y As Integer
        Public W As Integer
        Public H As Integer
        Public LV As Integer

        'dependence情報
        Public Depend() As Depend_Type

    End Structure

    ' デペンデンシーdependency情報
    '--------------------------------
    Public Structure Depend_Type
        Public ID As Long
        Public Type As String
        Public State As String
        Public Name As String
        Public PathO As String 'オリジンパス
        Public PathC As String 'コピーパス
        Public PathV As String　'ビューパス

    End Structure





    Public Structure ListView_Type
        Public Name As String
        'コントロール
        Public ListView As System.Windows.Forms.ListView     'LabのList作業の対象のオブジェクト
        'カラム情報
        Public Columns As Columns_Type

    End Structure


    Public Structure RepoBlockType
        Public Name As String
        Public NodeList() As Long
    End Structure
    Public Structure Columns_Type
        Public Type As String
        Public DBKeys As String 'DBの戻り値
        Public Name() As String
        Public Width() As String　'文字列にしてあるので適用時に数値化すること
    End Structure

    'Public Structure Box_Type
    '    'Public Name As String
    '    'Public Type As String

    '    Public Code As String
    '    Public Dataset As DataSet
    '    Public WID As Long
    '    Public OwnerID As Long
    '    Public Base As Panel
    '    Public Split() As SplitContainer
    '    Public Pict() As PictureBox
    '    Public Text() As TextBox
    '    Public WebB() As WebView2
    '    Public Draw() As AxMSINKAUTLib.AxInkPicture
    '    Public CogT() As AxINKEDLib.AxInkEdit
    '    Public Combo() As ComboBox
    '    ' Public Tab() As TabControl
    '    Public Dtop() As Panel
    '    Public List() As ListView
    '    Public Btn() As Button



    'End Structure


    Public Structure Pbox_Type
        Public ID As Long
        Public BaseID As Long
        Public BoxType As String
        Public BoxPath As String
        Public Style As Integer
        Public PictureBox As PictureBox
        ' Public DbAcsKey As String 後で書き直す
        Public DBACS As DBACS_Type
    End Structure

    Public Structure Tbox_Type
        Public ID As Long
        Public BaseID As Long
        Public BoxType As String
        Public BoxPath As String
        Public Style As Integer
        Public TextBox As TextBox
        ' Public DbAcsKey As String 後で書き直す
        Public DBACS As DBACS_Type
    End Structure


    Public Structure Node_Type
        Public ID As Long
        Public State As Integer     'State
        Public NodeType As String
        Public Name As String
        Public text As String
        Public Owner As String
        Public OwnerID As Long
        Public PID As Integer     'ノードの 親ID
        Public RID As Integer     'レポジトリID

        'Public DT() As DicTbl_Type    'DicTbl DType単位
        'Public DS() As DicTbl_Type    'DicTbl Suggest情報 


        ' コピーを作成するメソッド
        Public Function Clone() As Node_Type
            Dim cloned As Node_Type = CType(MemberwiseClone(), Node_Type)
            ' 参照型フィールドの複製を作成する
            cloned.Name = CType(MyClass.Name.Clone, String)
            If MyClass.text IsNot Nothing Then cloned.text = CType(MyClass.text.Clone, String)
            'cloned.Path = CType(MyClass.Path.Clone, String)
            'If Not MyClass.DT Is Nothing Then cloned.DT = CType(MyClass.DT.Clone(), DicTbl_Type())
            'If Not MyClass.DS Is Nothing Then cloned.DS = CType(MyClass.DS.Clone(), DicTbl_Type())
            Return cloned
        End Function
    End Structure



    Public Structure DBACS_Type

        Public DB_Path As String    'DBへのパス
        Public DB_Name As String    'DBの名前

        Public TBL_Name As String
        Public TBL_Type As String
        Public TBL_Field As String


        ' コピーを作成するメソッド
        Public Function Clone() As DBACS_Type
            Dim cloned As DBACS_Type = CType(MemberwiseClone(), DBACS_Type)
            ' 参照型フィールドの複製を作成する
            cloned.DB_Path = CType(MyClass.DB_Path.Clone, String)
            cloned.DB_Name = CType(MyClass.DB_Name.Clone, String)

            cloned.TBL_Type = CType(MyClass.TBL_Type.Clone, String)
            cloned.TBL_Field = CType(MyClass.TBL_Field.Clone, String)
            cloned.TBL_Name = CType(MyClass.TBL_Name.Clone, String)

            Return cloned
        End Function

    End Structure










    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■




    Public Structure Deco_Type

        Public BackColor As String
        Public ForeColor As String
    End Structure



    Public Class ListViewItemComparer
        Implements IComparer
        Private _column As Integer

        Public Sub New(ByVal col As Integer)
            _column = col
        End Sub

        'xがyより小さいときはマイナスの数、大きいときはプラスの数、
        '同じときは0を返す
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            'ListViewItemの取得
            Dim TmpR As Integer
            Dim Text1, Text2 As String
            Dim itemx As ListViewItem = CType(x, ListViewItem)
            Dim itemy As ListViewItem = CType(y, ListViewItem)
            If itemx Is Nothing Or itemy Is Nothing Then Return 0
            If itemx.SubItems(_column).Text = "" Then Return 0
            If itemy.SubItems(_column).Text = "" Then Return 0

            'xとyを文字列として比較する
            Text1 = itemx.SubItems(_column).Text
            Text2 = itemy.SubItems(_column).Text

            TmpR = String.Compare(Text1, Text2)
            Return TmpR

        End Function
    End Class
    Public Class ListViewItemComparer_F
        Implements IComparer
        Private _column As Integer

        Public Sub New(ByVal col As Integer)
            _column = col
        End Sub

        'xがyより小さいときはマイナスの数、大きいときはプラスの数、
        '同じときは0を返す
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            'ListViewItemの取得
            Dim TmpR As Integer
            Dim Text1, Text2 As String
            Dim itemx As ListViewItem = CType(x, ListViewItem)
            Dim itemy As ListViewItem = CType(y, ListViewItem)
            If itemx Is Nothing Or itemy Is Nothing Then Return 0

            Try
                If itemx.SubItems(_column).Text = "" Then Return 0
                If itemy.SubItems(_column).Text = "" Then Return 0
                'xとyを文字列として比較する
                Text1 = itemx.SubItems(_column).Text
                Text2 = itemy.SubItems(_column).Text

                TmpR = String.Compare(Text1, Text2)
                Return TmpR

            Catch


            End Try
            Return 0
        End Function
    End Class

    Public Class ListViewItemComparer_L

        Implements IComparer
        Private _column As Integer

        Public Sub New(ByVal col As Integer)
            _column = col
        End Sub

        'xがyより小さいときはマイナスの数、大きいときはプラスの数、
        '同じときは0を返す
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            'ListViewItemの取得
            Dim TmpR As Integer
            Dim Text1, Text2 As String
            Dim itemx As ListViewItem = CType(x, ListViewItem)
            Dim itemy As ListViewItem = CType(y, ListViewItem)
            If itemx Is Nothing Or itemy Is Nothing Then Return 0
            If itemx.SubItems(_column).Text = "" Then Return 0
            If itemy.SubItems(_column).Text = "" Then Return 0

            'xとyを文字列として比較する
            Text1 = itemx.SubItems(_column).Text
            Text2 = itemy.SubItems(_column).Text

            TmpR = String.Compare(Text1, Text2)
            Return -TmpR

        End Function
    End Class

    Public Class ListViewItemComparer_BF
        Implements IComparer
        Private _column As Integer

        Public Sub New(ByVal col As Integer)
            _column = col
        End Sub

        'xがyより小さいときはマイナスの数、大きいときはプラスの数、
        '同じときは0を返す
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            'ListViewItemの取得
            Dim TmpR As Integer
            Dim Text1, Text2 As String
            Dim itemx As ListViewItem = CType(x, ListViewItem)
            Dim itemy As ListViewItem = CType(y, ListViewItem)
            If itemx Is Nothing Or itemy Is Nothing Then Return 0
            If itemx.SubItems(_column).Text = "" Then Return 0
            If itemy.SubItems(_column).Text = "" Then Return 0

            'xとyを文字列として比較する
            Text1 = StrReverse(itemx.SubItems(_column).Text)
            Text2 = StrReverse(itemy.SubItems(_column).Text)

            TmpR = String.Compare(Text1, Text2)
            Return TmpR

        End Function
    End Class

    Public Class ListViewItemComparer_BL
        Implements IComparer
        Private _column As Integer

        Public Sub New(ByVal col As Integer)
            _column = col
        End Sub

        'xがyより小さいときはマイナスの数、大きいときはプラスの数、
        '同じときは0を返す
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            'ListViewItemの取得
            Dim TmpR As Integer
            Dim Text1, Text2 As String
            Dim itemx As ListViewItem = CType(x, ListViewItem)
            Dim itemy As ListViewItem = CType(y, ListViewItem)
            If itemx Is Nothing Or itemy Is Nothing Then Return 0
            If itemx.SubItems(_column).Text = "" Then Return 0
            If itemy.SubItems(_column).Text = "" Then Return 0

            'xとyを文字列として比較する
            Text1 = StrReverse(itemx.SubItems(_column).Text)
            Text2 = StrReverse(itemy.SubItems(_column).Text)

            TmpR = String.Compare(Text1, Text2)
            Return -TmpR

        End Function
    End Class



    'グローバル
    '----------------------------------------------------
    '----------------------------------------------------
    '----------------------------------------------------

    Public gLab As Lab_Type         'Lab　グローバル

    Public gX_owner As String             '現在の作業者
    Public gX_ownerID As Long
    Public gX_Token As String

    Public gOwner As String             'LABの所有者
    Public gOwnerID As Long
    Public gDeco As Deco_Type          'デザインカラー
    Public gApp_Path As String          'EXE位置
    Public gIcon_Path As String         'Iconデータへのパス  
    Public gWORK_Path As String         '各種WORKデータへのパス  
    Public gENV_Path As String         '各種設定データへのパス  
    Public gLOG_Path As String         'ログ出力先のパス  
    Public gDB_Path As String
    Public gDB_Acs As String        'DBへのアクセスキー
    Public gImg_List As New Image_List_Type     'イメージリスト

    Public gSelectNode As TreeNode      '現在フォーカスしてるRepoNode 
    Public gView(5) As Integer  '画面表示制御

    Public gProjKey As String
    Public gProjPath As String
    Public gProjID As String
    Public gProjOwnerID As String 'カレントのレポジトリID  
    Public gBtnGrid As DataGridView
    Public gFontView As PictureBox
    Public gFontName As String
    Public gFontSize As Integer

    Public gColorName As String
    Public gColorView As RichTextBox
    Public gWorkPanel() As Panel
    Public gMap As Panel
    Public CatName() As String
    '----------------------------------------------------
    '----------------------------------------------------
    '----------------------------------------------------
    '----------------------------------------------------

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '●●


        Lab_Set_Env(gLab) '環境設定
        Handler_Set() ' 'イベントハンドラを追加する

        'OWNERの設定
        '-----------------------------------

        gX_ownerID = Owner_Get_ID(gLab)


        If gX_ownerID = 0 Then
            'ユーザ仮登録

            'gX_ownerID = 1
            'Owner_Save_ID(gX_ownerID)
            'gLab.OwnerID = gX_ownerID
            ''プロジェクト新設
            'Proj_Make_New(gLab)

        Else
            gLab.OwnerID = gX_ownerID
            '  Owner_Set_Env(gLab, gX_ownerID)
            'Owner_Show_Data(gLab)
        End If


        'スタックの一覧設定 現在は全部出すが本来はユーザ依存？
        Box_Show_List()

        Proj_Get_ProjList(gLab) 'Projectの一覧作成
        Dim Plist As DataRow()
        Plist = gLab.ProjList.Select("State = 'Nomal'")
        Proj_Show_List(Plist, True)



    End Sub
    Private Sub ColorGrid_Set(ByRef Workpanel As Panel)

        ' Dim gColorView As RichTextBox
        gColorView = New RichTextBox
        Workpanel.Controls.Add(gColorView)
        With gColorView
            .Multiline = False
            .Height = 23
            .SelectionAlignment = HorizontalAlignment.Center

            .Dock = DockStyle.Bottom
        End With


        Dim ColorGrid As DataGridView
        ColorGrid = New DataGridView
        Workpanel.Controls.Add(ColorGrid)



        Dim ColName() As String = Split("black|dimgray|gray|darkgray|silver|lightgray|gainsboro|whitesmoke|white|snow|ghostwhite|floralwhite|linen|antiquewhite|papayawhip|blanchedalmond|bisque|moccasin|navajowhite|peachpuff|mistyrose|lavenderblush|seashell|oldlace|ivory|honeydew|mintcream|azure|aliceblue|lavender|lightsteelblue|lightslategray|slategray|steelblue|royalblue|midnightblue|navy|darkblue|mediumblue|blue|dodgerblue|cornflowerblue|deepskyblue|lightskyblue|skyblue|lightblue|powderblue|paleturquoise|lightcyan|cyan|aqua|turquoise|mediumturquoise|darkturquoise|lightseagreen|cadetblue|darkcyan|teal|darkslategray|darkgreen|green|forestgreen|seagreen|mediumseagreen|mediumaquamarine|darkseagreen|aquamarine|palegreen|lightgreen|springgreen|mediumspringgreen|lawngreen|chartreuse|greenyellow|lime|limegreen|yellowgreen|darkolivegreen|olivedrab|olive|darkkhaki|palegoldenrod|cornsilk|beige|lightyellow|lightgoldenrodyellow|lemonchiffon|wheat|burlywood|tan|khaki|yellow|gold|orange|sandybrown|darkorange|goldenrod|peru|darkgoldenrod|chocolate|sienna|saddlebrown|maroon|darkred|brown|firebrick|indianred|rosybrown|darksalmon|lightcoral|salmon|lightsalmon|coral|tomato|orangered|red|crimson|mediumvioletred|deeppink|hotpink|palevioletred|pink|lightpink|thistle|magenta|fuchsia|violet|plum|orchid|mediumorchid|darkorchid|darkviolet|darkmagenta|purple|indigo|darkslateblue|blueviolet|mediumpurple|slateblue|mediumslateblue", "|")

        With ColorGrid
            .Dock = DockStyle.Fill
            .AllowDrop = True
            .AllowUserToAddRows = False
            .ColumnHeadersVisible = False
            .AllowUserToResizeRows = False
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            .AllowUserToResizeColumns = False
            '行ヘッダーを非表示にする
            .RowHeadersVisible = False
            .AllowUserToResizeRows = False
            ' .ReadOnly = True
            .ColumnCount = 5
            .RowCount = ColName.Count / .ColumnCount
            AddHandler .CellClick, AddressOf ColorGrid_Click
            AddHandler .MouseMove, AddressOf ColorGrid_MouseMove
            AddHandler .DragOver, AddressOf ColorGrid_DragOver
            '  AddHandler .DragDrop, AddressOf ColorGrid_DragDrop

            '  .RowTemplate.Height = 4
            For K = 0 To .ColumnCount - 1
                .Columns(K).Width = (.Width / .ColumnCount) - 3
                For L = 0 To .RowCount - 1
                    .Rows(L).Height = 8
                    .Rows(L).Cells(K).Style.BackColor = Color.FromName(ColName(L + K * (.RowCount - 1)))
                    .Rows(L).Cells(K).ToolTipText = ColName(L + K * (.RowCount - 1))

                Next

            Next

            .DefaultCellStyle.SelectionBackColor = Color.FromArgb(100, Color.White)


        End With

        Workpanel.BringToFront()

    End Sub
    Private Sub Font_MouseEnter(sender As Object, e As System.EventArgs)
        gLab.Flg_BoxMouse = "OnFont"
    End Sub
    Private Sub FontList_Set(ByRef Workpanel As Panel)

        'Dim Fontview As PictureBox
        gFontView = New PictureBox
        Workpanel.Controls.Add(gFontView)
        With gFontView
            .Dock = DockStyle.Fill
            .BackColor = Color.FromName("black")

            AddHandler .MouseEnter, AddressOf Font_MouseEnter
        End With

        Dim FontGrid As DataGridView
        FontGrid = New DataGridView
        Workpanel.Controls.Add(FontGrid)

        Dim Ifc As New System.Drawing.Text.InstalledFontCollection
        'インストールされているすべてのフォントファミリアを取得
        Dim Ffs As FontFamily() = Ifc.Families

        With FontGrid
            .Width = 120
            .Dock = DockStyle.Right
            .AllowUserToAddRows = False
            .ColumnHeadersVisible = False
            .AllowUserToResizeRows = False
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            .AllowUserToResizeColumns = False
            '行ヘッダーを非表示にする
            .RowHeadersVisible = False
            .AllowUserToResizeRows = False
            .ColumnCount = 1
            .RowCount = Ffs.Count
            AddHandler .CellContentClick, AddressOf FontGrid_Click



        End With
        Dim cnt As Integer
        Dim ff As FontFamily
        For Each ff In Ffs
            Dim fnt As New System.Drawing.Font(ff, 8)
            With FontGrid.Rows(cnt).Cells(0)
                .Value = fnt.Name
                .Style.Font = fnt
                .Style.BackColor = Color.FromName("black")
                .Style.ForeColor = Color.FromName("white")

            End With
            cnt += 1
        Next


    End Sub
    Private Sub FontGrid_Click(sender As Object, e As DataGridViewCellEventArgs)
        gFontName = sender(e.ColumnIndex, e.RowIndex).Value

        FontView_Show(gFontName, gFontSize)


    End Sub

    Private Sub FontView_Show(FontName As String, FontSize As Integer)
        Dim canvas As New Bitmap(gFontView.Width, gFontView.Height) 'ImageオブジェクトのGraphicsオブジェクトを作成する
        Dim g As Graphics = Graphics.FromImage(canvas) 'ImageオブジェクトのGraphicsオブジェクトを作成する

        'フォントオブジェクトの作成
        Dim fnt As New System.Drawing.Font(FontName, FontSize, FontStyle.Regular)
        'fnt = New System.Drawing.Font(FontName, 12, FontStyle.Regular)

        Dim drawBrush As SolidBrush
        drawBrush = New SolidBrush(Color.Black)

        Dim rect As New RectangleF(0, 0, 100, 200)
        g.FillRectangle(Brushes.Black, rect) 'rectの四角を描く

        Dim sf As New StringFormat()
        sf.FormatFlags = StringFormatFlags.DirectionVertical   ''縦書きにする
        g.DrawString(FontName, fnt, Brushes.White, rect, sf) '文字を書く

        fnt.Dispose() 'リソースを解放する
        g.Dispose()

        gFontView.Image = canvas    '表示する
    End Sub

    Private Sub ColorGrid_Click(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs)

        gColorView.Text = sender(e.ColumnIndex, e.RowIndex).ToolTipText

        gColorView.BackColor = sender(e.ColumnIndex, e.RowIndex).Style.BackColor
        gColorName = gColorView.Text
    End Sub
    Private Sub Map_Set_MapView(ByRef Workpanel As Panel)
        '.WorldX = -2000 'MAPデスクの初期X位置
        '.WorldY = -2000  'MAPデスクの初期Y位置
        '.WorldW = 24000  'MAPデスクの幅
        '.WorldH = 24000  'MAPデスクの高さ


        '

        gMap = New Panel
        Workpanel.Controls.Add(gMap)
        With gMap
            .BackColor = Color.FromName("royalblue")
            If gLab.DeskID = 0 Then
                .Left = Math.Abs(Int(gLab.WorldX) * Workpanel.Width * 4 / gLab.WorldW)
                .Top = Math.Abs(Int(gLab.WorldY) * Workpanel.Height * 4 / gLab.WorldH)
                .Left = Math.Abs(Int(gLab.WorldX) * Workpanel.Width * 4 / gLab.WorldW)
                .Top = Math.Abs(Int(gLab.WorldY) * Workpanel.Height * 4 / gLab.WorldH)

            Else
                .Left = Math.Abs(Int(gLab.WorldX - gLab.Desk(gLab.DeskID).ShiftX) * Workpanel.Width * 4 / gLab.WorldW)
                .Top = Math.Abs(Int(gLab.WorldY - gLab.Desk(gLab.DeskID).ShiftY) * Workpanel.Height * 4 / gLab.WorldH)
                .Left = Math.Abs(Int(gLab.WorldX - gLab.Desk(gLab.DeskID).ShiftX) * Workpanel.Width * 4 / gLab.WorldW)
                .Top = Math.Abs(Int(gLab.WorldY - gLab.Desk(gLab.DeskID).ShiftY) * Workpanel.Height * 4 / gLab.WorldH)
            End If



            .Width = 30 'Int(Workpanel.Width * 8 / gLab.WorldW)
            .Height = 15 'Int(Workpanel.Height * 8 / gLab.WorldH)
            AddHandler .MouseMove, AddressOf Map_MouseMove
            AddHandler .MouseDown, AddressOf Map_MouseDown
            AddHandler .MouseUp, AddressOf Map_Mouseup
        End With


    End Sub

    Private Sub Set_BtnList(ByRef Workpanel As Panel)


        gBtnGrid = New DataGridView
        Workpanel.Controls.Add(gBtnGrid)

        With gBtnGrid
            .Dock = DockStyle.Fill
            .AllowUserToAddRows = False
            .ColumnHeadersVisible = False
            .AllowUserToResizeRows = False
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            .AllowUserToResizeColumns = False

            .RowHeadersVisible = False '行ヘッダーを非表示にする
            .AllowUserToResizeRows = False
            ' .ReadOnly = True
            .ColumnCount = 3
            .RowCount = 1
            AddHandler .CellClick, AddressOf BtnGrid_Click

            AddHandler .CellMouseEnter, AddressOf BtnGrid_MouseEnter
            '  .RowTemplate.Height = 4
            For K = 0 To .ColumnCount - 1
                .Columns(K).Width = (.Width / .ColumnCount) - 3
                .Rows(0).Height = 8
                .Rows(0).Cells(K).Tag = K.ToString
            Next


        End With

    End Sub

    Private Sub BtnGrid_Click(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs)
        Dim PID As Integer
        PID = Val(sender(e.ColumnIndex, e.RowIndex).tag)

        gWorkPanel(PID).BringToFront()
        'gColorView.Text = sender(e.ColumnIndex, e.RowIndex).ToolTipText
        'gColorView.BackColor = sender(e.ColumnIndex, e.RowIndex).Style.BackColor
    End Sub

    Private Sub BtnGrid_MouseEnter(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs)
        sender.CurrentCell = sender(e.ColumnIndex, e.RowIndex)

        Dim PID As Integer
        PID = Val(sender(e.ColumnIndex, e.RowIndex).tag)

        gWorkPanel(PID).BringToFront()
        'gColorView.Text = sender(e.ColumnIndex, e.RowIndex).ToolTipText
        'gColorView.BackColor = sender(e.ColumnIndex, e.RowIndex).Style.BackColor
    End Sub
    Private Sub Set_WorkPanel()
        gFontSize = 8
        gFontName = "Arial"

        Dim WorkBasePanel As SplitContainer
        WorkBasePanel = New SplitContainer
        SplitContainer2.Panel2.Controls.Add(WorkBasePanel)
        With WorkBasePanel
            .Orientation = Orientation.Horizontal
            .Dock = DockStyle.Fill
            .BackColor = Color.FromName("black")
            .Panel2MinSize = 8
            .SplitterDistance = .Height - 10
            '.Panel2.Height = 10
            .Panel2.BackColor = Color.FromName("red")
            .FixedPanel = FixedPanel.Panel2
        End With


        ReDim gWorkPanel(3)
        For L = 0 To 2
            gWorkPanel(L) = New Panel
            WorkBasePanel.Panel1.Controls.Add(gWorkPanel(L))
            With gWorkPanel(L)
                .BackColor = Color.FromName("black")
                .Dock = DockStyle.Fill

            End With
        Next

        'Dim ColorGrid As DataGridView
        'Dim ColorView As TextBox
        Set_BtnList(WorkBasePanel.Panel2)
        Map_Set_MapView(gWorkPanel(0))
        ColorGrid_Set(gWorkPanel(1))
        FontList_Set(gWorkPanel(2))
        gBtnGrid.CurrentCell = gBtnGrid.Rows(0).Cells(0)
        gWorkPanel(0).BringToFront()

    End Sub

    Function Person_Set_NewOwner() As Long

        Lab_Start(gLab)



    End Function


    Private Sub Proj_Show_ListByCmd(ByRef Lab As Lab_Type, Cmd As String)
        '研究リストを表示
        '------------------
        Dim Plist As DataRow()
        With Lab
            ' DS = DBAcs_Get_DataS(Lab, "Proj", .ProjAT.View.Columns.DBKeys, Cmd)

            Plist = .ProjList.Select(Cmd)

            Button3.Text = "Hits　　" & Plist.Count.ToString
            Proj_Show_List(Plist, True)
        End With

    End Sub

    Private Sub Lab_Start(ByRef Lab As Lab_Type)
        '  Proj_Show_ListByCmd(gLab, " OwnerID = " & gX_ownerID) 'オーナーのレポ表示
        '    ListView1.Items(ListView1.Items.Count - 1).Selected = True '(本来なら終了する前のカレントノードだが今はさぼる)

        'gSelectNode = Lab.Desk(1).Node_Opened
        ''Proj_Set_Env(Lab) 'Proj環境設定
        'Proj_Read_Dat(Lab, gOwnerID) 'Projの読み込み
        ''Desk_Set_Env(Lab) 'Desk環境設定
        ''Desk_Read_NodeData(Lab) 'Desk読み込み
        ''Desk_Show_NodeAll(Lab) 'Desk表示
        'Element_Set_Env(Lab) 'ElementDesk環境設定


    End Sub


    Sub Handler_Set()

        'イベントハンドラを追加する
        '
        ''ListView1
        AddHandler ListView1.ItemDrag, AddressOf ListView1_ItemDrag
        'AddHandler ListView1.DragOver, AddressOf ListView1_DragOver
        'AddHandler ListView1.DragDrop, AddressOf ListView1_DragDrop

        ''ListView2
        AddHandler ListView2.ItemDrag, AddressOf ListView2_ItemDrag
        AddHandler ListView2.DragOver, AddressOf ListView2_DragOver
        AddHandler ListView2.DragDrop, AddressOf ListView2_DragDrop

        ''ListView3
        AddHandler ListView3.ItemDrag, AddressOf ListView2_ItemDrag
        AddHandler ListView3.DragOver, AddressOf ListView2_DragOver
        AddHandler ListView3.DragDrop, AddressOf ListView2_DragDrop
        'AddHandler ListView6.ItemDrag, AddressOf ListView6_ItemDrag

        'AddHandler TabControl3.SelectionChanged, AddressOf TabControl1_SelectionChanged

        'ColumnClickイベントハンドラの追加
        AddHandler ListView1.ColumnClick, AddressOf ListView1_ColumnClick
        AddHandler ListView2.ColumnClick, AddressOf ListView1_ColumnClick

        'TabControl1.SelectTab(1)


        'AddHandler RichTextBox1.LostFocus, AddressOf RichTextBox1_LostFocus
        'AddHandler RichTextBox2.LostFocus, AddressOf RichTextBox2_LostFocus


        ' AddHandler TabControl1.SelectionChanged, AddressOf TabControl1_SelectionChanged

        'AddHandler PictureBox2.Paint, New PaintEventHandler(AddressOf PictureBox2_Paint)
        'New PaintEventHandler()
        '  ProgressBar1.Controls.Add(Info2)

        'Label1の位置をPictureBox1内の位置に変更する
        'Info2.Top = Info2.Top - ProgressBar1.Top
        'Info2.Left = Info2.Left - ProgressBar1.Left

        'Info2.Top = ProgressBar1.Top + 10
        'Info2.Left = ProgressBar1.Left

    End Sub
    Private Sub Proj_Save_Text(Lab As Lab_Type, RID As Long, Type As String, Text As String)

        'With Lab.ProjAT.DBACS
        '    DBAcs_Update_Data_ByKey(.Db, .Db.TBL_Name, RID, Type & " = '" & Text & "'")
        'End With

    End Sub






    'Private Sub TabControl1_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs)
    '    '選択されたタブの番号を取得
    '    Dim selectedIndex As Integer = TabControl1.SelectedIndex + 1

    '    '選択されたタブの情報を表示
    '    TextBlock1.Text = selectedIndex.ToString() + "番目のタブが選択されました"
    'End Sub


    Private Sub ListView1_ColumnClick(ByVal sender As Object, ByVal e As ColumnClickEventArgs)
        'ListViewItemSorterを指定する

        If e.Column <> 1 And e.Column <> 2 Then Return

        Dim Key, Type As String

        With ListView1



            Key = Mid(.Columns(e.Column).Text, 2)
            Type = Mid(.Columns(e.Column).Text, 1, 1)
            If System.Windows.Forms.Control.ModifierKeys = Keys.Control Then

                Select Case Type
                    Case "▲"
                        Type = "▼"
                    Case "▼"
                        Type = "▲"
                    Case "▽"
                        Type = "▲"
                    Case "△"
                        Type = "▼"
                    Case Else
                        Type = "▲"
                        Key = .Columns(e.Column).Text
                End Select
            Else
                Select Case Type
                    Case "▲"
                        Type = "▽"
                    Case "▼"
                        Type = "▲"
                    Case "▽"
                        Type = "△"
                    Case "△"
                        Type = "▽"
                    Case Else
                        Type = "△"
                        Key = .Columns(e.Column).Text
                End Select
            End If

            .Columns(e.Column).Text = Type & Key

            '並び替える（ListViewItemSorterを設定するとSortが自動的に呼び出される）
            Select Case Type
                Case "▽"
                    .ListViewItemSorter = New ListViewItemComparer_F(e.Column)

                Case "△"
                    .ListViewItemSorter = New ListViewItemComparer_L(e.Column)
                Case "▲"
                    .ListViewItemSorter = New ListViewItemComparer_BF(e.Column)

                Case "▼"
                    .ListViewItemSorter = New ListViewItemComparer_BL(e.Column)

            End Select

        End With

    End Sub



    Private Sub ListView6_ItemDrag(ByVal sender As Object, ByVal e As ItemDragEventArgs)
        'ノードのドラッグを開始
        '--------------------------------
        Dim Lv As ListView = CType(sender, ListView)
        'Lv.SelectedNode = CType(e.Item, ListView)
        'Lv.Focus()

        'ノードのドラッグを開始する
        '  Dim dde As DragDropEffects = Lv.DoDragDrop(e.Item, DragDropEffects.All)

        ' 複数項目選択？
        If Lv.SelectedItems.Count >= 2 Then
            DoDragDrop(Lv.SelectedItems, DragDropEffects.All)
        Else
            DoDragDrop(e.Item, DragDropEffects.All)
        End If


    End Sub
    Private Sub ListView5_ItemDrag(ByVal sender As Object, ByVal e As ItemDragEventArgs)
        'ノードのドラッグを開始
        '--------------------------------
        Dim Lv As ListView = CType(sender, ListView)
        'Lv.SelectedNode = CType(e.Item, ListView)
        'Lv.Focus()

        'ノードのドラッグを開始する
        '  Dim dde As DragDropEffects = Lv.DoDragDrop(e.Item, DragDropEffects.All)

        ' 複数項目選択？
        If Lv.SelectedItems.Count >= 2 Then
            DoDragDrop(Lv.SelectedItems, DragDropEffects.All)
        Else
            DoDragDrop(e.Item, DragDropEffects.All)
        End If



        ''移動した時は、ドラッグしたノードを削除する
        'If (dde And DragDropEffects.Move) = DragDropEffects.Move Then
        '    Lv.Nodes.Remove(CType(e.Item, TreeNode))
        'End If
    End Sub

    Private Sub ListView1_ItemDrag(ByVal sender As Object, ByVal e As ItemDragEventArgs)
        'ノードのドラッグを開始
        '--------------------------------
        Dim Lv As ListView = CType(sender, ListView)
        'Lv.SelectedNode = CType(e.Item, ListView)
        'Lv.Focus()

        'ノードのドラッグを開始する
        '  Dim dde As DragDropEffects = Lv.DoDragDrop(e.Item, DragDropEffects.All)

        ' 複数項目選択？
        If Lv.SelectedItems.Count >= 2 Then
            DoDragDrop(Lv.SelectedItems, DragDropEffects.All)
        Else
            DoDragDrop(e.Item, DragDropEffects.All)
        End If



        ''移動した時は、ドラッグしたノードを削除する
        'If (dde And DragDropEffects.Move) = DragDropEffects.Move Then
        '    Lv.Nodes.Remove(CType(e.Item, TreeNode))
        'End If
    End Sub

    Private Sub ListView1_DragOver(ByVal sender As Object, ByVal e As DragEventArgs)
        'ドラッグしている時
        '--------------------------------
        ListView2.Focus()
        '   ListView_DropOvwer_Check(sender, e)
    End Sub


    Private Sub ListView1_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs)
        'ドロップされたとき
        '------------------------------------
        Proj_Drop_At_ListView(sender, e)

    End Sub


    Private Sub ListView2_ItemDrag(ByVal sender As Object, ByVal e As ItemDragEventArgs)
        'ノードのドラッグを開始
        '--------------------------------
        Dim Lv As ListView = CType(sender, ListView)
        'Lv.SelectedNode = CType(e.Item, ListView)
        'Lv.Focus()

        'ノードのドラッグを開始する
        Dim dde As DragDropEffects = Lv.DoDragDrop(e.Item, DragDropEffects.All)

        ''移動した時は、ドラッグしたノードを削除する
        'If (dde And DragDropEffects.Move) = DragDropEffects.Move Then
        '    Lv.Nodes.Remove(CType(e.Item, TreeNode))
        'End If
    End Sub
    Private Sub Panel1_DragOver(ByVal sender As Object, ByVal e As DragEventArgs)
        'ドラッグしている時
        '--------------------------------
        ListView2.Focus()
        e.Effect = DragDropEffects.Move

    End Sub

    'Private Sub Panel1_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs)
    '    'ドロップされたとき
    '    '------------------------------------
    '    Dim Dat As VariantType
    '    Dim Drops() As String
    '    Dim Element() As Element_Type
    '    Dim cnt As Integer
    '    cnt = 0
    '    ReDim Preserve Element(cnt)

    '    ' FILE
    '    Drops = e.Data.GetData(DataFormats.FileDrop, False)
    '    If Not (Drops Is Nothing) Then

    '        For L = 0 To Drops.Length - 1
    '            Element(cnt).Path = Drops(L)
    '            Element(cnt).Name = System.IO.Path.GetFileNameWithoutExtension(Drops(L))
    '            Element(cnt).Type = "File"
    '            cnt = cnt + 1
    '            ReDim Preserve Element(cnt)
    '        Next

    '    End If

    '    '' Html
    '    'Drops(0) = e.Data.GetData(DataFormats.Html, False)
    '    'If Not (Drops Is Nothing) Then
    '    '    For L = 0 To Drops.Length - 1
    '    '        With Element(cnt)
    '    '            .Path = Drops(L)
    '    '            .Type = "Html"
    '    '            .Name = Mid(.Path, 1, 20) 'タイトルから取るべきだが
    '    '        End With

    '    '        cnt = cnt + 1
    '    '        ReDim Preserve Element(cnt)
    '    '    Next
    '    'End If


    '    ' TEXT
    '    ReDim Drops(0)
    '    Drops(0) = e.Data.GetData(DataFormats.Text, False)
    '    If Not (Drops(0) = "") Then
    '        For L = 0 To Drops.Length - 1
    '            With Element(cnt)
    '                .Path = Drops(L)
    '                If InStr(.Path, "http") > 0 Or InStr(.Path, ".htm") > 0 Then
    '                    .Type = "Url"
    '                    .Name = Mid(.Path, 1, 20) 'タイトルから取るべきだが
    '                Else
    '                    .Type = "Text"
    '                    .Name = Mid(.Path, 1, 20)
    '                End If
    '            End With

    '            cnt = cnt + 1
    '            ReDim Preserve Element(cnt)
    '        Next
    '    End If

    '    ' Bitmap
    '    Dim img As System.Drawing.Image
    '    img = e.Data.GetData(DataFormats.Bitmap, False)
    '    If Not (img Is Nothing) Then
    '        For L = 0 To Drops.Length - 1
    '            With Element(cnt)
    '                .Path = Drops(L)
    '                .Type = "Bitmap"
    '                .Name = "画像"
    '            End With

    '            cnt = cnt + 1
    '            ReDim Preserve Element(cnt)
    '        Next
    '    End If
    '    Element_Add_Data(gLab, gProjID, Element)
    '    'Proj_Action_Files("NOT", DropS)
    '    ' Proj_Drop_At_ListView(sender, e)


    'End Sub
    Private Sub Base_DragOver(ByVal sender As Object, ByVal e As DragEventArgs)
        'ドラッグしている時
        '--------------------------------
        e.Effect = DragDropEffects.Move

    End Sub

    Private Sub Base_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs)
        'ドロップされたとき
        '------------------------------------
        Dim cp As Point = sender.PointToClient(New Point(e.X, e.Y))
        Elements_Drop(e)
        gLab.Flg_BoxMouse = "OnBOX"

    End Sub
    Private Sub Text_KeyUp(sender As Object, e As KeyEventArgs)
        '文字数でサイズ調整
        '------------------------------------
        ''Dim textbox As TextBox = CType(sender, TextBox)
        'Dim cnt As Integer
        'For K = 0 To 100
        '    If textbox.GetFirstCharIndexFromLine(K) < 0 Then
        '        Exit For
        '    End If
        '    cnt += 1
        'Next
        'Dim proposedSize = New Size(100, 50)
        'Dim SZ As Size
        'SZ = TextRenderer.MeasureText("ABC", sender.Font, proposedSize)
        'sender.Height = SZ.Height * (cnt + 1)
        'gLab.Base.Height = sender.Height + 8
    End Sub
    Private Sub Text_Drop(ByVal sender As Object, ByVal e As DragEventArgs)
        'ドロップされたとき
        '------------------------------------
        Dim cp As Point = sender.PointToClient(New Point(e.X, e.Y))
        Elements_Drop(e)
        gLab.Flg_BoxMouse = "OnBOX"

    End Sub
    Private Sub ColorGrid_DragOver(ByVal sender As Object, ByVal e As DragEventArgs)
        'ドラッグしている時
        '--------------------------------
        sender.Focus()
        e.Effect = DragDropEffects.Move

    End Sub

    Private Sub ListView2_DragOver(ByVal sender As Object, ByVal e As DragEventArgs)
        'ドラッグしている時
        '--------------------------------
        ListView2.Focus()
        e.Effect = DragDropEffects.Move

    End Sub


    Private Sub ListView2_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs)
        'ドロップされたとき
        '------------------------------------
        Dim cp As Point = sender.PointToClient(New Point(e.X, e.Y))
        Elements_Drop(e)


    End Sub
    Private Function Elements_Drop(ByRef e As DragEventArgs) As Element_Type()
        'ドロップされたとき
        '------------------------------------
        Dim Dat As VariantType
        Dim Drops() As String
        Dim Element() As Element_Type
        Dim cnt As Integer
        cnt = 0
        ReDim Preserve Element(cnt)

        ' FILE
        Drops = e.Data.GetData(DataFormats.FileDrop, False)
        If Not (Drops Is Nothing) Then

            For L = 0 To Drops.Length - 1
                Element(cnt).Path = Drops(L)
                Element(cnt).Name = System.IO.Path.GetFileNameWithoutExtension(Drops(L))
                Element(cnt).Type = "File"

                '  If flag = 1 Then Box_MakeFromElement(gLab, "FILE", cp, Element(cnt)) 'DESK上
                cnt = cnt + 1
                ReDim Preserve Element(cnt)
            Next


        End If

        '' Html
        'Drops(0) = e.Data.GetData(DataFormats.Html, False)
        'If Not (Drops Is Nothing) Then
        '    For L = 0 To Drops.Length - 1
        '        With Element(cnt)
        '            .Path = Drops(L)
        '            .Type = "Html"
        '            .Name = Mid(.Path, 1, 20) 'タイトルから取るべきだが
        '        End With

        '        cnt = cnt + 1
        '        ReDim Preserve Element(cnt)
        '    Next
        'End If


        ' TEXT
        ReDim Drops(0)
        Drops(0) = e.Data.GetData(DataFormats.Text, False)
        If Not (Drops(0) = "") Then
            For L = 0 To Drops.Length - 1
                With Element(cnt)
                    .Path = Drops(L)
                    If InStr(.Path, "http") > 0 Or InStr(.Path, ".htm") > 0 Then
                        .Type = "Url"
                        .Name = Mid(.Path, 1, 20) 'タイトルから取るべきだが
                    Else
                        .Type = "Text"
                        .Name = Mid(.Path, 1, 20)
                    End If
                End With

                cnt = cnt + 1
                ReDim Preserve Element(cnt)
            Next

        End If

        ' Bitmap
        Dim img As System.Drawing.Image
        img = e.Data.GetData(DataFormats.Bitmap, False)
        If Not (img Is Nothing) Then
            For L = 0 To Drops.Length - 1
                With Element(cnt)
                    .Path = Drops(L)
                    .Type = "Bitmap"
                    .Name = "画像"
                End With

                cnt = cnt + 1
                ReDim Preserve Element(cnt)
            Next
        End If
        ' Element_Add_Data(gLab, gProjID, Element) 'DB登録
        Return Element

    End Function

    Private Function Elements_Paste() As Element_Type()
        'pasteされたとき
        '------------------------------------

        Dim Drops() As String
        Dim Element() As Element_Type
        Dim cnt As Integer
        cnt = 0
        ReDim Preserve Element(cnt)

        'クリップボードのデータチェック
        Dim data As IDataObject = Clipboard.GetDataObject()
        If data Is Nothing Then Return Nothing　'空


        With My.Computer.Clipboard
            '◆ここから

            'テキスト
            'If .ContainsText() Then
            '    Dim Dat() As String
            '    ReDim Dat(0)
            '    Dat(0) = .GetText()

            '    If Not (Dat(0) = "") Then
            '        For L = 0 To Dat.Length - 1
            '            With Element(cnt)
            '                .Path = Dat(L)
            '                If InStr(.Path, "http") > 0 Or InStr(.Path, ".htm") > 0 Then
            '                    .Type = "Url"
            '                    .Name = Mid(.Path, 1, 20) 'タイトルから取るべきだが
            '                Else
            '                    If InStr(.Path, "【すたっく】") > 0 Then
            '                        .Type = "Stack"
            '                    Else
            '                        .Type = "Text"
            '                        .Name = Mid(.Path, 1, 20)
            '                    End If
            '                End If
            '            End With

            '            cnt = cnt + 1
            '            ReDim Preserve Element(cnt)
            '        Next

            '    End If

            'End If

            'ファイル
            If .ContainsFileDropList() Then
                Dim Dat As System.Collections.Specialized.StringCollection
                Dat = .GetFileDropList()
                For Each fileName In Dat
                    Element(cnt).Path = fileName
                    Element(cnt).Name = System.IO.Path.GetFileNameWithoutExtension(fileName)
                    Element(cnt).Type = "File"
                    cnt = cnt + 1
                    ReDim Preserve Element(cnt)
                Next fileName
            End If

            ''画像
            'If .ContainsImage() Then
            '    Dim Dat As System.Drawing.Image
            '    Dat = .GetImage()
            '    PictureBox1.Image = Dat
            'End If

            ''音声
            'If .ContainsAudio() Then

            'End If
        End With


        Element_Add_Data(gLab, gProjID, Element)
        Return Element

    End Function




    Private Sub ListView3_ItemDrag(ByVal sender As Object, ByVal e As ItemDragEventArgs)
        'ノードのドラッグを開始
        '--------------------------------
        Dim Lv As ListView = CType(sender, ListView)
        'Lv.SelectedNode = CType(e.Item, ListView)
        'Lv.Focus()

        'ノードのドラッグを開始する
        '  Dim dde As DragDropEffects = Lv.DoDragDrop(e.Item, DragDropEffects.All)

        ' 複数項目選択？
        If Lv.SelectedItems.Count >= 2 Then
            DoDragDrop(Lv.SelectedItems, DragDropEffects.All)
        Else
            DoDragDrop(e.Item, DragDropEffects.All)
        End If



        ''移動した時は、ドラッグしたノードを削除する
        'If (dde And DragDropEffects.Move) = DragDropEffects.Move Then
        '    Lv.Nodes.Remove(CType(e.Item, TreeNode))
        'End If
    End Sub





    'Private Sub Proj_Drop_At_TreeView(ByVal sender As Object, ByVal e As DragEventArgs)
    '    'TreeViewにドロップしたとき
    '    '------------------------------------
    '    'Proj_Actで始まる関数は、基本的にHISTORYに記憶して、進む、戻るの対象になるもの


    '    '--------------------
    '    'タイプ別アクション
    '    '--------------------
    '    Dim Act_Type As String
    '    Dim Act_Effect As String
    '    Dim Type_T(2), Type_N(2), Type_S(2) As String

    '    Act_Type = ""
    '    'Action設定
    '    Act_Type = ""
    '    Select Case e.Effect
    '        Case 1
    '            Act_Type = "COPY"
    '        Case 2
    '            Act_Type = "MOVE"
    '    End Select

    '    Act_Effect = ""


    '    '--------------------
    '    'ドロップ先のTreeView
    '    '--------------------

    '    Dim Target_Tree As TreeView = CType(sender, TreeView)
    '    'ドロップ先のTreeNodeを取得する
    '    Dim Target_Node As TreeNode = Target_Tree.GetNodeAt(Target_Tree.PointToClient(New Point(e.X, e.Y)))



    '    '-----------------
    '    'ドロップチェック
    '    '-----------------

    '    '------------------
    '    'ドロップ元がTreeNode
    '    '------------------
    '    If e.Data.GetDataPresent(GetType(TreeNode)) Then
    '        e.Effect = DragDropEffects.None

    '        'ドロップされたデータ(TreeNode)を取得
    '        Dim Source_Node As TreeNode = CType(e.Data.GetData(GetType(TreeNode)), TreeNode)

    '        'それぞれのアクションに振り分け
    '        'Act_Effect = Proj_Drop_Select_TreeViewToTreeView(gLab, Source_Node, Target_Node, Act_Type)

    '        Select Case Act_Effect
    '            Case "Non"
    '                e.Effect = DragDropEffects.None
    '            Case "Move" '元ノードを削除
    '                e.Effect = DragDropEffects.Move
    '        End Select

    '        Return
    '    End If


    '    '------------------
    '    'ドロップ元がListViewItem
    '    '------------------
    '    Dim Dmax As Integer
    '    Dim Source() As ListViewItem
    '    'ドロップされたデータがListViewItemか調べる
    '    Dmax = 0
    '    If e.Data.GetDataPresent(GetType(ListViewItem)) Then
    '        ReDim Source(Dmax)
    '        Source(Dmax) = CType(e.Data.GetData(GetType(ListViewItem)), ListViewItem)

    '        'それぞれのアクションに振り分け
    '        Proj_Drop_Select_ListViewToTreeView(gLab, Source, Target_Node, Act_Type)


    '    End If

    '    'ドロップされたデータがListViewItem(複数)か調べる
    '    If e.Data.GetDataPresent(GetType(ListView.SelectedListViewItemCollection)) Then
    '        Dim SourceS As ListView.SelectedListViewItemCollection = CType(e.Data.GetData(GetType(ListView.SelectedListViewItemCollection)), ListView.SelectedListViewItemCollection)
    '        If Not (SourceS Is Nothing) Then

    '            For L = 0 To SourceS.Count - 1
    '                ReDim Preserve Source(Dmax)
    '                Source(Dmax) = SourceS.Item(L)
    '                Dmax += 1
    '            Next

    '            'それぞれのアクションに振り分け
    '            Proj_Drop_Select_ListViewToTreeView(gLab, Source, Target_Node, Act_Type)

    '        End If
    '    End If


    'End Sub


    Private Sub Proj_Drop_At_ListView(ByVal sender As Object, ByVal e As DragEventArgs)
        'ListViewにドロップしたとき
        '------------------------------------      
        '現在はなにもしない


    End Sub

    Private Sub Desk_Paste(cp As Point)
        ' Deskにペーストしたとき
        '----------------------------------



        Dim DeskID As Integer
        DeskID = gLab.DeskID


        Dim Elm() As Element_Type
        Elm = Elements_Paste() 'DBに登録
        If Elm IsNot Nothing Then Desk_Drops_Show(gLab, Elm, DeskID, cp)



        Desk_Drops_Show(gLab, Elm, gLab.DeskID, cp)

    End Sub
    Private Sub Desk_Drops(ByVal sender As Object, ByVal e As DragEventArgs)
        ' Deskにドロップしたとき
        '------------------------------------

        '-----------------
        'ドロップチェック
        '-----------------
        Dim DropType As String
        'ドロップされたデータが何か調べる
        If e.Data.GetDataPresent(GetType(ListViewItem)) Then DropType = "Box"
        If e.Data.GetDataPresent(GetType(ListViewItem)) Then

        End If

        Dim cp As Point = sender.PointToClient(New Point(e.X, e.Y))
        With gLab
            .DeskID = Val(sender.tag)
            .DeskTopID = Val(sender.name)
        End With

        '------------------
        'ドロップ元がListViewItem スタック
        '------------------
        Dim Dmax As Integer
        Dim Source() As ListViewItem
        Dmax = 0

        If e.Data.GetDataPresent(GetType(ListViewItem)) Then
            ReDim Source(Dmax)
            Source(Dmax) = CType(e.Data.GetData(GetType(ListViewItem)), ListViewItem) '複数であっても１つだけ

            With gLab
                .DeskID = gLab.DeskID
                .DeskName = gLab.DeskID.ToString
                .BluePrint = Box_Get_BluePrint(gLab, Source(0).Tag) '設計
                .BluePrint = Replace(.BluePrint, "[BX]", (cp.X + gLab.WorldX).ToString)
                .BluePrint = Replace(.BluePrint, "[BY]", (cp.Y + gLab.WorldY).ToString)
                Box_Makes_OnDESK(gLab, .BluePrint)
            End With
            Return

        End If
        '------------------
        'ドロップ元がListViewItem ELEMENT
        '------------------


        If e.Data.GetDataPresent(GetType(String)) Then
            If gLab.Flg_BoxMouse <> "OnBOX" Then
                Dim tmpP As Integer = InStr(CType(e.Data.GetData(GetType(String)), String), "[Col]")
                If tmpP > 0 Then
                    '------------------
                    'ドロップ元がcolorGrid
                    '------------------
                    Desk_Set_Color(gLab, gLab.DeskID, Mid(CType(e.Data.GetData(GetType(String)), String), tmpP + 5))

                    Return
                End If
            End If
        End If



        '------------------
        'ドロップ元がFontGrid
        '------------------
        '------------------
        'ドロップ元が外部　 ELEMENT
        '------------------
        Dim Elm() As Element_Type
        Elm = Elements_Drop(e) 'タイプ別に分類
        If Elm IsNot Nothing Then Desk_Drops_Show(gLab, Elm, gLab.DeskID, cp)


    End Sub
    Private Sub Desk_Drops_Show(ByRef lab As Lab_Type, Elm() As Element_Type, DeskID As Integer, cp As Point)
        For L = 0 To Elm.Count - 1
            With Elm(L)
                If .Name <> Nothing Then
                    Select Case .Type
                        Case "File"

                            With gLab
                                .DeskID = DeskID

                                .BluePrint = Box_Get_BluePrint(gLab, "Viewer") '設計★TYPE|★PATH|★NAME"
                                .BluePrint = Replace(.BluePrint, "[BX]", (cp.X + lab.WorldX).ToString)
                                .BluePrint = Replace(.BluePrint, "[BY]", (cp.Y + lab.WorldY).ToString)
                                .BluePrint = Replace(.BluePrint, "★TYPE", Elm(L).Type)
                                .BluePrint = Replace(.BluePrint, "★PATH", Elm(L).Path)
                                .BluePrint = Replace(.BluePrint, "★NAME", Elm(L).Name)
                                Box_Makes_OnDESK(gLab, .BluePrint)
                            End With
                        Case "Url"
                            With gLab
                                .DeskID = DeskID
                                .DeskName = DeskID.ToString
                                .BluePrint = Box_Get_BluePrint(gLab, "Viewer") '設計★TYPE|★PATH|★NAME"
                                .BluePrint = Replace(.BluePrint, "[BX]", (cp.X + lab.WorldX).ToString)
                                .BluePrint = Replace(.BluePrint, "[BY]", (cp.Y + lab.WorldY).ToString)
                                .BluePrint = Replace(.BluePrint, "★TYPE", Elm(L).Type)
                                .BluePrint = Replace(.BluePrint, "★PATH", Elm(L).Path)
                                .BluePrint = Replace(.BluePrint, "★NAME", Elm(L).Name)
                                Box_Makes_OnDESK(gLab, .BluePrint)
                            End With

                        Case "Text"
                            '"@TEXT|3|BASE|TBLR|0|8|200|80|-1|lemonchiffon|black|Arial|8|"

                            With gLab
                                .DeskID = DeskID
                                .DeskName = DeskID.ToString
                                .BluePrint = Box_Get_BluePrint(gLab, "Post_it") '設計★TYPE|★PATH|★NAME"
                                .BluePrint = Replace(.BluePrint, "[BX]", (cp.X + lab.WorldX).ToString)
                                .BluePrint = Replace(.BluePrint, "[BY]", (cp.Y + lab.WorldY).ToString)
                                .BluePrint &= Elm(L).Path
                                ' .BoxID = Box_SaveDB(gLab, .BluePrint) 'DBに格納

                                '文字列記録場所を確保
                                Dim DPath As String = Box_Get_DependPath(lab.BoxList, lab.BoxID, "3", "dat.txt")
                                System.IO.File.WriteAllText(DPath, Elm(L).Path)
                                Box_Factory(gLab, .BoxID, .BluePrint)
                            End With

                        Case "Bitmap"
                    End Select

                End If

            End With


        Next
    End Sub

    Private Function BxCogT_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As AxINKEDLib.AxInkEdit
        '@CogT|0|NEXT|Dock|TEXT|"
        Dim CogT As AxINKEDLib.AxInkEdit
        CogT = New AxINKEDLib.AxInkEdit

        'AddHandler TBox.DragEnter, AddressOf BxText_DragEnter
        'AddHandler TBox.DragDrop, AddressOf BxText_Drop

        OBJ.Controls.Add(CogT)
        With CogT
            .Dock = STK_PRT(3)

        End With
        Return CogT

        'AddHandler Text.DragEnter, AddressOf Text_DragEnter
        'AddHandler Text.DragDrop, AddressOf Text_Drop
    End Function
    Private Function BxDraw_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As AxMSINKAUTLib.AxInkPicture
        '@Draw|0|NEXT|Dock|DaT| 
        Dim imgconv As New ImageConverter()
        Dim TEGAKI As AxMSINKAUTLib.AxInkPicture
        TEGAKI = New AxMSINKAUTLib.AxInkPicture

        'TEGAKI.Ink.
        'myink = New MSINKAUTLib.InkDisp
        'Dim DrID As Long

        'OBJ.Controls.Add(TEGAKI)
        'TEGAKI.Dock = STK_PRT(3)

        ''AddHandler TEGAKI.DragEnter, AddressOf BxTEGAKI_DragEnter
        ''AddHandler Draw.DragDrop, AddressOf BxDraw_Drop
        'AddHandler TEGAKI.MouseLeaveEvent, AddressOf BxDraw_MouseLeave
        ''  TEGAKI.SizeMode = PictureBoxSizeMode.Zoom
        'TEGAKI.AllowDrop = True

        'Select Case Mid(STK_PRT(4), 1, 1)
        '    Case "∇" 'DB登録
        '        '記録場所を確保
        '        DrID = DBAcs_Insert(Lab, "Bdat", "BoxID,Num,Type", Lab.BoxID.ToString & "," & STK_PRT(1).ToString & ",'Draw'")
        '        'Blueprint修正 (DB側のみ)
        '        Box_Re_Draw(Lab, Lab.BoxID, Val(STK_PRT(1)), "∃" & DrID.ToString)
        '        TEGAKI.Tag = DrID.ToString
        '        Dim TEGAKIData() As Byte
        '        TEGAKIData = TEGAKI.Ink.Save()
        '        My.Computer.FileSystem.WriteAllBytes(gWORK_Path & "\DDX" & DrID.ToString, TEGAKIData, True)
        '    Case "∃"
        '        DrID = Val(Mid(STK_PRT(4), 2))
        '        TEGAKI.Tag = DrID.ToString
        '        Dim TEGAKIData() As Byte

        '        TEGAKIData = My.Computer.FileSystem.ReadAllBytes(gWORK_Path & "\DDX" & DrID.ToString)

        '        'TEGAKI.InkEnabled = False
        '        ' TEGAKI.Ink.Load(CType(imgconv.ConvertFrom(TEGAKIData), Image))
        '        TEGAKI.Ink.Load(TEGAKIData)
        '        'TEGAKI.InkEnabled = True
        '        'TEGAKI.InkEnabled = True
        '    Case Else

        'End Select
        'TEGAKI.Name = TEGAKI.ToString

        Return TEGAKI
    End Function



    Private Function BxTabP_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As TabPage
        '@TABp|0|TAB|Text|BackColor|"
        Dim Tabp As TabPage
        Tabp = New TabPage

        'AddHandler TBox.DragEnter, AddressOf BxText_DragEnter
        'AddHandler TBox.DragDrop, AddressOf BxText_Drop

        OBJ.Controls.Add(Tabp)
        With Tabp
            .Text = STK_PRT(3)
            .BackColor = Color.FromName(STK_PRT(4))
        End With
        Return Tabp

        'AddHandler Text.DragEnter, AddressOf Text_DragEnter
        'AddHandler Text.DragDrop, AddressOf Text_Drop
    End Function

    Private Function BxTaB_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As TabControl
        '@TAB|0|BASE|Dock|Alignment|ItemSizeX|ItemSizeY|"
        Dim Tab As TabControl
        Tab = New TabControl

        'AddHandler TBox.DragEnter, AddressOf BxText_DragEnter
        'AddHandler TBox.DragDrop, AddressOf BxText_Drop

        OBJ.Controls.Add(Tab)
        With Tab
            .Dock = STK_PRT(3)
            .Alignment = Val(STK_PRT(4))
            .ItemSize = New Size(Val(STK_PRT(5)), Val(STK_PRT(6)))
        End With
        Return Tab

        'AddHandler Text.DragEnter, AddressOf Text_DragEnter
        'AddHandler Text.DragDrop, AddressOf Text_Drop
    End Function



    Private Function BxSPLIT_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As SplitContainer
        ' '@Split|0|NEXT|Dock|Orientation|SplitterDistance
        Dim SPLIT As SplitContainer
        SPLIT = New SplitContainer
        OBJ.Controls.Add(SPLIT)
        With SPLIT
            .Dock = STK_PRT(3)
            .Orientation = Val(STK_PRT(4))
            .SplitterDistance = Val(STK_PRT(5))
            If STK_PRT.Count > 7 Then
                If STK_PRT(6) <> "" Then .Width = Val(STK_PRT(6))
                If STK_PRT(7) <> "" Then .Height = Val(STK_PRT(7))

            End If
        End With
        Return SPLIT

    End Function
    Private Function BxWeb_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As WebView2
        ' "@WebView|0|Wb.Hed|Wb.Dock|Wb.S|"
        Dim WebBox As WebView2
        WebBox = New WebView2

        'AddHandler TBox.DragEnter, AddressOf BxText_DragEnter
        'AddHandler TBox.DragDrop, AddressOf BxText_Drop

        OBJ.Controls.Add(WebBox)
        With WebBox
            .Dock = STK_PRT(3)
            If STK_PRT(4) <> "" Then
                .Source = New Uri(STK_PRT(4))
            End If
        End With
        Return WebBox

        'AddHandler Text.DragEnter, AddressOf Text_DragEnter
        'AddHandler Text.DragDrop, AddressOf Text_Drop
    End Function

    Private Function BxText_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As TextBox
        ' "@TEXT|0|T.Hed|T.Dock|T.Mul|T.Bc|T.Fc|T.FName|T.FSize|T.Text|"
        '"@TEXT|3|BASE|TBLR|0|16|200|80|-1|lemonchiffon|black|Arial|8|∇"
        Dim TxID As Integer
        Dim TBox As TextBox
        TBox = New TextBox
        Dim BluePrint As String
        AddHandler TBox.MouseLeave, AddressOf BxText_MouseLeave
        'AddHandler TBox.DragDrop, AddressOf BxText_Drop

        OBJ.Controls.Add(TBox)
        With TBox
            Select Case STK_PRT(3)
                Case "0", "1", "2", "3", "4", "5"
                    .Dock = Val(STK_PRT(3))
                Case Else
                    For L = 1 To Len(STK_PRT(3))
                        Select Case Mid(STK_PRT(3), L, 1)
                            Case "T"
                                .Anchor = .Anchor Or AnchorStyles.Top
                            Case "B"
                                .Anchor = .Anchor Or AnchorStyles.Bottom
                            Case "L"
                                .Anchor = .Anchor Or AnchorStyles.Left
                            Case "R"
                                .Anchor = .Anchor Or AnchorStyles.Right
                        End Select
                    Next

            End Select

            .Left = Val(STK_PRT(4))
            .Top = Val(STK_PRT(5))

            .Width = OBJ.width - .Left * 2
            .Height = OBJ.height - .Top * 2
            .AllowDrop = True
            .Multiline = Val(STK_PRT(8))
            .BackColor = Color.FromName(STK_PRT(9))
            .ForeColor = Color.FromName(STK_PRT(10))
            Dim f As New System.Drawing.Font(STK_PRT(11), STK_PRT(12))

            .Font = f
            .ScrollBars = ScrollBars.None
            .BorderStyle = BorderStyle.None

            .WordWrap = True

            '文字列記録場所を確保
            Dim DPath As String = Box_Get_DependPath(Lab.BoxList, Lab.BoxID, STK_PRT(1), "dat.txt")
            If System.IO.File.Exists(DPath) Then
                .Text = IO.File.ReadAllText(DPath)
            Else
                If STK_PRT(13) <> "" Then
                    System.IO.File.WriteAllText(DPath, STK_PRT(13))
                    .Text = STK_PRT(13)

                End If
            End If
            .Tag = DPath
        End With
        AddHandler DragDrop, AddressOf Text_Drop
        AddHandler KeyUp, AddressOf Text_KeyUp
        Return TBox

        'AddHandler Text.DragEnter, AddressOf Text_DragEnter

    End Function

    Private Function Box_Get_RowDat(BoxList As DataTable, DeskID As Integer, BoxID As Integer, Key As String) As String

        Dim Blist() As DataRow
        Dim dtRow As DataRow

        'Blist = BoxList.Select("BoxID = '" & BoxID.ToString & "'")
        'Blist = BoxList.Select("DeskID = '" & DeskID.ToString & "'")

        Blist = BoxList.Select("BoxID = '" & BoxID.ToString & "' AND DeskID = '" & DeskID.ToString & "'")
        If Blist.Count = 0 Then Return ""
        dtRow = Blist(0)

        Return dtRow(Key)

    End Function
    Private Function Box_Get_DependPath(ByRef BoxList As DataTable, ByRef BoxID As Integer, DependID As String, FileName As String) As String
        Dim Path As String
        Path = Box_Get_RowDat(BoxList, gLab.DeskID, BoxID, "Path")
        Path = System.IO.Path.GetDirectoryName(Path)
        Path &= "\" & DependID
        Dir_Check(Path)
        Path &= "\" & FileName
        Return Path
    End Function

    Private Function BxViewer_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As Object
        ' "@Viewer|2|BASE|★TYPE|★PATH|★NAME"



        Dim Type, Path, Name, Atr As String
        Type = STK_PRT(3)
        Path = STK_PRT(4)
        Name = STK_PRT(5)
        Atr = System.IO.Path.GetExtension(Path)

        Dim WebBox As WebView2
        WebBox = New WebView2
        OBJ.Controls.Add(WebBox)


        Dim flag0 As Integer
        flag0 = 0

        With WebBox
            '後でちゃんと直そうね
            .Anchor = .Anchor Or AnchorStyles.Top
            .Anchor = .Anchor Or AnchorStyles.Bottom
            .Anchor = .Anchor Or AnchorStyles.Left
            .Anchor = .Anchor Or AnchorStyles.Right
            .Left = 0
            .Top = 8

            .Width = OBJ.width - 4
            .Height = OBJ.height - .Top * 2


            Select Case Type
                Case "Url"
                    .Source = New Uri(Path)
                Case "File"
                    If InStr(".txt|.pdf|", Atr & "|") > 0 Then
                        .Source = New Uri("file:///" & Path)
                        flag0 = 1
                    End If
                    If InStr(".ico|.bmp|.jpg|.gif|.png|.exig|.tiff|", Atr & "|") > 0 Then
                        .Source = New Uri("file:///" & Path)
                        flag0 = 1
                    End If
                    If InStr(".mp3|.wav|.mid|.mp4|", Atr & "|") > 0 Then
                        .Source = New Uri("file:///" & Path)
                        flag0 = 1
                    End If

                    If InStr(".mp3|.wav|.mid|.mp4|", Atr & "|") > 0 Then
                        .Source = New Uri("file:///" & Path)
                        flag0 = 1
                    End If

                    If InStr(".ppt|.pptx|.pptm|.xls|.xlsx|.xlsm|.doc|.docx|.dotm|", Atr & "|") > 0 Then
                        .Source = New Uri("file:///" & gIcon_Path & "\settings.gif")
                        Dim OutF As String = gWORK_Path & "\" & gLab.BoxID & ".pdf"
                        Dim Cmd = Path & "," & OutF

                        '  Dim p As System.Diagnostics.Process = System.Diagnostics.Process.Start(gWORK_Path & "\PDFX.exe ", Path & "," & OutF)

                        flag0 = 1
                    End If

                    If InStr(".exe|", Atr & "|") > 0 Then
                        .Source = New Uri("file:///" & gIcon_Path & "\X.ico")
                        flag0 = 1
                    End If
                    'Dim answer As DialogResult
                    If flag0 = 0 Then
                        .Source = New Uri("file:///" & gIcon_Path & "\hatena.ico")
                        flag0 = 1
                    End If




            End Select

        End With


        Return WebBox

    End Function

    Private Function BxHeader_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As Object
        '@Header|2|BASE|★TYPE|★PATH|★NAME"
        Dim Type, Path, Name, Atr As String
        Type = STK_PRT(3)
        Path = STK_PRT(4)
        Name = STK_PRT(5)
        Atr = System.IO.Path.GetExtension(Path)


        Dim appIcon As Icon

        Dim pic As PictureBox
        pic = New PictureBox

        With pic
            .Dock = DockStyle.None
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left
            .Location = New Point(0, 0)
            .BackColor = Color.RoyalBlue
            .Height = 8
            .Width = 16
            .Padding = New Padding(0)
            .Margin = New Padding(0)
            .SizeMode = PictureBoxSizeMode.Zoom
            .WaitOnLoad = False

            .Tag = Path

            Select Case Type
                Case "File"
                    Try
                        appIcon = System.Drawing.Icon.ExtractAssociatedIcon(Path)
                        .Image = appIcon.ToBitmap()
                    Catch ex As Exception
                        .BackColor = Color.Red
                    End Try

                Case "Url"
                    .BackColor = Color.Red
                Case "Text"
                    .BackColor = Color.Blue
                Case "Pen"
                    .BackColor = Color.Green

            End Select
            AddHandler pic.DoubleClick, AddressOf BxPic_DoubleClick
            OBJ.Controls.Add(pic)



        End With

        Return pic

    End Function

    Private Function BxPict_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As PictureBox
        ' "@PICT|0|P.Hed|P.Dock|P.C|P.DAT|"
        Dim Pict As PictureBox
        Pict = New PictureBox
        Dim PxID As Long
        AddHandler Pict.DragEnter, AddressOf BxPict_DragEnter
        AddHandler Pict.DragDrop, AddressOf BxPict_Drop
        Pict.SizeMode = PictureBoxSizeMode.Zoom
        Pict.AllowDrop = True

        OBJ.Controls.Add(Pict)
        Pict.Dock = Val(STK_PRT(3))
        Pict.BackColor = Color.FromName(STK_PRT(4))

        Select Case Mid(STK_PRT(5), 1, 1)
            Case "∇" 'DB登録
                '記録場所を確保
                'パス記憶用  PxID = DBAcs_Insert(Lab, "PBox", "BoxID,Num", Lab.BoxID.ToString & "," & STK_PRT(1).ToString)
                'PxID = DBAcs_Insert(Lab, "BDat", "BoxID,Num", Lab.BoxID.ToString & "," & STK_PRT(1).ToString)
                'Blueprint修正 (DB側のみ)
                Box_Re_Pict(Lab, Lab.BoxID, Val(STK_PRT(1)), "∃" & PxID.ToString)

            Case "∃"
                Dim Ds As DataSet
                PxID = Val(Mid(STK_PRT(5), 2))
                'Bx_Get_Pict(Lab, Pict, PxID) 'ファイルパスから
                Bx_Get_Pict2(Lab, Pict, PxID) 'DBから

            Case Else

        End Select
        Pict.Tag = PxID.ToString

        Return Pict

    End Function

    Private Function BxBase_Set(ByRef Lab As Lab_Type, ByRef Obj As Object, ByRef STK_PRT() As String) As Panel
        'BASE生成
        '-------------------
        ' BluePrint &= "@BASE|0|B.Hed|B.Type|B.Dock|B.X|B.Y|B.Z|B.W|B.H|B.C|"

        Dim Base As Panel
        Base = New Panel

        AddHandler Base.MouseEnter, AddressOf Base_MouseEnter
        AddHandler Base.MouseLeave, AddressOf Base_MouseLeave
        AddHandler Base.MouseUp, AddressOf Base_MouseUp
        AddHandler Base.MouseDown, AddressOf Base_MouseDown
        AddHandler Base.MouseMove, AddressOf Base_MouseMove
        AddHandler Base.DoubleClick, AddressOf Base_DoubleClick
        AddHandler Base.DragDrop, AddressOf Base_DragDrop

        Obj.Controls.Add(Base)
        With Base
            .AllowDrop = True

            .Dock = Val(STK_PRT(3))
            .Left = Val(STK_PRT(4)) - Lab.WorldX
            .Top = Val(STK_PRT(5)) - Lab.WorldY
            .Width = Val(STK_PRT(7))
            .Height = Val(STK_PRT(8))
            ' .BackColor = Color.FromName("blue")
            .Padding = New Padding(5)
            .ContextMenuStrip = ContextMenuStrip3
        End With
        Base.Visible = True
        Return Base
    End Function

    Private Function BxLineH_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As Panel
        '@LineH|0|BASE|X|Y|Z|W|H|Col|" '設計図"@LineH|0|DESK|[BX]|[BY]|[BZ]|W|H|Col"
        Dim Line As Panel
        Line = New Panel
        If OBJ Is Nothing Then Return Nothing
        AddHandler Line.MouseEnter, AddressOf LineH_MouseEnter
        AddHandler Line.MouseLeave, AddressOf LineH_MouseLeave
        AddHandler Line.MouseUp, AddressOf LineH_MouseUp
        AddHandler Line.MouseDown, AddressOf Line_MouseDown
        AddHandler Line.MouseDoubleClick, AddressOf LineH_MouseDoubleClick
        AddHandler Line.MouseMove, AddressOf LineH_MouseMove
        AddHandler Line.DragOver, AddressOf Line_DragOver
        AddHandler Line.DragDrop, AddressOf Line_DragDrop
        '   AddHandler Line.DoubleClick, AddressOf Base_DoubleClick

        OBJ.Controls.Add(Line)
        With Line
            .Left = Val(STK_PRT(3)) - Lab.WorldX
            .Top = Val(STK_PRT(4)) - Lab.WorldY

            .Width = Val(STK_PRT(6))
            .Height = Val(STK_PRT(7))
            .BackColor = Color.FromName(STK_PRT(8))
            .Padding = New Padding(8)
            .AllowDrop = True
            .Tag = "LineH"
            .ContextMenuStrip = ContextMenuStrip3
        End With
        Return Line
        'AddHandler Text.DragEnter, AddressOf Text_DragEnter
        'AddHandler Text.DragDrop, AddressOf Text_Drop
    End Function


    Private Function BxLineV_Set(ByRef Lab As Lab_Type, ByRef OBJ As Object, ByRef STK_PRT() As String) As Panel
        '@LINE|0|BASE|X|Y|Z|W|H|Col|"
        Dim Line As Panel
        Line = New Panel
        If OBJ Is Nothing Then Return Nothing
        AddHandler Line.MouseEnter, AddressOf LineV_MouseEnter
        AddHandler Line.MouseLeave, AddressOf LineV_MouseLeave
        AddHandler Line.MouseUp, AddressOf LineV_MouseUp
        AddHandler Line.MouseDown, AddressOf Line_MouseDown
        AddHandler Line.MouseDoubleClick, AddressOf LineV_MouseDoubleClick
        AddHandler Line.MouseMove, AddressOf LineV_MouseMove
        AddHandler Line.DragOver, AddressOf Line_DragOver
        AddHandler Line.DragDrop, AddressOf Line_DragDrop
        '   AddHandler Line.DoubleClick, AddressOf Base_DoubleClick

        OBJ.Controls.Add(Line)
        With Line
            .Left = Val(STK_PRT(3)) - Lab.WorldX
            .Top = Val(STK_PRT(4)) - Lab.WorldY

            .Width = Val(STK_PRT(6))
            .Height = Val(STK_PRT(7))
            .BackColor = Color.FromName(STK_PRT(8))
            .Padding = New Padding(8)
            .AllowDrop = True
            .Tag = "LineV"
            ' .ContextMenuStrip = ContextMenuStrip3
        End With
        Return Line
        'AddHandler Text.DragEnter, AddressOf Text_DragEnter
        'AddHandler Text.DragDrop, AddressOf Text_Drop
    End Function
    Private Sub Box_Factory(ByRef Lab As Lab_Type, BoxID As Long, BluePrint As String)
        '設計図に従ってBOXの中身を作る
        '-----------------------------
        Dim STACKS() As String
        Dim STK_PRT() As String
        Dim ObjBase As Object
        Dim ObjLine As Object
        Dim ObjTab As Object
        Dim ObjParts As Object
        Dim ObjSpl As SplitContainer
        Dim ObjSpl1 As Object
        Dim ObjSpl2 As Object
        Dim Obj As Object

        STACKS = Split(BluePrint, "@")
        For L = 1 To STACKS.Count - 1
            STK_PRT = Split(STACKS(L), "|")

            Select Case STK_PRT(2)
                Case "DESK"
                    Obj = Lab.DesKTop(Lab.DeskTopID)
                Case "LINE"
                    Obj = ObjLine
                    'Lab.Base = Obj
                Case "BASE"
                    Obj = ObjBase
                   ' Lab.Base = Obj

                Case "TAB"
                    Obj = ObjTab
                Case "SPLIT1"
                    Obj = ObjSpl1
                Case "SPLIT2"
                    Obj = ObjSpl2
                Case "NEXT"
                    Obj = ObjParts
                Case Else
                    Obj = ObjParts
            End Select

            Select Case STK_PRT(0)
                Case "BASE"
                    ObjBase = BxBase_Set(Lab, Obj, STK_PRT)
                    ObjBase.name = BoxID.ToString
                    Lab.Base = ObjBase

                Case "LineH"
                    ObjLine = BxLineH_Set(Lab, Obj, STK_PRT)
                    If ObjLine Is Nothing Then Return
                    ObjLine.name = BoxID.ToString
                    Lab.Base = ObjLine
                    ObjLine.tag = "LineH"
                Case "LineV"
                    ObjLine = BxLineV_Set(Lab, Obj, STK_PRT)
                    If ObjLine Is Nothing Then Return
                    ObjLine.name = BoxID.ToString
                    Lab.Base = ObjLine
                    ObjLine.tag = "LineV"

                Case "Viewer"
                    ObjParts = BxViewer_Set(Lab, Obj, STK_PRT)

                Case "Header"
                    ObjParts = BxHeader_Set(Lab, Obj, STK_PRT)

                Case "TAB"
                    ObjTab = BxTaB_Set(Lab, Obj, STK_PRT)
                    ObjTab.name = BoxID.ToString & "." & STK_PRT(1).ToString

                Case "SPLIT"
                    ObjSpl = BxSPLIT_Set(Lab, Obj, STK_PRT)
                    ObjSpl1 = ObjSpl.Panel1
                    ObjSpl2 = ObjSpl.Panel2

                Case "PICT"
                    ObjParts = BxPict_Set(Lab, Obj, STK_PRT)

                Case "TEXT"
                    ObjParts = BxText_Set(Lab, Obj, STK_PRT)

                Case "WebView"
                    ObjParts = BxWeb_Set(Lab, Obj, STK_PRT)

                Case "TABp"
                    ObjParts = BxTabP_Set(Lab, Obj, STK_PRT)
                    ObjParts.name = BoxID.ToString & "." & STK_PRT(1).ToString
                Case "Draw"
                    ObjParts = BxDraw_Set(Lab, Obj, STK_PRT)

                Case "CogT"
                    ObjParts = BxCogT_Set(Lab, Obj, STK_PRT)


            End Select

        Next

    End Sub

    Private Sub Box_Makes_OnDESK(ByRef Lab As Lab_Type, BluePrint As String)
        Dim Blist() As DataRow
        Dim dtRow As DataRow
        With Lab
            If .BoxList Is Nothing Then
                Dim Ent() As String
                Ent = Split("BoxID,DeskID,ProjKey,OwnerID,State,Type,ParentID,LineID,LineNo,BluePrint,Path", ",")
                .BoxList = New DataTable("Box")

                For L = 0 To Ent.Count - 1
                    If Ent(L) <> "" Then .BoxList.Columns.Add(Ent(L))
                Next
                .BoxID = 1
            End If

            Blist = Lab.BoxList.Select("DeskID ='" & .DeskID.ToString & "'")
            If Blist.Count = 0 Then
                .BoxID = 1
            Else
                .BoxID = Blist.Count + 1
            End If
            With .BoxList
                dtRow = .NewRow
                dtRow("BoxID") = Lab.BoxID.ToString
                dtRow("DeskID") = Lab.DeskID.ToString
                dtRow("ProjKey") = Lab.ProjOwnerID.ToString & "\" & Lab.ProjID.ToString
                dtRow("OwnerID") = gX_ownerID.ToString
                dtRow("State") = "Nomal"
                dtRow("Type") = ""
                dtRow("LineID") = ""
                If InStr(BluePrint, "LineH") > 0 Then dtRow("LineID") = "LineH"
                If InStr(BluePrint, "LineV") > 0 Then dtRow("LineID") = "LineV"

                dtRow("LineNo") = ""

                dtRow("BluePrint") = BluePrint
                'Pa(Pa.Count - 4).ToString & "\" & Pa(Pa.Count - 3).ToString
                dtRow("Path") = gENV_Path & "\PRJ\" & Lab.ProjOwnerID.ToString & "\" & Lab.ProjID.ToString & "\" & Lab.DeskID.ToString & "\" & Lab.BoxID.ToString & "\box.txt"

                .Rows.Add(dtRow)
                Box_Save_Info(dtRow)  'ファイルに保存
            End With

            Box_Factory(Lab, .BoxID, BluePrint) '製造と配置と
        End With


        ' Lab.Base.BringToFront()
    End Sub
    'Private Function Box_SaveDB(ByRef Lab As Lab_Type, BluePrint As String) As Long
    '    Dim BoxID As Long '製造番号
    '    Dim Data As String
    '    With Lab
    '        Data = ""
    '        Data &= .DeskID.ToString & ","
    '        Data &= .ProjID.ToString & ","
    '        Data &= .OwnerID.ToString & ","
    '        Data &= "'" & BluePrint & "',"
    '        Data &= "1"
    '    End With

    '    BoxID = DBAcs_Insert(gLab, "Box", "DeskID,ProjKey,OwnerID,BluePrint,State", Data)


    '    Return BoxID
    'End Function

    Private Function Box_Get_BluePrint(ByRef Lab As Lab_Type, Type As String) As String
        '   設計図生成
        '-----------------------------
        Dim BluePrint As String

        'パラメータ一覧
        'Dock 1上　２下　1左　４→
        '@BASE|0|DESK|Dock|X|Y|Z|W|H|"
        '@PICT|0|BASE|Dock|C|DAT|"
        '@TEXT|0|BASE|Dock|Mul|Bc|Fc|FName|FSize|Text|"
        '@WebView|0|BASE|Dock|URL|"
        '@TAB|0|BASE|Dock|Alignment|ItemSizeX|ItemSizeY|"
        '@TABp|0|TAB|Text|BackColor|"
        '@CogT|0|NEXT|Dock|TEXT|"
        '@Draw|0|NEXT|Dock|DAT|
        '@Split|0|NEXT|Dock|Orientation|SplitterDistance|

        With Lab
            Select Case Type
                Case "Fig"
                    '設計図
                    BluePrint = "Fig"
                    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|100|100"
                    BluePrint &= "@PICT|2|BASE|5|cadetblue|"
                    BluePrint &= "@TEXT|3|BASE|2|-1|black|white|Arial|8|Fig."

                'Case "File"
                '    '設計図
                '    BluePrint = "FILE"
                '    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|100|20"

                '    BluePrint &= "@PICT|2|BASE|1|darkblue|★"
                '    BluePrint &= "@TEXT|3|BASE|5|0|black|white|Arial|16|"

                Case "Viewer"
                    '設計図
                    BluePrint = "Viewer"
                    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|200|50"
                    BluePrint &= "@Header|2|BASE|★TYPE|★PATH|★NAME"
                    BluePrint &= "@Viewer|3|BASE|★TYPE|★PATH|★NAME"

                Case "Post_it"
                    '設計図
                    BluePrint = "Post_it"
                    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|200|100"
                    BluePrint &= "@Header|2|BASE|Text|★PATH|Post_it"
                    BluePrint &= "@TEXT|3|BASE|TBLR|0|8|200|80|-1|black|white|Arial|12|"

                Case "LineH"
                    '設計図"@LineH|0|DESK|[BX]|[BY]|[BZ]|W|H|Col"
                    BluePrint = "LineH"
                    BluePrint &= "@LineH|1|DESK|[BX]|[BY]|Z|600|4|black"

                Case "LineV"
                    '設計図""@LineV|1|DESK|[BX]|[BY]|Z|2|200|blue"
                    BluePrint = "LineV"
                    BluePrint &= "@LineV|1|DESK|[BX]|[BY]|Z|4|300|black"
                Case "Pen_Memo"
                    '設計図
                    BluePrint = "Pen_Memo"
                    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|200|200"
                    BluePrint &= "@Header|2|BASE|Text|★PATH|Pen_Memo"
                    BluePrint &= "@Draw|3|NEXT|5|"


                Case "Web_View"
                    '設計図
                    BluePrint = "Web_View"
                    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|400|250"
                    BluePrint &= "@WebView|2|BASE|5|https://www.google.com/"



                Case "G-Docs"
                    '設計図
                    BluePrint = "Web_View"
                    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|400|250"
                    BluePrint &= "@WebView|2|BASE|5|https://docs.google.com/document/d/1MFxmfOIaZKf7KFYoyE7vDC1-77IASuqEb82M9AnP11U/edit#|"



                Case "Web_J-STAGE"
                    '設計図
                    BluePrint = "Web_View"
                    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|400|250"
                    BluePrint &= "@WebView|2|BASE|5|https://www.jstage.jst.go.jp/browse/-char/ja"


                Case "ID_Card"
                    BluePrint = "ID_Card"
                    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|200|200|"
                    BluePrint &= "@SPLIT|2|BASE|5|0|80|"
                    BluePrint &= "@TEXT|3|SPLIT2|5|-1|white|black|Arial|8|introduce"
                    BluePrint &= "@SPLIT|4|SPLIT1|5|1|60|"
                    BluePrint &= "@PICT|5|SPLIT1|5|white|"
                    BluePrint &= "@TEXT|6|SPLIT2|1|0|white|black|Arial|8|Tel"
                    BluePrint &= "@TEXT|7|SPLIT2|1|0|white|black|Arial|8|Mail"
                    BluePrint &= "@TEXT|8|SPLIT2|1|0|white|black|Arial|8|Name"
                    BluePrint &= "@TEXT|9|SPLIT2|1|0|red|white|Arial|8|Title"

                Case "ID_Card2"
                    BluePrint = "ID_Card"
                    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|200|200|"
                    BluePrint &= "@SPLIT|2|BASE|5|0|80|"
                    BluePrint &= "@TEXT|3|SPLIT2|5|-1|white|black|Arial|8|introduce"
                    BluePrint &= "@SPLIT|4|SPLIT1|5|1|60|"
                    BluePrint &= "@PICT|5|SPLIT1|5|white|"
                    BluePrint &= "@TEXT|6|SPLIT2|1|0|white|black|Arial|8|Tel"
                    BluePrint &= "@TEXT|7|SPLIT2|1|0|white|black|Arial|8|Mail"
                    BluePrint &= "@TEXT|8|SPLIT2|1|0|white|black|Arial|8|Name"
                    BluePrint &= "@TEXT|9|SPLIT2|1|0|red|white|Arial|8|Director of Laboratory"



                Case "Summary"
                    BluePrint = "Summary"
                    BluePrint &= "@BASE|1|DESK|0|[BX]|[BY]|[BZ]|650|100|"
                    BluePrint &= "@SPLIT|2|BASE|5|1|50|"

                    BluePrint &= "@TEXT|3|SPLIT1|1|0|white|black|Arial|8|Worries"
                    BluePrint &= "@TEXT|4|SPLIT1|1|0|black|white|Arial|8|Hypothesis"
                    BluePrint &= "@TEXT|5|SPLIT1|1|0|black|white|Arial|8|Summary"
                    BluePrint &= "@TEXT|6|SPLIT1|1|0|black|white|Arial|8|TITLE"

                    BluePrint &= "@TEXT|7|SPLIT2|1|0|white|black|Arial|8|"
                    BluePrint &= "@TEXT|8|SPLIT2|1|0|white|black|Arial|8|"
                    BluePrint &= "@TEXT|9|SPLIT2|1|0|white|black|Arial|8|"
                    BluePrint &= "@TEXT|10|SPLIT2|1|0|white|black|Arial|8|"


            End Select

        End With
        Return BluePrint
    End Function





    Private Sub Dir_Check(ByVal Path As String)
        If System.IO.Directory.Exists(Path) Then Return
        System.IO.Directory.CreateDirectory(Path)
    End Sub
    Private Sub Lab_Set_Env(ByRef Lab As Lab_Type)
        '環境変数セット
        '------------------------------------------
        Log_Show(1, "環境変数セット")


        'グローバル変数セット
        '--------------------
        gApp_Path = IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)    'EXE位置
        gWORK_Path = gApp_Path & "\WORK" : Dir_Check(gWORK_Path)   '各種WORKデータへのパス  
        gENV_Path = gApp_Path & "\ENV" : Dir_Check(gENV_Path)    '各種設定データへのパス  
        gLOG_Path = gApp_Path & "\LOG" : Dir_Check(gLOG_Path)  'ログ出力先のパス  
        gIcon_Path = gApp_Path & "\ENV\ICONS" : Dir_Check(gIcon_Path)  'ICON格納フォルダ

        With gLab

            .WorldX = -2000 'MAPデスクの初期X位置
            .WorldY = -2000  'MAPデスクの初期Y位置
            .WorldW = 24000  'MAPデスクの幅
            .WorldH = 24000  'MAPデスクの高さ
        End With

        'イメージ設定 (アイコンフォルダの画像イメージを利用にかかわらずすべて登録
        '--------------------
        ' Log_Show_Splash("イメージ設定")
        gLab.Image_List = Image_Init(gIcon_Path)

        'デコレーション
        '-----------------------------------
        'gDeco.BackColor = "DarkRed"
        ' gDeco.ForeColor = "Black"

        gDeco.BackColor = "Black"
        gDeco.ForeColor = "White"

        Set_WorkPanel()

        'コントロール設定
        '--------------------

        Controls_Init()

        'AT登録

        With gLab.ProjAT.View
            .ListView = ListView1
            Columns_Set_Env(.Columns, "[Proj]")
        End With

        With gLab.ElementAT.View
            .ListView = ListView2
            Columns_Set_Env(.Columns, "[Element]")
        End With

        With gLab.StackAT.View
            .ListView = ListView3
            Columns_Set_Env(.Columns, "[Stack]")
        End With
        With SplitContainer1
            .FixedPanel = FixedPanel.Panel2
            .IsSplitterFixed = True
        End With


        'With gLab.DeskAT
        '    .View = TabControl3


        '    'ReDim .Desk(2) ' 0-データパーツ用 ,1--Proj用
        '    'ReDim .Person(0)
        'End With
        ' RepoList_Init(gLab.View) '一覧設定
        gLab.DeskViewMax = 16

        With DataGridView3

            .ColumnHeadersVisible = False
            .RowHeadersVisible = False
            .ColumnCount = gLab.DeskViewMax
            .RowCount = 1
            .CellBorderStyle = DataGridViewCellBorderStyle.None
            For L = 0 To 15
                .Columns(L).Width = 80 '.Width / 10
                .Rows(0).Cells(L).Value = ""
                .Rows(0).Cells(L).Style.BackColor = Color.FromName("dimgray")
                '  AddHandler .Rows(0).Cell, AddressOf Desk_MouseDown
                .Rows(0).Cells(L).ReadOnly = True
            Next

        End With




        'コンテキストメニューの設定
        '-----------------------------------
        Menu_Set(ContextMenuStrip1, "プロジェクトの作成/プロジェクト名変更/プロジェクトの削除", "NEW/RENAME/DEL")
        Menu_Set(ContextMenuStrip2, "デスクの追加/デスク名変更/デスクの削除", "NEW/RENAME/DEL")
        Menu_Set(ContextMenuStrip3, "スタックのコピー/スタックの貼付け/スタックの削除", "COPY/PASTE/DEL")
        Menu_Set(ContextMenuStrip4, "資料の削除", "DEL")



        gLab.ProjAT.View.ListView.ContextMenuStrip = ContextMenuStrip1
        DataGridView3.ContextMenuStrip = ContextMenuStrip2
        ListView2.ContextMenuStrip = ContextMenuStrip4


    End Sub

    Private Sub Box_Set_OnDESK(ByRef lab As Lab_Type, Desk As Panel)
        'BOXをDESKに再配置（データを再度SAVEしないように注意）
        '-----------------------------
        Dim Blist As DataRow()


        With lab
            Blist = .BoxList.Select("State = 'Nomal' AND DeskID = '" & .DeskID.ToString & "'")
            If Blist.Count = 0 Then Return

            For Each row As DataRow In Blist
                .BoxID = row("BoxID")

                Box_Factory(gLab, .BoxID, row("BluePrint")) '再配置

                If row("LineID") <> "" Then
                    .Base.Tag = row("LineID")
                    '  .Base.Name = .BoxID.ToString
                    ' MsgBox("●" & .DeskID.ToString & "/" & .Base.Name & "/" & .BoxID.ToString)
                Else
                    If .Base.Tag = "" Then .Base.Tag = "OnDESK"
                    '   .Base.Name = .BoxID.ToString
                    'MsgBox(.DeskID.ToString & "/" & .Base.Name & "/" & .BoxID.ToString)
                End If
            Next

        End With

    End Sub

    Private Sub Box_Show_List()
        'スタックをLISTVIEWに表示
        '-----------------------------

        With gLab.StackAT.View
            Columns_Set_Listview(.ListView, .Columns)
        End With

        Box_Set_Item(gLab, "Post_it", "Post_it", "Post_it")
        Box_Set_Item(gLab, "Pen_Memo", "Pen_Memo", "Pen_Memo")
        Box_Set_Item(gLab, "LineH", "LineH", "LineH")
        Box_Set_Item(gLab, "LineV", "LineV", "LineV")

        'Box_Set_Item(gLab, "Fig", "Fig", "Fig")
        'Box_Set_Item(gLab, "Webブラウザ", "Web_View", "web")
        'Box_Set_Item(gLab, "スタンプ（公開）", "Stamp(OPEN)", "Stamp_OPEN")
        'Box_Set_Item(gLab, "スタンプ（非公開）", "Stamp(CLOSE)", "Stamp_CLOSE")
        ''  Box_Set_Item(gLab, "スタンプ（welcome）", "Stamp(welcome)", "Stamp_Welcome")
        'Box_Set_Item(gLab, "ID_Card", "ID_Card", "ID_Card")
        'Box_Set_Item(gLab, "サマリー", "Summary", "Summary")
        'Box_Set_Item(gLab, "Webブラウザ(Googleドキュメント)", "G-Docs", "G-Docs")
        'Box_Set_Item(gLab, "Webブラウザ(J-Stage)", "Web_J-STAGE", "J-STAGE_logo")
        'Box_Set_Item(gLab, "クリッパー(製作中)", "ハサミ", "ハサミ")
        'Box_Set_Item(gLab, "引用マーカ(製作中)", "マーカー", "マーカー")
        'Box_Set_Item(gLab, "スパイダー(製作中)", "SPIDER_J-STAGE", "スパイダー")
        'Box_Set_Item(gLab, "フォローJ(製作中)", "Web_J-STAGE", "ターゲット")


    End Sub

    Private Sub Box_Set_Item(ByRef Lab As Lab_Type, Name As String, Key As String, IconKey As String)
        'ElementデータをLISTVIEWに表示
        '-----------------------------
        Dim Itm As ListViewItem

        With Lab.StackAT.View.ListView

            Itm = .Items.Add(Name)
            Itm.ImageIndex = Image_Get_ID(Lab.Image_List, IconKey)
            ' Itm.ImageIndex = Image_Get_ID(Lab.Image_List, IconKey)
            Itm.Tag = Key

        End With

    End Sub


    Private Sub Controls_Init()

        'バックから―
        With gDeco
            '
            'SplitContainer1.Panel1.BackColor = Color.FromName(.BackColor)
            'SplitContainer1.Panel2.BackColor = Color.FromName(.BackColor)

            SplitContainer5.Panel1.BackColor = Color.FromName(.BackColor)

            ListView1.BackColor = Color.FromName(.BackColor)
            ListView2.BackColor = Color.FromName(.BackColor)
            ListView3.BackColor = Color.FromName(.BackColor)

            'SplitContainer1.Panel1.ForeColor = Color.FromName(.ForeColor)
            'SplitContainer1.Panel2.ForeColor = Color.FromName(.ForeColor)



            SplitContainer5.Panel1.ForeColor = Color.FromName(.ForeColor)


            ListView1.ForeColor = Color.FromName(.ForeColor)
            ListView2.ForeColor = Color.FromName(.ForeColor)
            ListView3.ForeColor = Color.FromName(.ForeColor)


            TextBox1.ForeColor = Color.FromName(.ForeColor)
            'TextBox2.ForeColor = Color.FromName(.ForeColor)
            'TextBox3.ForeColor = Color.FromName(.ForeColor)
            'TextBox4.ForeColor = Color.FromName(.ForeColor)
            'TextBox5.ForeColor = Color.FromName(.ForeColor)
            TextBox6.ForeColor = Color.FromName(.ForeColor)

            TextBox1.BackColor = Color.FromName(.BackColor)
            'TextBox2.BackColor = Color.FromName(.BackColor)
            'TextBox3.BackColor = Color.FromName(.BackColor)
            'TextBox4.BackColor = Color.FromName(.BackColor)
            'TextBox5.BackColor = Color.FromName(.BackColor)
            TextBox6.BackColor = Color.FromName(.BackColor)


        End With




        'ツールチップ
        With ToolTip1
            .ToolTipTitle = ""
            '.ToolTipIcon = "" ' ToolTipIcon.Info
            '.SetToolTip(PictureBox2, "研究画面を開きます。")
            ''.SetToolTip(PictureBox3, "研究を探します。")
            '.SetToolTip(PictureBox3, "新しい研究を始めます。")
            '.SetToolTip(PictureBox4, "私について")

            '.SetToolTip(PictureBox3, "研究ノートを開きます。")

        End With

        'LISTVIEW

        With ListView1
            .View = View.Details
            '.LargeImageList = gImg_List.Image
            '.SmallImageList = gImg_List.Image
            '.StateImageList = gImg_List.Image

        End With
        With ListView2 'スタック用
            '.View = View.LargeIcon
            .View = View.Details
            '.LabelWrap = True
            ' .LargeImageList = gImg_List.Image
            .SmallImageList = gImg_List.Image
            ' .StateImageList = gImg_List.Image
        End With
        With ListView3 '資料用
            '   .View = View.LargeIcon
            .View = View.Details
            '  .LabelWrap = True
            '.LargeImageList = gImg_List.Image
            .SmallImageList = gImg_List.Image
            '.StateImageList = gImg_List.Image

        End With





    End Sub

    Function Owner_Get_ID(ByRef lab As Lab_Type) As Integer
        'オーナー番号取得
        Dim tmpID As Integer
        Dim Owner_Path As String = gENV_Path & "\ginger.id"


        If System.IO.File.Exists(Owner_Path) = False Then
            Return 0
        Else
            With lab
                gX_ownerID = Owner_Get_OwnerID(lab.OwnerList, Owner_Path)
            End With

        End If

        Return gX_ownerID
    End Function

    Private Function Owner_Get_OwnerID(ByRef OwnerList As DataTable, ByRef Owner_Path As String) As Integer
        If OwnerList Is Nothing Then
            OwnerList = New DataTable
        End If

        Dim Ent() As String
        Ent = Split("UserID,Name,Alias,Token", ",")
        For L = 0 To Ent.Count - 1
            If Ent(L) <> "" Then OwnerList.Columns.Add(Ent(L))
        Next

        Dim TmpD As String
        TmpD = IO.File.ReadAllText(Owner_Path)
        Dim TmpR() As String = Split(TmpD, vbCrLf)
        Dim key() As String
        Dim dtRow As DataRow
        With OwnerList
            dtRow = .NewRow

            For L = 0 To TmpR.Count - 1
                key = Split(TmpR(L), vbTab)
                If Trim(key(0)) <> "" Then
                    dtRow(key(0)) = key(1)
                End If

            Next
            .Rows.Add(dtRow)
            gX_ownerID = dtRow("UserID")
            gX_owner = dtRow("Name")
            gX_Token = dtRow("Token")
            Return dtRow("UserID")
        End With


    End Function

    Private Sub Owner_Add_OwnerListByFile(ByRef DeskList As DataTable, File As String)

        Dim Pa() As String
        Pa = Split(File, "\")

        Dim TmpD As String
        TmpD = IO.File.ReadAllText(File)

        Dim TmpR() As String = Split(TmpD, vbCrLf)
        Dim key() As String
        Dim dtRow As DataRow
        With DeskList
            dtRow = .NewRow

            For L = 0 To TmpR.Count - 1
                key = Split(TmpR(L), vbTab)
                If Trim(key(0)) <> "" Then
                    dtRow(key(0)) = key(1)
                End If

            Next
            dtRow("Path") = File
            dtRow("ProjKey") = Pa(Pa.Count - 4).ToString & "\" & Pa(Pa.Count - 3).ToString
            dtRow("DeskID") = Pa(Pa.Count - 2).ToString

            .Rows.Add(dtRow)
        End With
    End Sub
    Sub Owner_Save_ID(ByVal ID As Integer)
        'OWNER ID をファイルに保存
        ' Dim tmpDir As String = System.IO.Path.GetTempPath()
        Dim tmpDir As String = System.Windows.Forms.Application.StartupPath() & "\ENV"
        Dim sw As New System.IO.StreamWriter(tmpDir & "\ginger.id")

        sw.WriteLine(ID.ToString)
        sw.Close()


    End Sub

    Sub Owner_Show_Data(Lab As Lab_Type)
        'OWNER Data を表示
        '-------------------

        'With Lab.Person(0)
        '    TextBox4.Text = .Name
        '    TextBox2.Text = .Tel
        '    TextBox3.Text = .Mail
        '    TextBox5.Text = .Prof
        '    If System.IO.File.Exists(.Image) = True Then PictureBox1.ImageLocation = .Image
        'End With

    End Sub

    Sub Ping(tmpS As String)
        'Pingオブジェクトの作成
        Dim p As New System.Net.NetworkInformation.Ping()
        '"www.yahoo.com"にPingを送信する
        Dim reply As System.Net.NetworkInformation.PingReply = p.Send(tmpS)

        '結果を取得
        If reply.Status = System.Net.NetworkInformation.IPStatus.Success Then
            'Console.WriteLine("Reply from {0}:bytes={1} time={2}ms TTL={3}",
            '    reply.Address, reply.Buffer.Length,
            '    reply.RoundtripTime, reply.Options.Ttl)
            Log_Show(2, "Ping送信OK" & tmpS & reply.Status)
        Else
            Log_Show(2, "Ping送信に失敗。" & tmpS & reply.Status)
        End If

        p.Dispose()
    End Sub


    Private Sub ToolStripButton_Set_Env(ByRef btn As ToolStripButton, ByRef key As String)

        With btn
            '.ImageKey = Image_Get_ID(key, gImg_List)
            .ImageKey = key
            .ToolTipText = key
        End With

    End Sub
    Private Sub Button_Set_Env(ByRef btn As Button, ByRef key As String)
        'ボタンに絵を出す

        With btn
            .ImageList = gImg_List.Image
            .ImageKey = key
            '.ToolTipText = key
        End With

    End Sub

    Private Sub Proj_Make_New(ByRef Lab As Lab_Type)
        'Projの追加 (DB修正)
        '-------------------------------------------
        Dim ProjID As Integer
        With Lab
            'TBL追加
            If .ProjList Is Nothing Then
                .ProjList = New DataTable("Project")
                Dim Ent() As String
                Ent = Split("ProjID|Title|OwnerID|Path|State|Summary|Hypothesis|Purpose|Problem", "|")
                For L = 0 To Ent.Count - 1
                    If Ent(L) <> "" Then .ProjList.Columns.Add(Ent(L))
                Next
                ProjID = 1
            Else
                Dim Maxrow() As DataRow = .ProjList.Select("OwnerID = '" & gX_ownerID.ToString & "'")
                If Maxrow.Count = 0 Then
                    ProjID = 1
                Else
                    ProjID = Maxrow.Count + 1
                End If

            End If

            Dim dtRow As DataRow
            With .ProjList
                dtRow = .NewRow
                dtRow("ProjID") = ProjID
                dtRow("ProjKey") = gX_ownerID.ToString & "\" & ProjID.ToString
                dtRow("Title") = "新しいプロジェクト"
                dtRow("OwnerID") = gX_ownerID.ToString
                dtRow("Path") = gENV_Path & "\PRJ\" & gX_ownerID.ToString & "\" & ProjID.ToString & "\" & "Proj.txt"
                dtRow("State") = "Nomal"
                dtRow("Summary") = ""
                dtRow("Hypothesis") = ""
                dtRow("Purpose") = ""
                dtRow("Problem") = ""
            End With

            Proj_Save_Info(dtRow) 'ファイル書き込み
            Proj_Add_List(dtRow)  'LISTVIEW追加

            Desk_Clear_Name(Lab)
            Desk_Clear_Desk(Lab)

            Desk_Make_New(Lab) 'デスク追加
        End With

    End Sub



    'Private Function ProjDB_Insert(ByRef Lab As Lab_Type, Name As String, ByRef Key As String, Val As String) As Long
    '    'DB Node　を追加
    '    '--------------------------
    '    Dim tmpR As String

    '    Dim DBCon As SqlConnection
    '    Dim DBCom As SqlCommand
    '    Dim TmpCmd As String


    '    With Lab
    '        Select Case .Db_Type
    '            Case "SQLite"
    '            Case "MSSQL"
    '                DBCon = New SqlConnection(Lab.DbPath & ";Initial Catalog =" & .Db_Type & Name)
    '                TmpCmd = "INSERT INTO " & "TBL_" & Name
    '                TmpCmd &= Key & "VALUES" & Val
    '                TmpCmd &= "; SELECT SCOPE_IDENTITY();"
    '        End Select

    '    End With

    '    DBCom = New SqlCommand(TmpCmd, DBCon)
    '    DBCom.CommandTimeout = 0

    '    DBCom.Connection.Open()

    '    tmpR = -1
    '    For I As Integer = 1 To 50
    '        Try
    '            tmpR = DBCom.ExecuteScalar()

    '            Exit For
    '        Catch ex As Exception
    '            Log_Show(1, "DB登録エラー" & I & ex.Message)
    '            For J As Integer = 1 To 100
    '                System.Windows.Forms.Application.DoEvents()
    '            Next
    '        End Try
    '    Next

    '    DBCom.Connection.Close()
    '    DBCon.Close()
    '    DBCon.Dispose()
    '    Return tmpR

    'End Function
    'Private Function ProjDB_Insert_OLD(ByRef Lab As Lab_Type, Name As String, ByRef Proj As Proj_Type) As Integer
    '    'DB Node　を追加
    '    '--------------------------
    '    Dim tmpR As String

    '    Dim DBCon As SqlConnection
    '    Dim DBCom As SqlCommand
    '    Dim TmpCmd As String

    '    DBCon = New SqlConnection(Lab.DbPath & ";Initial Catalog =" & Lab.DBname & Name)
    '    With Proj

    '        TmpCmd = "INSERT INTO " & "TBL_" & Name
    '        TmpCmd &= " ( Name, Prof, Note, OwnerID, Owner, PID) VALUES ("

    '        TmpCmd &= "'" & .Name & "',"

    '        TmpCmd &= "'" & .OwnerID & "',"
    '        TmpCmd &= "'" & .Owner & "',"

    '        TmpCmd &= ")"
    '        TmpCmd &= "; SELECT SCOPE_IDENTITY();"
    '    End With

    '    DBCom = New SqlCommand(TmpCmd, DBCon)
    '    DBCom.CommandTimeout = 0

    '    DBCom.Connection.Open()

    '    tmpR = -1
    '    For I As Integer = 1 To 50
    '        Try
    '            tmpR = DBCom.ExecuteScalar()

    '            Exit For
    '        Catch ex As Exception
    '            Log_Show(1, "DB登録エラー" & I & ex.Message)
    '            For J As Integer = 1 To 100
    '                System.Windows.Forms.Application.DoEvents()
    '            Next
    '        End Try
    '    Next

    '    DBCom.Connection.Close()
    '    DBCon.Close()
    '    DBCon.Dispose()
    '    Return tmpR

    'End Function



    Private Sub Menu_Act_Dicbox(ByRef Proj As Lab_Type, ByRef Cmd() As String)
        'メニュー選択時のアクションまとめ TREE
        '-----------------------------------------
        Dim DeskID As Integer
        DeskID = 0
        Dim Act_Effect As String
        Dim N_Name, O_Name As String
        Dim O_NID As Integer
        Dim TmpV As String

        Dim D_Node() As TreeNode
        D_Node = Nothing
        Dim N_Type, N_State As String
        N_Type = "" : N_State = ""


        N_Name = ""

        'Dim History As History_Type
        'History = Proj.Desk(DeskID).History



    End Sub


    Private Sub Menu_Act_DicList(ByRef Proj As Lab_Type, ByRef LV As ListView, ByRef Cmd() As String)
        'メニュー選択時のアクションまとめ DIC
        '-----------------------------------------
        'Dim DeskID As Integer
        'DeskID = 0

        'Dim Flag_RID, Flag_DID As Integer
        'Dim Item() As ListViewItem
        'Dim Cb_data As String

        'With LV
        '    '対象ノード非選択
        '    If .SelectedItems.Count = 0 Then Return
        '    '対象ノード
        '    ReDim Item(.SelectedItems.Count - 1)
        '    For L = 0 To .SelectedItems.Count - 1
        '        Item(L) = .SelectedItems(L)
        '        Progres_Show(L + 1, .SelectedItems.Count, False)
        '    Next

        '    Select Case Cmd(1)

        '        Case "コピー"
        '            gBff_Act_Type = "COPY_ITEM"
        '            gBff_Node = Nothing
        '            gBff_Item = Item.Clone


        '            'クリップボードにコピー
        '            Clipboard.Clear()   'クリップボードを初期化
        '            Clipboard.SetText(ListView_Get_TextData(LV))



        '        Case "切り取り"
        '            gBff_Act_Type = "MOVE_ITEM"
        '            gBff_Node = Nothing
        '            gBff_Item = Item.Clone

        '        Case "貼り付け"
        '            '■何もしない
        '            Log_Show(2, "指定のオブジェクトは貼り付けできません")

        '        Case "削除"
        '            '''Histroy_Start(gLab.Desk(DeskID).History)
        '            'For L = Item.Length - 1 To 0 Step -1
        '            '    Proj_Act_Dat_Del_Item(Proj, Item(L), True)
        '            '    Item(L).Remove()
        '            'Next
        '            '''Histroy_End(gLab.Desk(DeskID))

        '        Case "新規"
        '            '----------------------------------------------------
        '            '■何もしない
        '            Log_Show(2, "新規/変更はできません")

        '        Case "変更"
        '            '----------------------------------------------------
        '            '■何もしない
        '            Log_Show(2, "新規/変更はできません")
        '    End Select

        'End With

        ''現在開いているものと一致するなら再描画
        ''If Flag_RID <> 0 Then Proj_Show_List(Proj, gDic_TblID)
        ''If Flag_DID <> 0 Then Proj_Show_Atr(Proj, gAtr_NID, gAtr_LID)



    End Sub

    'Private Function ListView_Get_TextData(ByRef LV As ListView) As String
    '    '選択されているListViewのアイテムをテキスト化する
    '    '-----------------------------------------
    '    Dim Lv_data As String
    '    Dim Item() As ListViewItem
    '    Lv_data = ""

    '    With LV
    '        '対象ノード非選択
    '        If .SelectedItems.Count = 0 Then Return Lv_data

    '        '対象ノード
    '        ReDim Item(.SelectedItems.Count - 1)
    '        For L = 0 To .SelectedItems.Count - 1
    '            Item(L) = .SelectedItems(L)

    '            For K = 0 To .SelectedItems(L).SubItems.Count - 1
    '                Lv_data &= .SelectedItems(L).SubItems(K).Text & Chr(9)
    '            Next
    '            Progres_Show(L + 1, .SelectedItems.Count, False)
    '            Lv_data &= Chr(13) & Chr(10)
    '        Next

    '    End With
    '    Return Lv_data
    'End Function

    Private Sub Menu_Act_DicAtr(ByRef Proj As Lab_Type, ByRef Cmd() As String)
        'メニュー選択時のアクションまとめ ATR
        '-----------------------------------------
        '画面上の操作のみで、HISTORYの対象外

        'Dim DeskID As Integer
        'DeskID = 1


        'Dim Item() As ListViewItem


        'With ListView2
        '    '対象ノード非選択
        '    If .SelectedItems.Count = 0 Then Return

        '    '対象ノード
        '    ReDim Item(.SelectedItems.Count - 1)
        '    For L = 0 To .SelectedItems.Count - 1
        '        Item(L) = .SelectedItems(L)
        '    Next


        '    Select Case Cmd(1)

        '        Case "コピー"
        '            gBff_Act_Type = "COPY_ATR"
        '            gBff_Node = Nothing
        '            gBff_Item = Item.Clone

        '        Case "切り取り"
        '            gBff_Act_Type = "MOVE_ATR"
        '            gBff_Node = Nothing
        '            gBff_Item = Item.Clone

        '            ListView_Del(ListView2)

        '        Case "貼り付け"

        '            Select Case gBff_Act_Type
        '                Case "COPY_NODE", "MOVE_NODE" '同じあつかい　元ノードは消さない
        '                    '----------------------------------------------------

        '                    .Items.Insert(Item(0).Index, Item(0).Clone)
        '                    Dim Itm As String
        '                    '    Itm = Proj_ATR_Make_Item_byNodeName(gBff_Node.Text)
        '                   ' Proj_Atr_Rep_Item(.SelectedItems(0), Itm)

        '                Case "COPY_ITEM", "MOVE_ITEM"
        '                    '----------------------------------------------------
        '                    '■何もしない
        '                    Log_Show(2, "指定のオブジェクトは貼り付けできません")

        '                Case "MOVE_ATR", "COPY_ATR" '同じあつかい　元ノードは消されてる
        '                    For K = 0 To gBff_Item.Length - 1
        '                        .Items.Insert(Item(0).Index, gBff_Item(K).Clone)
        '                    Next

        '            End Select

        '        Case "削除"
        '            ListView_Del(ListView2)

        '        Case "新規"
        '            ''If Cmd.Length <3 Then Return
        '            ''Select Case Cmd(2)
        '            ''    Case "S", "C", "G"
        '            ''        Proj_Atr_New_Item(ListView2, Cmd(2))
        '            ''End Select
        '            'Proj_Atr_New_Item(ListView2, "S")

        '        Case "変更"
        '            'If Cmd.Length < 3 Then Return
        '            'Dim Itm As String
        '            'Itm = Proj_ATR_Make_Item_byMenuCmd(Cmd)
        '            'Proj_Atr_Rep_Item(.SelectedItems(0), Itm)
        '    End Select


        'End With

    End Sub




    Private Sub Menu_Set(ByRef Menu As ContextMenuStrip, ByRef Tx As String, ByRef Kx As String)

        'メニューに追加
        '-----------------------------------------
        Dim dx() As String
        Dim tag() As String

        If Tx = "" Then Return
        dx = Split(Tx, "/")
        tag = Split(Kx, "/")
        Dim Itm(dx.Length) As ToolStripMenuItem
        For K = 0 To dx.Length - 1
            If dx(K) <> "" Then
                Itm(K) = New ToolStripMenuItem
                Itm(K).Text = dx(K)
                Itm(K).Tag = tag(K)
                Menu.Items.Add(Itm(K))
            End If
        Next

        Return

    End Sub



    Private Sub Owner_Set_Env(ByRef Lab As Lab_Type, ID As Integer)

        Dim Ds As DataSet
        Dim Cnd As String
        Cnd = "ID = " & ID.ToString


        'With Lab.Person(0)
        '    DBAcs_Init2(.Db, "OwnerDB", "OWNER", "OWNER")
        '    Ds = DBAcs_Get_DataS(.Db, "Name,Mail,Tel,Prof,Image", Cnd)
        '    If Ds.Tables(0).Rows.Count > 0 Then
        '        .ID = ID
        '        .Name = Trim(Ds.Tables(0).Rows(0).Item(0))
        '        .Mail = Trim(Ds.Tables(0).Rows(0).Item(1))
        '        .Tel = Trim(Ds.Tables(0).Rows(0).Item(2))
        '        .Prof = Trim(Ds.Tables(0).Rows(0).Item(3))
        '        .Image = Trim(Ds.Tables(0).Rows(0).Item(4))
        '        Lab.Owner = .Name
        '        Lab.OwnerID = ID
        '    End If
        'End With




    End Sub



    Private Function Person_Set_NewData(ByRef Lab As Lab_Type) As Long
        'OWNER プロファイル登録
        '-------------------
        'Dim key, tmpD As String
        'Dim Person As Person_Type
        'With Person
        '    .Name = "Unknown"
        '    .Tel = "Tel"
        '    .Mail = "Mail"
        '    .Prof = "Prof"
        '    .Address = "Address"
        'End With
        'key = "(Name)"()
        ''DB登録
        'Proj.ProjID = ProjDB_Insert(Lab, "Proj", "(OwnerID,Name)", "(" & Lab.OwnerID.ToString & ",'新しいプロジェクト')")
        ''With Lab.Person(0)
        ''    DBAcs_Init2(.Db, "OwnerDB", "OWNER", "OWNER")
        ''    tmpD = ""
        ''    .ID = Owner_Get_ID()
        ''    If .ID = 0 Then

        ''        key = Mid(.Db.TBL_Field, Len("ID,") + 1)
        ''        tmpD &= "'" & .Name & "',"
        ''        tmpD &= "'" & .Mail & "',"
        ''        tmpD &= "'" & .Tel & "',"
        ''        tmpD &= "'" & .Prof & "',"
        ''        tmpD &= "'" & .Image & "'"
        ''        .ID = DBAcs_Insert(.Db, .Db.TBL_Name, key, tmpD) 'DB登録
        ''        Lab.OwnerID = .ID
        ''        Owner_Save_ID(.ID) 'ファイル保存
        ''    Else
        ''        tmpD &= "Name = '" & .Name & "',"
        ''        tmpD &= "Mail = '" & .Mail & "',"
        ''        tmpD &= "Tel = '" & .Tel & "',"
        ''        tmpD &= "Prof = '" & .Prof & "',"
        ''        tmpD &= "Image = '" & .Image & "'"
        ''        DBAcs_Update_Data_ByKey(.Db, .Db.TBL_Name, .ID, tmpD) 'DB修正
        ''        Lab.OwnerID = .ID

        ''    End If



        ''End With

        ''Owner_Show_Data(Lab)


    End Function

    'Private Sub Desk_Set_NewData(ByRef Lab As Lab_Type)
    '    'OWNER プロファイル登録
    '    '-------------------
    '    Dim key, tmpD As String

    '    'With Lab.Desk
    '    '    DBAcs_Init2(.DBACS, "OwnerDB", "OWNER", "OWNER")
    '    '    tmpD = ""
    '    '    .ID = Owner_Get_ID()
    '    '    If .ID = 0 Then

    '    '        key = Mid(.Db.TBL_Field, Len("ID,") + 1)
    '    '        tmpD &= "'" & .Name & "',"
    '    '        tmpD &= "'" & .Mail & "',"
    '    '        tmpD &= "'" & .Tel & "',"
    '    '        tmpD &= "'" & .Prof & "',"
    '    '        tmpD &= "'" & .Image & "'"
    '    '        .ID = DBAcs_Insert(.Db, .Db.TBL_Name, key, tmpD) 'DB登録
    '    '        Lab.OwnerID = .ID
    '    '        Owner_Save_ID(.ID) 'ファイル保存
    '    '    Else
    '    '        tmpD &= "Name = '" & .Name & "',"
    '    '        tmpD &= "Mail = '" & .Mail & "',"
    '    '        tmpD &= "Tel = '" & .Tel & "',"
    '    '        tmpD &= "Prof = '" & .Prof & "',"
    '    '        tmpD &= "Image = '" & .Image & "'"
    '    '        DBAcs_Update_Data_ByKey(.Db, .Db.TBL_Name, .ID, tmpD) 'DB修正
    '    '        Lab.OwnerID = .ID

    '    '    End If



    '    'End With

    '    Owner_Show_Data(Lab)


    'End Sub

    'Private Sub DBAcs_Init(ByRef Lab As Lab_Type, ByRef Key As String)
    '    'DBアクセスの為の型とエントリ名を設定

    '    With Lab
    '        Select Case Key

    '            Case "Proj"
    '                .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT,State INTEGER ,OwnerID int,Owner  TEXT ,Name  TEXT ,Title  TEXT,Purpose text,Summary text,Hypothesis Text,Problem text)")
    '                .DBACS_Field.Add(Key, "ID,State,OwnerID,Owner,Name,Title,Purpose,Summary,Hypothesis,Problem")


    '            Case "Person"
    '                .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT ,  Name  TEXT ,Mail  TEXT,Tel   TEXT , Prof text, Image  TEXT ,State int)")
    '                .DBACS_Field.Add(Key, "ID,Name,Mail,Tel,Prof,Image ,State")


    '            Case "Desk"
    '                .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT  ,OwnerID int,ProjID int,Page int,Type  TEXT ,Title  TEXT ,State int,ShiftX int,ShiftY int,ZM int)")
    '                .DBACS_Field.Add(Key, "ID,OwnerID,ProjID,Page,Type,Title ,State,ShiftX,ShiftY,ZM")

    '            Case "Box"
    '                .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT  ,DeskID int,ProjID int,OwnerID int,BluePrint text ,State int)")
    '                .DBACS_Field.Add(Key, "ID,DeskID,ProjKey,OwnerID,BluePrint ,State")


    '            Case "Tbox"
    '                .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT  ,BoxID int,Num int,Text Text)")
    '                .DBACS_Field.Add(Key, "ID,BoxID,Num,Text ,State")

    '            'Case "Draw"
    '            '    .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT  ,BoxID int,Num int,Text Text ,State int,Path string)")
    '            '    .DBACS_Field.Add(Key, "ID,BoxID,Num,Text , State,Path")

    '            Case "Pbox"
    '                .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT  ,BoxID int,Num int,Pict TEXT ,State int)")
    '                .DBACS_Field.Add(Key, "ID,BoxID,Num,Pict ,State")

    '            Case "BDat"
    '                .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT  ,BoxID int,Num int,Dat BLOB ,State int,Type String,Path string)")
    '                .DBACS_Field.Add(Key, "ID,BoxID,Num,Dat ,State,Type,Path")

    '            'Key= "Webbox"
    '            '    .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT  ,BaseID int,Type  TEXT ,Style int,Text Text)"
    '            '    .DBACS_Field.Add(Key,"ID,BaseID,Type,Path,Style,Text"
    '            'DB_Check(Lab, key)

    '            Case "Element"
    '                .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT  ,ProjID int,Type  TEXT ,Name  TEXT,Path TEXT ,State int)")
    '                .DBACS_Field.Add(Key, "ID,ProjID,Type,Name,Path ,State")

    '            Case "Role"
    '                .DBACS_Type.Add(Key, "(ID INTEGER  PRIMARY KEY  AUTOINCREMENT  ,LID INTEGER ,PID int,Role  TEXT ,State INTEGER )")
    '                .DBACS_Field.Add(Key, "ID,LID,PID,Role ,State")

    '        End Select

    '    End With
    '    'SQLサーバ
    '    'With Lab
    '    '    Select Case Key

    '    '        Case "Proj"
    '    '            .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1),State int ,OwnerID int,Owner TEXT(100) ,Name nvarchar(200) ,Title nvarchar(200),Purpose text,Summary text,Hypothesis Text,Problem text)")
    '    '            .DBACS_Field.Add(Key, "ID,State,OwnerID,Owner,Name,Title,Purpose,Summary,Hypothesis,Problem")


    '    '        Case "Person"
    '    '            .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1),  Name nchar(200) ,Mail nchar(100),Tel  nvarchar(100) , Prof text, Image nvarchar(100) ,State int)")
    '    '            .DBACS_Field.Add(Key, "ID,Name,Mail,Tel,Prof,Image ,State")


    '    '        Case "Desk"
    '    '            .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1) ,OwnerID int,ProjID int,Page int,Type nvarchar(100) ,Title nvarchar(300) ,State int)")
    '    '            .DBACS_Field.Add(Key, "ID,OwnerID,ProjID,Page,Type,Title ,State")

    '    '        Case "Box"
    '    '            .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1) ,DeskID int,ProjID int,OwnerID int,BluePrint text ,State int)")
    '    '            .DBACS_Field.Add(Key, "ID,DeskID,ProjKey,OwnerID,BluePrint ,State")


    '    '        Case "Tbox"
    '    '            .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1) ,BoxID int,Num int,Text Text)")
    '    '            .DBACS_Field.Add(Key, "ID,BoxID,Num,Text ,State")

    '    '        Case "Draw"
    '    '            .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1) ,BoxID int,Num int,Text Text ,State int)")
    '    '            .DBACS_Field.Add(Key, "ID,BoxID,Num,Text , State")

    '    '        Case "Pbox"
    '    '            .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1) ,BoxID int,Num int,Pict TEXT ,State int)")
    '    '            .DBACS_Field.Add(Key, "ID,BoxID,Num,Pict ,State")

    '    '        Case "BDat"
    '    '            .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1) ,BoxID int,Num int,Dat varbinary(max) ,State int)")
    '    '            .DBACS_Field.Add(Key, "ID,BoxID,Num,Dat ,State")

    '    '        'Key= "Webbox"
    '    '        '    .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1) ,BaseID int,Type nvarchar(100) ,Style int,Text Text)"
    '    '        '    .DBACS_Field.Add(Key,"ID,BaseID,Type,Path,Style,Text"
    '    '        'DB_Check(Lab, key)

    '    '        Case "Element"
    '    '            .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1) ,ProjID int,Type nvarchar(100) ,Name nvarchar(100),Path nvarchar(500) ,State int)")
    '    '            .DBACS_Field.Add(Key, "ID,ProjID,Type,Name,Path ,State")

    '    '        Case "Role"
    '    '            .DBACS_Type.Add(Key, "(ID int  primary key IDENTITY (1,1) ,LID int ,PID int,Role nvarchar(100) ,State int )")
    '    '            .DBACS_Field.Add(Key, "ID,LID,PID,Role ,State")

    '    '    End Select

    '    'End With
    '    DB_Check(Lab, Key)


    'End Sub


    'Private Function DBAcs_Del_Dat(ByRef Db As DBACS_Type, ByRef Cnd As String) As Integer
    '    'DBレコードの削除
    '    '----------------
    '    'DB 指定Key の　値　を格納
    '    '--------------------------
    '    Dim tmpR As String
    '    Dim DBCon As SqlConnection
    '    Dim DBCom As SqlCommand
    '    Dim TmpCmd As String

    '    With Db
    '        DBCon = New SqlConnection(.DB_Path & "initial catalog = " & .DB_Name)
    '        TmpCmd = "DELETE   FROM " & .TBL_Name & "  WHERE " & Cnd
    '    End With

    '    DBCom = New SqlCommand(TmpCmd, DBCon)
    '    DBCom.CommandTimeout = 0

    '    DBCom.Connection.Open()


    '    For I As Integer = 1 To 50
    '        Try
    '            tmpR = DBCom.ExecuteScalar()

    '            Exit For
    '        Catch ex As Exception
    '            Log_Show(2, "DB登録エラー" & I & ex.Message)
    '            For J As Integer = 1 To 100
    '                System.Windows.Forms.Application.DoEvents()
    '            Next
    '        End Try
    '    Next

    '    DBCom.Connection.Close()
    '    DBCon.Close()
    '    DBCon.Dispose()




    '    Return tmpR
    'End Function



    'Private Function DBAcs_Tbl_CheckExist(ByRef Db As DBACS_Type) As Integer

    '    Dim TmpR As Integer
    '    With Db
    '        TmpR = DB_Tbl_CheckExist(.DB_Path, .DB_Name, .TBL_Name, .TBL_Type)
    '    End With


    '    Return TmpR
    'End Function
    'Private Sub DB_Check(ByRef Lab As Lab_Type, Key As String)
    '    With Lab

    '        Select Case .Db_Type
    '            Case "SQLite"
    '                DB_Tbl_CheckExistSQLite(Lab, Key, .DBACS_Type(Key))'TBL存在チェック&ないなら作成
    '            Case "MSSQL", "MSSQLW"
    '                DB_Check_Exist_DB_MSSQL(.Db_Type, .DbPath, .Db_Type & Key) 'DB 存在チェック&ないなら作成
    '                DB_Tbl_CheckExistMSSQL(.Db_Type, .DbPath, .Db_Type & Key, "TBL_" & Key, .DBACS_Type(Key)) 'TBL存在チェック&ないなら作成
    '        End Select

    '    End With

    'End Sub

    'Private Sub DB_Check_MSSQL(ByRef Lab As Lab_Type, Key As String)
    '    With Lab
    '        DB_Check_Exist_DB_MSSQL(.Db_Type, .DbPath, .Db_Type & Key) 'DB 存在チェック&ないなら作成
    '        DB_Tbl_CheckExistMSSQL(.Db_Type, .DbPath, .Db_Type & Key, "TBL_" & Key, .DBACS_Type(Key)) 'TBL存在チェック&ないなら作成
    '    End With

    'End Sub


    'Private Function DB_Check_Exist_DB(ByRef DB_type As String, ByRef Db_Path As String) As Integer
    '    'DBがないときは空DBをコピー
    '    Try
    '        If System.IO.File.Exists(gDB_Path) Then
    '            Return True
    '        Else
    '            System.IO.File.Copy(System.Windows.Forms.Application.StartupPath & "\ENV\GI.sqlite3", gDB_Path)
    '        End If
    '        Return True

    '    Catch ex As Exception
    '        Log_Show(1, ex.Message)
    '    End Try

    '    Return False

    'End Function

    'Private Function DB_Check_Exist_DB_MSSQL(ByRef DB_type As String, ByRef Db_Path As String, ByRef DB_Name As String) As Integer

    '    Dim DBCon, cn As SqlConnection
    '    Dim DBCom, Cmd As SqlCommand
    '    Dim TmpR As Integer

    '    TmpR = True
    '    'DBの存在チェック
    '    'cn = New SqlConnection("Server=.\SQLEXPRESS;" & "Database=master;" & "integrated security=true")
    '    Log_Show(1, "DB存在チェック")
    '    Try
    '        cn = New SqlConnection(Db_Path)
    '        Cmd = New SqlCommand("SELECT COUNT(*) FROM sysdatabases WHERE name='" & DB_Name & "'", cn)

    '        Cmd.Connection.Open()

    '        TmpR = True
    '    Catch ex As Exception
    '        Log_Show(1, ex.Message)
    '        TmpR = False
    '    End Try

    '    'データベースが存在しない
    '    If (Cmd.ExecuteScalar() = 0) Then

    '        'データベースを作成
    '        DBCon = New SqlConnection(Db_Path)
    '        DBCom = New SqlCommand("create database " & DB_Name, DBCon)

    '        DBCom.CommandTimeout = 0

    '        DBCom.Connection.Open()

    '        Try
    '            DBCom.ExecuteNonQuery()
    '            TmpR = True
    '        Catch ex As Exception
    '            Log_Show(1, ex.Message)
    '            TmpR = False
    '        End Try
    '        DBCom.Connection.Close()
    '        DBCon.Close()
    '    End If

    '    cn.Close()

    '    Return TmpR


    'End Function
    'Private Function DB_Tbl_CheckExistSQLite(ByRef Lab As Lab_Type, ByRef Tbl_Name As String, ByRef Tbl_Type As String) As Integer

    '    'テーブルの存在チェック
    '    '---------------------------------
    '    Try
    '        Dim TmpR As Integer
    '        TmpR = False
    '        If Tbl_Name = "" Then Return TmpR


    '        ' コネクション作成
    '        Dim con As SQLiteConnection
    '        con = New SQLiteConnection(GetConnectionString(Lab.DbPath))
    '        con.Open()

    '        'テーブルチェック
    '        Using cmd = con.CreateCommand()
    '            cmd.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE TYPE = 'table' AND name= '" & Tbl_Name & "'"
    '            Dim tmpC As Integer = cmd.ExecuteScalar()
    '            If (tmpC <> 0) Then
    '                con.Close()
    '                Return True
    '            End If

    '        End Using

    '        ' テーブル作成SQL
    '        Dim TmpCmd As String
    '        TmpCmd = "CREATE TABLE " & Tbl_Name & Tbl_Type
    '        Using cmd = con.CreateCommand()
    '            cmd.CommandText = TmpCmd
    '            cmd.ExecuteNonQuery()
    '        End Using
    '        con.Close()
    '        Return True

    '    Catch ex As Exception

    '        Log_Show(1, ex.Message)
    '    End Try
    '    Return False

    'End Function


    'Private Function DB_Tbl_CheckExistMSSQL(ByRef DB_Type As String, ByRef Db_Path As String, ByRef DB_Name As String, ByVal Tbl_Name As String, ByRef Tbl_Type As String) As Integer

    '    'テーブルの存在チェック
    '    '---------------------------------
    '    '戻り値　存在　2　新規　1　エラー0　
    '    Dim TmpR As Integer
    '    TmpR = 0
    '    If Tbl_Name = "" Then Return TmpR

    '    Dim DBCon, cn As SqlConnection
    '    Dim DBCom, Cmd As SqlCommand



    '    'テーブルの存在チェック
    '    cn = New SqlConnection(Db_Path & ";Initial Catalog =" & DB_Name)

    '    Cmd = New SqlCommand("Select COUNT(*) FROM sys.objects WHERE object_id = OBJECT_ID( '" & Tbl_Name & "')", cn)

    '    For J = 0 To 50
    '        Try
    '            Cmd.Connection.Open()
    '            Exit For

    '        Catch ex As Exception

    '            Log_Show(1, ex.Message)
    '        End Try

    '    Next

    '    'SELECT COUNT(*) FROM sys.objects WHERE object_id = OBJECT_ID('dbo.Table')
    '    'テーブルが存在しない
    '    If (Cmd.ExecuteScalar() = 0) Then

    '        'テーブルを作成
    '        Dim TmpCmd As String
    '        TmpCmd = " create table " & Tbl_Name & Tbl_Type

    '        DBCon = New SqlConnection(Db_Path & ";Initial Catalog =" & DB_Name)
    '        DBCom = New SqlCommand(TmpCmd, DBCon)
    '        Try


    '            DBCom.Connection.Open()
    '            DBCom.ExecuteNonQuery()
    '            DBCom.Connection.Close()
    '            DBCon.Close()
    '            '  DBCon.Dispose()

    '            TmpR = 1

    '        Catch ex As Exception

    '            Log_Show(1, ex.Message)
    '        End Try
    '    Else
    '        TmpR = 2
    '    End If
    '    cn.Close()

    '    Return TmpR

    'End Function

    'Private Function DBACS_Get_ID_ByCnd(ByRef DB As DBACS_Type, ByRef Cnd As String) As Integer()
    '    'DB 指定条件 の　ID　を返却
    '    '--------------------------
    '    Dim XR() As Integer
    '    ReDim XR(0)
    '    Dim DBCon As SqlConnection
    '    Dim DBCom As SqlCommand
    '    Dim DBReader As SqlClient.SqlDataReader

    '    Dim TmpCmd, TmpDB As String

    '    With DB
    '        TmpDB = .DB_Path & ";Initial Catalog = " & .DB_Name
    '        TmpCmd = "Select * FROM [" & .DB_Name & "].[dbo].[" & .TBL_Name & "] WHERE  " & Cnd
    '    End With

    '    DBCon = New SqlConnection(TmpDB)
    '    DBCom = New SqlCommand(TmpCmd, DBCon)

    '    DBCom.CommandTimeout = 0
    '    DBCom.Connection.Open()

    '    For I As Integer = 1 To 50
    '        Try
    '            DBReader = DBCom.ExecuteReader()

    '            While DBReader.Read
    '                XR(XR.Length - 1) = Val(DBReader(0) & ".0")
    '                ReDim Preserve XR(XR.Length)
    '            End While
    '            DBReader.Read()


    '            Exit For
    '        Catch ex As Exception
    '            Log_Show(2, "DBエラー" & I & ex.Message)
    '            For J As Integer = 1 To 50
    '                System.Windows.Forms.Application.DoEvents()
    '            Next
    '        End Try
    '    Next
    '    DBCom.Dispose()
    '    DBReader.Close()
    '    DBCon.Close()
    '    DBCon.Dispose()

    '    Return XR



    'End Function



    'Private Sub DBAcs_Update_Data_ByKey(ByRef DB As DBACS_Type, ByRef TBL_Name As String, ByRef ID As Integer, ByRef Cmd As String)
    '    'DB 指定Key の　値　を格納
    '    '--------------------------
    '    Dim tmpR As String
    '    Dim DBCon As SqlConnection
    '    Dim DBCom As SqlCommand
    '    Dim TmpCmd As String


    '    TmpCmd = "UPDATE " & TBL_Name & " Set  " & Cmd & "  WHERE ID = " & ID

    '    With DB
    '        DBCon = New SqlConnection(.DB_Path & ";Initial Catalog =" & .DB_Name)
    '    End With

    '    DBCom = New SqlCommand(TmpCmd, DBCon)
    '    DBCom.Connection.Open()


    '    For I As Integer = 1 To 50
    '        Try
    '            tmpR = DBCom.ExecuteScalar()

    '            Exit For
    '        Catch ex As Exception
    '            Log_Show(2, "DB登録エラー" & I & ex.Message)
    '            For J As Integer = 1 To 100
    '                System.Windows.Forms.Application.DoEvents()
    '            Next
    '        End Try
    '    Next

    '    DBCom.Connection.Close()
    '    DBCon.Close()
    '    DBCon.Dispose()




    'End Sub
    'Private Function DBAcs_Insert(ByRef Lab As Lab_Type, ByRef Name As String, ByRef Key As String, ByRef Val As String) As Long
    '    'DB 指定Key の　値　を格納
    '    '--------------------------
    '    Dim tmpR As Integer
    '    Try
    '        Dim TmpCmd As String

    '        TmpCmd = "INSERT INTO " & Name
    '        TmpCmd &= " ( " & Key & ")   VALUES ( " & Val & ")"

    '        Using con = New SQLiteConnection(GetConnectionString(Lab.DbPath))
    '            con.Open()
    '            Using cmd = con.CreateCommand()
    '                cmd.CommandText = TmpCmd
    '                cmd.ExecuteNonQuery()
    '            End Using
    '            Using cmd = con.CreateCommand()
    '                'cmd.CommandText = "SELECT last_insert_rowid();"
    '                cmd.CommandText = "Select max(ID) FROM " & Name
    '                tmpR = cmd.ExecuteScalar()
    '            End Using
    '            con.Close()
    '            Return tmpR
    '        End Using

    '    Catch ex As Exception
    '        Log_Show(1, ex.Message)
    '        Return False
    '    End Try
    '    Return tmpR
    'End Function

    'Private Function DBAcs_Insert_MSSQL(ByRef Lab As Lab_Type, ByRef Name As String, ByRef Key As String, ByRef Val As String) As Long
    '    'DB 指定Key の　値　を格納
    '    '--------------------------
    '    Dim tmpR As Integer

    '    Dim DBCon As SqlConnection
    '    Dim DBCom As SqlCommand
    '    Dim TmpCmd As String
    '    tmpR = -1

    '    With Lab

    '        DBCon = New SqlConnection(.DbPath & ";Initial Catalog =" & .Db_Type & Name)

    '        TmpCmd = "INSERT INTO " & "TBL_" & Name
    '        TmpCmd &= " ( " & Key & ")  OUTPUT INSERTED.ID  VALUES ( " & Val & ")"
    '        TmpCmd &= "; Select SCOPE_IDENTITY();"
    '    End With

    '    DBCom = New SqlCommand(TmpCmd, DBCon)
    '    DBCom.Connection.Open()



    '    For I As Integer = 1 To 50
    '        Try
    '            tmpR = CInt(DBCom.ExecuteScalar())

    '            Exit For
    '        Catch ex As Exception
    '            Log_Show(1, "DB登録エラー" & I & ex.Message)
    '            For J As Integer = 1 To 100
    '                System.Windows.Forms.Application.DoEvents()
    '            Next
    '        End Try
    '    Next

    '    DBCom.Connection.Close()
    '    DBCon.Close()
    '    DBCon.Dispose()

    '    Return tmpR
    'End Function
    Private Sub Box_Update(ByRef Lab As Lab_Type, ByRef DeskID As Integer, ByRef BoxID As Integer, ByRef Key As String, ByRef Dat As String)
        Dim Blist() As DataRow
        Dim dtRow As DataRow
        With Lab

            Blist = .BoxList.Select("BoxID = '" & BoxID.ToString & "' and DeskID='" & DeskID.ToString & "'")
            If Blist.Count = 0 Then Return
            dtRow = Blist(0)
            dtRow(Key) = Dat
            Box_Save_Info(dtRow)

        End With




    End Sub
    'Private Function DBAcs_Update(ByRef Lab As Lab_Type, ByRef Name As String, ByRef Key As String, ByRef Val As String) As Integer
    '    'DB 指定Key の　値　を格納
    '    '--------------------------
    '    Try
    '        Dim TmpCmd As String
    '        Return False

    '        TmpCmd = "UPDATE " & Name
    '        TmpCmd &= "  SET " & Key & " WHERE  " & Val

    '        Using con = New SQLiteConnection(GetConnectionString(Lab.DbPath))
    '            con.Open()
    '            Using cmd = con.CreateCommand()
    '                cmd.CommandText = TmpCmd
    '                cmd.ExecuteNonQuery()
    '            End Using
    '            con.Close()
    '        End Using


    '    Catch ex As Exception
    '        Log_Show(1, ex.Message)
    '        Return False
    '    End Try

    '    Return True
    'End Function
    'Private Function DBAcs_UpdateMSSQL(ByRef Lab As Lab_Type, ByRef Name As String, ByRef Key As String, ByRef Val As String) As Integer
    '    'DB 指定Key の　値　を格納
    '    '--------------------------
    '    Dim tmpR As Integer

    '    Dim DBCon As SqlConnection
    '    Dim DBCom As SqlCommand
    '    Dim TmpCmd As String
    '    tmpR = -1

    '    With Lab

    '        DBCon = New SqlConnection(.DbPath & ";Initial Catalog =" & .Db_Type & Name)
    '        TmpCmd = "UPDATE " & "TBL_" & Name
    '        TmpCmd &= "  SET " & Key & " WHERE  " & Val
    '        '  TmpCmd &= "; Select SCOPE_IDENTITY();"

    '    End With

    '    DBCom = New SqlCommand(TmpCmd, DBCon)
    '    DBCom.Connection.Open()

    '    For I As Integer = 1 To 50
    '        Try
    '            tmpR = CInt(DBCom.ExecuteScalar())

    '            Exit For
    '        Catch ex As Exception
    '            Log_Show(1, "DB登録エラー" & I & ex.Message)
    '            For J As Integer = 1 To 100
    '                System.Windows.Forms.Application.DoEvents()
    '            Next
    '        End Try
    '    Next

    '    DBCom.Connection.Close()
    '    DBCon.Close()
    '    DBCon.Dispose()

    '    Return tmpR
    'End Function
    'Private Function DBAcs_Get_MaxID(ByRef Lab As Lab_Type, ByRef Name As String, ByRef Key As String) As Long

    '    Dim tmpR As Integer

    '    Dim DBCon As SqlConnection
    '    Dim DBCom As SqlCommand
    '    Dim TmpCmd As String
    '    tmpR = -1

    '    With Lab

    '        DBCon = New SqlConnection(.DbPath & ";Initial Catalog =" & .DBname & Name)
    '        TmpCmd = "SELECT MAX(" & Key & ") FROM " & "TBL_" & Name

    '    End With

    '    DBCom = New SqlCommand(TmpCmd, DBCon)
    '    DBCom.Connection.Open()



    '    For I As Integer = 1 To 50
    '        Try
    '            tmpR = CLng(DBCom.ExecuteScalar())

    '            Exit For
    '        Catch ex As Exception
    '            Log_Show(1, "DB登録エラー" & I & ex.Message)
    '            For J As Integer = 1 To 100
    '                System.Windows.Forms.Application.DoEvents()
    '            Next
    '        End Try
    '    Next

    '    DBCom.Connection.Close()
    '    DBCon.Close()
    '    DBCon.Dispose()

    '    Return tmpR

    'End Function

    Private Function DBAcs_Get_DataS(ByRef Lab As Lab_Type, Name As String, Action As String, ByRef Condition As String) As DataSet

        'DB 指定Key の　値　を格納
        '--------------------------
        Try
            Dim adapter = New SQLiteDataAdapter()
            '  Dim dtTbl As New DataTable()
            Dim ds As DataSet
            ds = New DataSet()

            Dim TmpCmd As String
            TmpCmd = "Select " & Action & " FROM " & Name & " WHERE " & Condition

            Using con = New SQLiteConnection(GetConnectionString(Lab.DbPath))
                con.Open()
                Using cmd = con.CreateCommand()
                    cmd.CommandText = TmpCmd
                    adapter = New SQLiteDataAdapter(TmpCmd, con)
                    '  adapter.Fill(dtTbl)
                    adapter.Fill(ds)
                End Using
                adapter.Dispose()
                con.Close()
            End Using

            Return ds


        Catch ex As Exception
            Log_Show(1, ex.Message)

        End Try

    End Function

    'Private Function DBAcs_Get_DataS_MSSQL(ByRef Lab As Lab_Type, Name As String, Action As String, ByRef Condition As String) As DataSet

    '    'DB 指定Key の　値　を格納
    '    '--------------------------
    '    Dim tmpR As Integer

    '    Dim DBCon As SqlConnection
    '    Dim TmpCmd As String
    '    tmpR = -1
    '    With Lab
    '        TmpCmd = "Select " & Action & " FROM [" & .Db_Type & Name & "].[dbo].[" & "TBL_" & Name & "] WHERE " & Condition
    '        DBCon = New SqlConnection(.DbPath & ";Initial Catalog =" & .Db_Type & Name)
    '        DBCon.Open()

    '    End With
    '    'DBCom = New SqlCommand(TmpCmd, DBCon)

    '    Dim adapter As SqlDataAdapter
    '    Dim ds As DataSet
    '    ds = New DataSet()
    '    adapter = New SqlDataAdapter()
    '    'データの取得

    '    For I As Integer = 1 To 50
    '        Try

    '            adapter.SelectCommand = New SqlCommand(TmpCmd, DBCon)

    '            adapter.SelectCommand.CommandType = CommandType.Text

    '            adapter.Fill(ds)
    '            Exit For
    '        Catch ex As Exception
    '            Log_Show(1, "エラー" & I & ex.Message)
    '            For J As Integer = 1 To 100
    '                System.Windows.Forms.Application.DoEvents()
    '            Next
    '        End Try
    '    Next
    '    DBCon.Close()
    '    DBCon.Dispose()
    '    adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey

    '    DBCon.Close()

    '    Return ds



    'End Function



    Private Function File_Set_byDir(ByRef FolderName As String, ByRef Key As String, ByRef TmpL() As String) As Integer
        '拡張子指定でファイル名リストを作成する
        '---------------------------------------
        Dim FileName As String
        Dim Lmax As Integer
        '指定のフォルダがない
        If System.IO.Directory.Exists(FolderName) = False Then
            Return 0
        End If

        '指定のフォルダ内の指定の拡張子のファイルを全て列挙する
        If TmpL Is Nothing Then
            ReDim TmpL(0)
            Lmax = 0
        Else
            Lmax = TmpL.Length
        End If

        If FolderName.Length > 1 Then
            '指定の拡張子のファイルだけ取得する場合
            For Each FileName In System.IO.Directory.GetFiles(FolderName, "*" & Key, System.IO.SearchOption.TopDirectoryOnly)

                ReDim Preserve TmpL(Lmax)
                TmpL(Lmax) = FileName
                Lmax = Lmax + 1
            Next
        End If

        Return Lmax

    End Function





    Private Sub Log_Show(ByRef ID As Integer, ByRef TmpM As String)

        Select Case ID

            Case 1
                If TmpM = "DEL" Then
                    ToolStripStatusLabel1.Text = ""
                Else

                    ToolStripStatusLabel1.Text = TmpM & Chr(13) & Chr(10)
                End If

            Case 2

                ToolStripStatusLabel2.Text = TmpM


            Case 5
                ToolStripStatusLabel3.Text = TmpM


        End Select
        System.Windows.Forms.Application.DoEvents()


    End Sub



    Private Sub Proj_Rename(ByRef lab As Lab_Type, NewName As String)
        If lab.ProjOwnerID <> gX_ownerID Then
            MsgBox（"オーナーではありません"）
            Return
        End If
        Dim Plist() As DataRow
        Dim dtRow As DataRow
        With gLab
            .ProjName = NewName
            Plist = .ProjList.Select("ProjKey ='" & .ProjKey & "'")

            dtRow = Plist(0)
            dtRow("Title") = NewName
            Proj_Save_Info(dtRow) 'ファイル書き込み
        End With



    End Sub

    Private Sub Desk_Check_Rename(ByRef lab As Lab_Type)
        With DataGridView3
            If .CurrentCell Is Nothing Then Return '選択されたセルがない
            If .CurrentCell.Tag = "" Then Return

            Dim Dlist() As DataRow
            Dim dtRow As DataRow
            ' Dim Key As String = lab.DesKTop(Val(.CurrentCell.Tag)).Tag
            'Dlist = lab.DeskList.Select("DeskID ='1'")
            Dlist = lab.DeskList.Select("DeskID ='" & lab.DesKTop(Val(.CurrentCell.Tag)).Tag & "'")
            If Dlist.Count = 0 Then Return
            dtRow = Dlist(0)
            If dtRow("OwnerID") <> gX_ownerID Then
                MsgBox（"オーナーではありません"）
                Return
            End If
            lab.Flg_DeskRename = True
            .CurrentCell.ReadOnly = False
            .BeginEdit(True)
        End With

    End Sub

    Private Function Desk_Rename(ByRef lab As Lab_Type, DeskID As String, NewName As String) As Integer

        Dim Dlist() As DataRow
        Dim dtRow As DataRow
        Dlist = lab.DeskList.Select("DeskID ='" & DeskID & "'")
        If Dlist.Count = 0 Then Return False

        dtRow = Dlist(0)

        dtRow("Name") = NewName
        Desk_Save_Info(dtRow) 'ファイル書き込み

    End Function
    Private Sub Desk_Set_Color(ByRef lab As Lab_Type, DeskID As String, tmpC As String)

        Dim Dlist() As DataRow
        Dim dtRow As DataRow
        Dlist = lab.DeskList.Select("DeskID ='" & DeskID & "'")
        If Dlist.Count = 0 Then Return
        dtRow = Dlist(0)
        dtRow("Color") = tmpC
        Desk_Save_Info(dtRow) 'ファイル書き込み
        lab.DesKTop(lab.DeskTopID).BackColor = Color.FromName(tmpC)

    End Sub
    Private Sub DeskList_Update(ByRef lab As Lab_Type, DeskID As String, Key() As String, Dat() As String)

        Dim Dlist() As DataRow
        Dim dtRow As DataRow
        Dlist = lab.DeskList.Select("DeskID ='" & DeskID & "'")
        If Dlist.Count = 0 Then Return
        dtRow = Dlist(0)

        For L = 0 To Key.Count - 1
            If Key(L) <> "" Then dtRow(Key(L)) = Dat(L)
        Next
        Desk_Save_Info(dtRow) 'ファイル書き込み


    End Sub

    Private Function Image_Init(ByRef TmpD As String) As Image_List_Type
        'ICON類の収集
        '----------------------------------------
        Log_Show(1, "ICON類の収集")


        Dim Image0 As System.Drawing.Image
        Dim Key As String
        If System.IO.Directory.Exists(TmpD) = False Then
            Log_Show(5, "イメージリソースが見つかりません ")
            Return Nothing
        End If



        Dim Flist() As String
        ReDim Flist(0)
        Dim Cnt As Integer
        File_Set_byDir(TmpD, ".ico", Flist)
        File_Set_byDir(TmpD, ".gif", Flist)
        File_Set_byDir(TmpD, ".png", Flist)
        File_Set_byDir(TmpD, ".jpg", Flist)
        File_Set_byDir(TmpD, ".jpeg", Flist)
        With gImg_List
            .Int(Flist.Length)
            .Image.ImageSize = New Size(32, 32)
            For L = 0 To Flist.Length - 1
                If Flist(L) <> "" Then
                    Key = IO.Path.GetFileNameWithoutExtension(Flist(L))

                    Image0 = Image.FromFile(Flist(L))
                    .Image.Images.Add(Key, Image0)

                    ReDim Preserve .Key(.Image.Images.Count - 1)
                    .Key(.Image.Images.Count - 1) = Key
                End If
            Next
        End With
        Return gImg_List

        ' Button20.Image = gImg_List.Image.Images(19)  
    End Function


    Private Function Image_Get_ID(ByRef Img As Image_List_Type, ByRef Key As String) As Integer
        'ICONのID取得
        '----------------------------------------
        Dim tmpR As Integer
        tmpR = 0
        With Img
            For L = 0 To .Image.Images.Count - 1
                If Key = .Key(L) Then
                    Return L
                End If
            Next
        End With

        Return 0
    End Function





    Private Sub Element_Act_Item(ByRef Lab As Lab_Type, DID As Long)
        Dim Ds As DataSet
        Dim Cnd As String

        Cnd = "ID = " & DID.ToString
        Ds = DBAcs_Get_DataS(Lab, "Element", "Type,path", Cnd)

        With Ds.Tables(0).Rows(0)
            Select Case .Item(0)
                Case "File"
                    If System.IO.File.Exists(.Item(1)) Then
                        System.Diagnostics.Process.Start(.Item(1))
                        Log_Show(1, "OK")
                    Else

                        Log_Show(1, "ファイルが消えてる")
                    End If

                Case "Text"
                Case "Url"
                    If InStr(.Item(1), "http") > 0 Then
                        System.Diagnostics.Process.Start(.Item(1))
                    Else
                        System.Diagnostics.Process.Start("https://" & .Item(1))
                    End If
            End Select

        End With





    End Sub
    Private Sub Box_MakeFromElement(ByRef Lab As Lab_Type, Type As String, Cp As Point, ELM As Element_Type)
        'ポトンされたデペンドをスタック化
        '----------------------------
        Select Case Type
            Case "FILE"
                Select Case ELM.Type.ToLower
                    Case "ico", "bmp", "jpg", "gif", "png", "exig", "tiff"
                    Case "text", "bmp", "jpg", "gif", "png", "exig", "tiff"
                    Case "ico", "bmp", "jpg", "gif", "png", "exig", "tiff"
                End Select


                'Box_Makes_OnDESK(gLab, "Viewer", Cp.X, Cp.Y, 1)


            Case Else
        End Select


        'スタック作成
        '   Box_Makes_OnDESK(gLab, Source(0).Tag, cp.X, cp.Y, 1)

        'Element(cnt).Path = Drops(L)
        'Element(cnt).Name = System.IO.Path.GetFileNameWithoutExtension(Drops(L))
        'Element(cnt).Type = "File"


    End Sub

    Private Sub Element_Add_Data(ByRef Lab As Lab_Type, RID As Long, ByRef ELM() As Element_Type)
        'ポトンされたデペンドを登録 (同じものでも別登録になる/サボリ後で考える)
        '----------------------------
        If RID = 0 Then Return
        Dim tmpD As String
        Dim DID As Long
        Dim Cnd As String
        Dim Key As String
        'DB登録
        Cnd = "ID In ( "
        For L = 0 To ELM.Length - 1
            tmpD = ""
            If ELM(L).Name <> "" Then

                With ELM(L)
                    tmpD = gProjID.ToString & ","
                    tmpD &= "'" & .Type & "',"
                    tmpD &= "'" & .Name & "',"
                    tmpD &= "'" & .Path & "',"
                    tmpD &= "1"
                End With
                With Lab.ElementAT
                    Key = Mid(Lab.DBACS_Field("Element"), Len("ID,") + 1)
                    'DID = DBAcs_Insert(Lab, "Element", Key, tmpD) 'ElementDB登録
                End With
                Cnd = Cnd & DID.ToString & ","
                ELM(L).ID = DID
            End If

        Next


        'ElementDBから読み出して、Listviewに登録
        Cnd = "ProjID= " & Str(Lab.ProjID)

        Dim Ds As DataSet
        With Lab.ElementAT
            Ds = DBAcs_Get_DataS(Lab, "Element", .View.Columns.DBKeys, Cnd)
            Element_Show_List(Lab, Ds, True) '表示
        End With



    End Sub


    Private Sub Button25_DragEnter(sender As Object, e As System.Windows.Forms.DragEventArgs)

        'ドラッグを受け付けます。
        If e.Data.GetDataPresent(DataFormats.FileDrop) = True Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If


    End Sub
    Private Sub Button25_DragDrop(sender As Object, e As System.Windows.Forms.DragEventArgs)
        Dim DropS As String() = CType(e.Data.GetData(DataFormats.FileDrop, False), String())

        'Proj_Action_Files("NOT", DropS)


    End Sub




    Private Sub Button21_DragEnter(sender As Object, e As System.Windows.Forms.DragEventArgs)

        'ドラッグを受け付けます。
        If e.Data.GetDataPresent(DataFormats.FileDrop) = True Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If


    End Sub

    Private Sub Button21_DragDrop(sender As Object, e As System.Windows.Forms.DragEventArgs)
        Dim DropS As String() = CType(e.Data.GetData(DataFormats.FileDrop, False), String())

        'Proj_Action_Files("FIXED", DropS)

    End Sub



    Private Sub Button25_Click(sender As System.Object, e As System.EventArgs)
        'Histroy_Go(gLab, 1)
    End Sub









    Private Sub Columns_Set_Listview(ByRef Listview As ListView, ByRef columns As Columns_Type)
        'LISTVIEWのカラム指定
        '-----------------------
        With Listview
            If .Items.Count > 0 Then
                .Clear()
            End If
        End With

        With columns
            For L = 0 To .Name.Length - 1
                Listview.Columns.Add(.Name(L), Integer.Parse(.Width(L)), HorizontalAlignment.Left)
            Next

        End With

    End Sub
    Private Sub Columns_Set_Env(ByRef columns As Columns_Type, ClmKey As String)
        'カラム指定
        '-----------------------
        Dim Name As String
        Dim Wid As String
        Dim DBkeys As String

        Select Case ClmKey

            Case "[Proj]"
                DBkeys = "Name,Owner,ID"
                Name = "Name,Owner,ID"
                Wid = "300,80,5"
            Case "[Element]"
                DBkeys = "Name,Type,Path,ID"
                Name = "Name,Type,Path,ID"
                Wid = "300,30,80,10"
            Case "[Stack]"
                DBkeys = "Name,Type,ID"
                Name = "Name,Type,ID"
                Wid = "300,30,10"
        End Select
        With columns
            .Type = ClmKey
            .DBKeys = DBkeys
            .Name = Split(Name, ",")
            .Width = Split(Wid, ",")

        End With
    End Sub









    Private Sub Proj_Get_ProjList(ByRef lab As Lab_Type)
        Dim Ent() As String
        Ent = Split("ProjID|Title|OwnerID|ProjKey|Path|State|Summary|Hypothesis|Purpose|Problem", "|")
        With lab
            .ProjList = New DataTable("Project")
            For L = 0 To Ent.Count - 1
                If Ent(L) <> "" Then .ProjList.Columns.Add(Ent(L))
            Next
        End With

        Dim ProjPath As String
        ProjPath = gApp_Path & "\ENV\PRJ\"
        Dim Owners() As String
        Dim ProFile() As String
        Owners = System.IO.Directory.GetDirectories(ProjPath, "*")

        For L = 0 To Owners.Count - 1
            ProFile = System.IO.Directory.GetDirectories(Owners(L), "*")
            For K = 0 To ProFile.Count - 1
                Proj_Add_ProjListByFile(lab, ProFile(K) & "\Proj.txt")
            Next
        Next


    End Sub

    Private Function DeskList_Set(ByRef Path As String) As DataTable
        Dim DeskList As DataTable

        Dim Ent() As String
        Ent = Split("Name,DeskID,DeskTopID,ProjKey,OwnerID,State,Type,Color,ShiftX,ShiftY,ZM,Path", ",")
        DeskList = New DataTable("Desk")

        For L = 0 To Ent.Count - 1
            If Ent(L) <> "" Then DeskList.Columns.Add(Ent(L))
        Next

        Dim Desks() As String
        Desks = System.IO.Directory.GetDirectories(Path & "\", "*")
        For L = 0 To Desks.Count - 1
            If InStr(Desks(L), "git") = 0 Then
                DeskList_Add_ByFile(DeskList, Desks(L) & "\Desk.txt")
            End If

        Next

        Return DeskList
    End Function


    Private Sub DeskList_Add_ByFile(ByRef DeskList As DataTable, File As String)
        Try

            Dim Pa() As String
            Pa = Split(File, "\")

            Dim TmpD As String
            TmpD = IO.File.ReadAllText(File)

            Dim TmpR() As String = Split(TmpD, vbCrLf)
            Dim key() As String
            Dim dtRow As DataRow
            With DeskList
                dtRow = .NewRow

                For L = 0 To TmpR.Count - 1
                    key = Split(TmpR(L), vbTab)
                    If Trim(key(0)) <> "" Then
                        dtRow(key(0)) = key(1)
                    End If

                Next
                dtRow("Path") = File
                dtRow("ProjKey") = Pa(Pa.Count - 4).ToString & "\" & Pa(Pa.Count - 3).ToString
                dtRow("DeskID") = Pa(Pa.Count - 2).ToString

                .Rows.Add(dtRow)
            End With

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

    Private Sub Box_Get_BoxList(ByRef BoxList As DataTable, ByRef Path As String)
        If BoxList Is Nothing Then
            Dim Ent() As String
            Ent = Split("BoxID,DeskID,ProjKey,OwnerID,State,Type,ParentID,LineID,LineNo,BluePrint,Path", ",")
            BoxList = New DataTable("Box")

            For L = 0 To Ent.Count - 1
                If Ent(L) <> "" Then BoxList.Columns.Add(Ent(L))
            Next
        End If

        Dim Boxs() As String
        Boxs = System.IO.Directory.GetDirectories(Path & "\", "*")
        For L = 0 To Boxs.Count - 1
            Box_Add_BoxListByFile(BoxList, Boxs(L) & "\Box.txt")
        Next

    End Sub

    Private Sub Box_Add_BoxListByFile(ByRef DeskList As DataTable, File As String)

        Dim Pa() As String
        Pa = Split(File, "\")

        If System.IO.File.Exists(File) = False Then Return
        Dim TmpD As String
        TmpD = IO.File.ReadAllText(File)

        Dim TmpR() As String = Split(TmpD, vbCrLf)
        Dim key() As String
        Dim dtRow As DataRow
        With DeskList
            dtRow = .NewRow

            For L = 0 To TmpR.Count - 1
                key = Split(TmpR(L), vbTab)
                If Trim(key(0)) <> "" Then
                    dtRow(key(0)) = key(1)
                End If

            Next

            dtRow("Path") = File
            dtRow("ProjKey") = Pa(Pa.Count - 5) & "\" & Pa(Pa.Count - 4)
            dtRow("DeskID") = Pa(Pa.Count - 3)
            dtRow("BoxID") = Pa(Pa.Count - 2)
            'MsgBox(Pa(Pa.Count - 1))
            .Rows.Add(dtRow)
        End With
    End Sub



    Private Sub Proj_Add_ProjListByFile(ByRef lab As Lab_Type, File As String)

        Dim TmpD As String
        TmpD = IO.File.ReadAllText(File)
        Dim TmpR() As String = Split(TmpD, vbCrLf)
        Dim key() As String
        Dim dtRow As DataRow
        With lab.ProjList
            dtRow = .NewRow

            For L = 0 To TmpR.Count - 1
                key = Split(TmpR(L), vbTab)
                If Trim(key(0)) <> "" Then
                    dtRow(key(0)) = key(1)
                End If

            Next
            dtRow("Path") = File
            dtRow("ProjKey") = dtRow("OwnerID") & "\" & dtRow("ProjID")
            .Rows.Add(dtRow)
        End With
    End Sub
    Private Sub Proj_Show_List(ByRef Plist As DataRow(), clearFlg As Boolean)
        'LISTVIEWにProjデータを表示
        '-----------------------------
        Try
            Dim Col(), w() As String
            Col = Split("Title,OwnerID,ProjID,Path", ",")
            w = Split("300,20,20,20", ",")

            Dim cnt As Integer
            cnt = 1

            Dim ITM As ListViewItem
            With ListView1
                If clearFlg = True Then .Clear()
                For L = 0 To Col.Count - 1
                    .Columns.Add(Col(L), Integer.Parse(w(L)), HorizontalAlignment.Left)
                Next
                For Each row As DataRow In Plist
                    ITM = .Items.Add(row("Title"))
                    ITM.Tag = row("ProjKey") 'クリック時のタグがプロジェクトのパス
                    ITM.SubItems.Add(row("OwnerID"))
                    ITM.SubItems.Add(row("ProjID"))
                    ITM.SubItems.Add(row("Path"))
                    cnt += 1
                Next
                .Visible = True
            End With

            Return
        Catch ex As Exception
            Log_Show(1, ex.Message)
        End Try

    End Sub

    Private Sub Proj_Add_List(ByRef dtRow As DataRow)
        'LISTVIEWにProjデータを表示
        '-----------------------------
        Try
            'Dim Col(), w() As String
            'Col = Split("Title,OwnerID,ProjID,Path", ",")

            Dim ITM As ListViewItem
            With ListView1

                ITM = .Items.Add(dtRow("Title"))

                ITM.SubItems.Add(dtRow("OwnerID"))
                ITM.SubItems.Add(dtRow("ProjID"))
                ITM.SubItems.Add(dtRow("Path"))
                ITM.Tag = dtRow("ProjKey") 'クリック時のタグがプロジェクトのパス
                .Visible = True
            End With
            ITM.Selected = True
            Return
        Catch ex As Exception
            Log_Show(1, ex.Message)
        End Try

    End Sub
    Private Sub Element_Show_List(ByRef Lab As Lab_Type, ByRef Ds As DataSet, clearFlg As Boolean)
        'ElementデータをLISTVIEWに表示
        '-----------------------------
        Dim Itm As ListViewItem
        Dim Img As Image

        With Lab.ElementAT.View
            If Ds.Tables(0).Rows.Count = 0 Then
                .ListView.Clear()
                Return
            End If

            If clearFlg = True Then
                Columns_Set_Listview(.ListView, .Columns)
            End If

            For L = 0 To Ds.Tables(0).Rows.Count - 1
                Itm = .ListView.Items.Add(Ds.Tables(0).Rows(L).Item(0))
                Itm.ImageIndex = Element_Get_ImageID(Lab, Ds.Tables(0).Rows(L).Item(2))

                Itm.Tag = Val(Ds.Tables(0).Rows(L).Item(Ds.Tables(0).Columns.Count - 1)) '最後のアイテムをDIDにしてるので
                For K = 1 To Ds.Tables(0).Columns.Count - 1
                    .ListView.Items(L).SubItems.Add(Ds.Tables(0).Rows(L).Item(K))
                Next
            Next

            'If Ds.Tables(0).Rows.Count > 0 Then
            '    .View.Items(0).Selected = True
            'End If
            .ListView.Refresh()
        End With
    End Sub

    Private Function Element_Get_ImageID(ByRef Lab As Lab_Type, Key As String) As Integer

        Dim tmpR As Integer
        Dim Ext As String
        tmpR = Image_Get_ID(Lab.Image_List, Key)
        If tmpR > 0 Then Return tmpR

        If InStr(Key, "http") > 0 Then
            tmpR = Image_Get_ID(Lab.Image_List, "http")
            If tmpR > 0 Then Return tmpR
        End If

        If System.IO.Directory.Exists(Key) Then
            tmpR = Image_Get_ID(Lab.Image_List, "Folder")
            If tmpR > 0 Then Return tmpR
        End If
        Try
            Ext = Path.GetExtension(Key)
            If Ext <> "" Then
                tmpR = Image_Get_ID(Lab.Image_List, Ext)
                If tmpR > 0 Then Return tmpR

                tmpR = Image_Add_DatabyPath(Lab.Image_List, Key, Ext)
                If tmpR > 0 Then Return tmpR

            End If
        Catch

        End Try




        Return 1

    End Function

    Private Function Image_Add_DatabyPath(ByRef Image_List As Image_List_Type, path As String, Key As String) As Integer
        'ファイルに関連したアイコンをイメージリストに設定
        Dim tmpIcon As Icon
        Dim max As Integer
        Try
            tmpIcon = Icon.ExtractAssociatedIcon(path)
        Catch
            Return 0
        End Try
        If tmpIcon Is Nothing Then Return 0

        With Image_List
            .Image.Images.Add(Key, tmpIcon.ToBitmap)
            max = .Image.Images.Count - 1
            ReDim Preserve .Key(max)
            .Key(max) = Key
            Return max
        End With

    End Function


    Private Function Icon_GetbyPath(Fpath As String) As Icon
        Dim tmpIcon As Icon
        tmpIcon = Icon.ExtractAssociatedIcon(Fpath)
        Return tmpIcon

    End Function


    'Private Sub Proj_ListView_Item_Remove(ByRef ObjCtrl As System.Windows.Forms.ListView)
    '    'Listから指定TAG 0を削除する
    '    '----------------------------------
    '    Dim DX() As Integer

    '    With ObjCtrl
    '        For L = .Items.Count - 1 To 0 Step -1
    '            If .Items(L).Tag = 0 Then
    '                .Items(L).Remove()
    '            End If
    '        Next
    '    End With

    'End Sub



    'Private Sub Proj_Show_Atr(ByRef Lab As Lab_Type, ByRef RID As Long)
    '    'レポジトリのATR表示
    '    '--------------------------
    '    Dim Ds As DataSet
    '    Ds = DBAcs_Get_DataS(Lab, "Proj", "Note,Prof", "ID=" & Str(RID))
    '    If Ds.Tables(0).Rows.Count = 0 Then Return
    '    With Ds.Tables(0)

    '        'Lab.ElementAT.Note.Text = .Rows(0).Item(0)
    '        'Lab.ElementAT.Prof.Text = .Rows(0).Item(1)
    '    End With
    '    Proj_Show_Element(Lab, RID)
    '    'RichTextBox3.Text = Lab.ElementAT.Note.Text


    'End Sub
    Private Sub Proj_Show_Element(ByRef Lab As Lab_Type, ByRef PID As Long)
        'レポジトリ対応のElement表示
        '--------------------------
        Dim Ds As DataSet
        Dim Cnd As String
        '■ちょっと待ってね
        'Ds = DBAcs_Get_DataS(Lab, "Element", "ID", "ProjID=" & Str(PID))
        'If Ds.Tables(0).Rows.Count = 0 Then
        '    Lab.ElementAT.View.ListView.Clear()
        '    '   Columns_Set_Listview(Lab.Element.View, Lab.Element.Columns)
        '    Return
        'End If

        'Cnd = "ID IN ( "
        'With Ds.Tables(0)
        '    For L = 0 To .Rows.Count - 2
        '        Cnd = Cnd & .Rows(L).Item(0) & ","
        '    Next
        '    Cnd = Cnd & .Rows(.Rows.Count - 1).Item(0) & ")"
        'End With

        Cnd = "ProjID= " & Str(PID)
        With Lab.ElementAT
            Ds = DBAcs_Get_DataS(Lab, "Element", .View.Columns.DBKeys, Cnd)
            Element_Show_List(Lab, Ds, True)

        End With

    End Sub




    Private Function Proj_Atr_Add_Item(ByRef LV As System.Windows.Forms.ListView, ByRef DID As Integer, ByRef BType As String, ByRef SType As String, ByRef FX As String, ByRef VX As String, ByRef CX As String, ByRef State As Integer, ByRef Ratio As Double) As ListViewItem
        'ATR ITEMをADD
        '---------------------

        Dim lvi As ListViewItem
        lvi = LV.Items.Add(BType)
        lvi.Tag = DID


        lvi.ImageIndex = Image_Get_ID(gImg_List, BType & "_")
        lvi.SubItems.Add(SType)
        lvi.SubItems.Add(FX)
        lvi.SubItems.Add(VX)
        lvi.SubItems.Add(CX)
        lvi.SubItems.Add(State)
        lvi.SubItems.Add(Ratio)

        Return lvi
    End Function

    Private Sub Proj_Atr_New_Item(ByRef ListV As ListView, ByRef BType As String)
        ''ATR ITEMを新規追加
        '-----------------------
        With ListV
            Proj_Atr_Add_Item(ListView2, 0, BType, "", "", "", "", 0, 0)
            .SelectedItems.Clear()
            .Items(.Items.Count - 1).EnsureVisible()
            .Items(.Items.Count - 1).Selected = True


        End With


    End Sub










    Private Sub Button14_Click(sender As System.Object, e As System.EventArgs)
        'ルールの適用

        'Rule_Apply(gStudy(ComboBox8.SelectedIndex)) 'ルールの適用

    End Sub


    Private Sub Button17_DragEnter(sender As Object, e As System.Windows.Forms.DragEventArgs)
        'ドラッグを受け付けます。
        If e.Data.GetDataPresent(DataFormats.FileDrop) = True Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub




    Private Sub ContextMenuStrip1_ItemClicked(sender As Object, e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ContextMenuStrip1.ItemClicked

        'Menu_Act(gBff_ObjName, "0\" & e.ClickedItem.ToString)

        Select Case e.ClickedItem.Tag
            Case "NEW"
                Proj_Make_New(gLab)
            Case "RENAME"
                If ListView1.SelectedItems.Count = 0 Then Return

                ListView1.SelectedItems(0).BeginEdit()
            Case "DEL"
                Proj_Del(gLab)

            Case Else
        End Select
    End Sub






    Private Sub ListView1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles ListView1.KeyDown


        Dim DeskID As Integer
        DeskID = 1
        With ListView1

            '削除
            If e.KeyData = Keys.Delete Then
                If IsNothing(.SelectedItems) Then Return
                If .SelectedItems.Count = 0 Then Return
                Proj_Del(gLab)

                If Not (.FocusedItem Is Nothing) Then
                    .FocusedItem.Selected = True
                End If

            End If

            '作成はしない
            If e.KeyData = Keys.Insert Then
                'If Proj_Node_Make_Check(.SelectedNode, "NEW") = True Then
                '     'Histroy_start(gLab.Desk(1).history)
                '    Proj_Act_Node_Make(gLab,1.History, .SelectedNode, "NEW", True)
                'End If
            End If

        End With




    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        'プロジェクトの読み込み

        With ListView1
            If .SelectedItems Is Nothing Then Return
            If .SelectedItems.Count <= 0 Then Return
            If .SelectedItems.Count > 1 Then Return

            If .SelectedItems(0).Tag Is Nothing Then Return

            gLab.ProjName = .SelectedItems(0).Text
            gLab.ProjKey = .SelectedItems(0).Tag
            gLab.ProjPath = gENV_Path & "\PRJ\" & gLab.ProjKey

            Dim SS() As String = Split(.SelectedItems(0).Tag, "\")
            gLab.ProjID = SS(1)
            gLab.ProjOwnerID = SS(0)


            '  Me.Text = "Ginger-breadboard       " & gLab.ProjName
        End With
        'Log_Show(1, gProjID & "|" & gLab.ProjName)
        Proj_Set(gLab)


    End Sub
    Private Sub Proj_Set(ByRef Lab As Lab_Type)
        ' '--------------------------
        With Lab
            .DeskList = DeskList_Set(.ProjPath)

        End With

        Desk_ShowByProjID(Lab)
    End Sub

    Private Sub Desk_ShowByProjID(ByRef Lab As Lab_Type)
        'デスクの表示
        '------------------------------
        Dim Dlist As DataRow()
        Dim Cnt As Integer
        Desk_Clear_Name(Lab)
        Desk_Clear_Desk(Lab)
        If Lab.BoxList IsNot Nothing Then
            Lab.BoxList.Clear()
        End If
        With Lab
            Dlist = .DeskList.Select("State = 'Nomal'")
            If Dlist.Count = 0 Then Return
            ReDim .Desk(Dlist.Count)
            .DeskTopID = 0
            For Each row As DataRow In Dlist
                .DeskID = row("DeskID")
                row("DeskTopID") = .DeskTopID
                Desk_Show(Lab, row)
                Desk_Show_Name(Lab, row)

                Box_Get_BoxList(Lab.BoxList, System.IO.Path.GetDirectoryName(row("Path")))
                Box_Set_OnDESK(Lab, .DesKTop(.DeskTopID))


                .DeskTopID += 1
            Next
        End With

        '本来なら以前の作業中のページ
        Lab.DeskTopID = 0
        DataGridView3(0, 0).Selected = False
        DataGridView3(0, 0).Selected = True
    End Sub
    Private Sub Desk_Show(ByRef Lab As Lab_Type, ByRef Row As DataRow)

        '  Dim DID As Integer
        With Lab

            ReDim Preserve .DesKTop(.DeskTopID)
            .DesKTop(.DeskTopID) = New Panel
            SplitContainer1.Panel1.Controls.Add(.DesKTop(.DeskTopID))

            With .DesKTop(.DeskTopID)
                .Width = Lab.WorldW  '4000
                .Height = Lab.WorldH '4000

                .Left = Lab.WorldX + Row("ShiftX")
                .Top = Lab.WorldY + Row("ShiftY")
                .BackColor = Color.FromName(Row("Color"))
                .Tag = Row("DeskID")
                .Name = Lab.DeskTopID.ToString
                .AllowDrop = True
                .AllowDrop = True
                .Visible = True

                AddHandler .DragOver, AddressOf Desk_DragOver
                AddHandler .DragDrop, AddressOf Desk_DragDrop
                AddHandler .MouseDown, AddressOf Desk_MouseDown
                AddHandler .MouseMove, AddressOf Desk_MouseMove
                AddHandler .MouseUp, AddressOf Desk_MouseUp
                AddHandler .MouseDoubleClick, AddressOf Desk_DoubleClick
                AddHandler .SizeChanged, AddressOf Desk_ReSize
            End With

        End With

    End Sub


    'Private Sub Box_ShowByDesk(ByRef Lab As Lab_Type)
    '    ''BOXの表示
    '    Dim Dlist() As DataRow
    '    Dim Blist() As DataRow
    '    Dim dtRow As DataRow
    '    With Lab
    '        Dlist = .DeskList.Select("State = 'Nomal' AND DeskID = '" & .DeskID.ToString & "'")
    '        If Dlist.Count = 0 Then Return
    '        dtRow = Dlist(0)

    '        Blist = .BoxList.Select("State = 'Nomal' AND DeskID = '" & .DeskID.ToString & "'")

    '    End With


    '    Box_Set_OnDESK(.BoxList, .DTop)


    '    ''End With
    '    '  Dim DID As Integer
    '    'With Lab

    '    '    ReDim Preserve .DesKTop(.DeskTopID)
    '    '    .DesKTop(.DeskTopID) = New Panel
    '    '    SplitContainer1.Panel1.Controls.Add(.DesKTop(.DeskTopID))

    '    '    With .DesKTop(.DeskTopID)
    '    '        .Width = Lab.WorldW  '4000
    '    '        .Height = Lab.WorldH '4000

    '    '        .Left = Lab.WorldX + Row("ShiftX")
    '    '        .Top = Lab.WorldY + Row("ShiftY")
    '    '        .BackColor = Color.FromName(Row("Color"))
    '    '        .Tag = Row("DeskID")
    '    '        .Name = Lab.DeskTopID.ToString
    '    '        .AllowDrop = True
    '    '        .AllowDrop = True
    '    '        .Visible = True

    '    '        AddHandler .DragOver, AddressOf Desk_DragOver
    '    '        AddHandler .DragDrop, AddressOf Desk_DragDrop
    '    '        AddHandler .MouseDown, AddressOf Desk_MouseDown
    '    '        AddHandler .MouseMove, AddressOf Desk_MouseMove
    '    '        AddHandler .MouseUp, AddressOf Desk_MouseUp
    '    '        AddHandler .MouseDoubleClick, AddressOf Desk_DoubleClick
    '    '        AddHandler .SizeChanged, AddressOf Desk_ReSize
    '    '    End With

    '    'End With

    'End Sub


    Private Sub Bx_Re_Draw(ByRef Lab As Lab_Type, ByRef Drbox As AxMSINKAUTLib.AxInkPicture)
        ''イメージデータを保存

        ''Dim Bs() As Byte
        ''Bs = Drbox.Ink.Save(1)
        ''DBACS_Update_Byte(Lab, "Draw", "Text", "ID =" & Drbox.name, Bs)

        ''BASE64
        'Dim Bs As String
        'Bs = Drbox.Ink.Save(1)
        'DBAcs_Update(Lab, "Draw", "Text = '" & Drbox.Name & "'", "ID =" & Bs)


    End Sub
    Private Sub Bx_Re_Text(ByRef Lab As Lab_Type, TboxID As Long, ByRef Tdata As String)
        'テキストを保存
        ' DBAcs_Update(Lab, "Tbox", "Text = '" & Tdata & "'", "ID = " & TboxID.ToString)
    End Sub

    Private Sub Bx_Re_Pict(ByRef Lab As Lab_Type, TAG As String, path As String)
        'イメージデータのパスを保存
        '  DBAcs_Update(Lab, "Pbox", "PICT = '" & path & "'", "ID = " & TAG)


    End Sub

    Private Sub Bx_Re_Pict2(ByRef Lab As Lab_Type, ByRef PID As String, ByRef File As String)
        'イメージデータを保存

        ' ファイルを読み込む
        Dim Fs As New System.IO.FileStream(File, System.IO.FileMode.Open, System.IO.FileAccess.Read)
        'ファイルを読み込むバイト型配列を作成する
        Dim Bs(Fs.Length - 1) As Byte
        Fs.Read(Bs, 0, Bs.Length)
        Fs.Close()

        'Dim BData() As Byte イメージデータのコンバータ
        'BData = ImageToByteArray(bs)

        DBACS_Update_Byte(Lab, "BDat", "Dat", "ID =" & PID, Bs)


    End Sub

    Private Sub DBACS_Update_Byte(ByRef Lab As Lab_Type, Key As String, Act As String, Cnd As String, ByRef BData() As Byte)
        'バイト型配列の保存
        '-----------------------
        'Dim strSQL As String
        'strSQL = ""
        'strSQL &= "UPDATE TBL_" & Key & " SET "
        'strSQL &= " Dat = @pic"
        'strSQL &= " WHERE " & Cnd


        ''    登録処理
        'Select Case Lab.Db_Type
        '    Case "SQLite"
        '    Case "MSSQL"
        '        Using con As New SqlClient.SqlConnection(Lab.DbPath & ";Initial Catalog =" & Lab.Db_Type & Key)
        '            Dim cmd As New SqlClient.SqlCommand(strSQL, con)
        '            cmd.Parameters.Add("@pic", SqlDbType.Binary, BData.Length).Value = BData
        '            con.Open()
        '            For I As Integer = 1 To 50
        '                Try
        '                    cmd.ExecuteNonQuery()
        '                    Exit For
        '                Catch ex As Exception
        '                    Log_Show(1, "DBエラー" & I & ex.Message)
        '                    For J As Integer = 1 To 100
        '                        System.Windows.Forms.Application.DoEvents()
        '                    Next
        '                End Try
        '            Next
        '            con.Close()
        '        End Using
        'End Select



    End Sub

    'Private Sub DBACS_Update_ByteMSSQL(ByRef Lab As Lab_Type, Key As String, Act As String, Cnd As String, ByRef BData() As Byte)
    '    'バイト型配列の保存
    '    '-----------------------
    '    Dim strSQL As String
    '    strSQL = ""
    '    strSQL &= "UPDATE TBL_" & Key & " SET "
    '    strSQL &= " Dat = @pic"
    '    strSQL &= " WHERE " & Cnd


    '    '    登録処理
    '    Using con As New SqlClient.SqlConnection(Lab.DbPath & ";Initial Catalog =" & Lab.Db_Type & Key)
    '        Dim cmd As New SqlClient.SqlCommand(strSQL, con)
    '        cmd.Parameters.Add("@pic", SqlDbType.Binary, BData.Length).Value = BData
    '        con.Open()
    '        For I As Integer = 1 To 50
    '            Try
    '                cmd.ExecuteNonQuery()
    '                Exit For
    '            Catch ex As Exception
    '                Log_Show(1, "DBエラー" & I & ex.Message)
    '                For J As Integer = 1 To 100
    '                    System.Windows.Forms.Application.DoEvents()
    '                Next
    '            End Try
    '        Next
    '        con.Close()
    '    End Using
    'End Sub





    Private Sub Bx_Get_Draw(ByRef Lab As Lab_Type, ByRef Draw As AxMSINKAUTLib.AxInkPicture, DrawID As Long)
        Dim imgconv As New ImageConverter()
        Dim DS As DataSet
        DS = DBAcs_Get_DataS(Lab, "Draw", "Text", "ID = " & DrawID.ToString)
        If DS.Tables(0).Rows.Count = 0 Then Return

        ' Draw.Ink.Load(CType(imgconv.ConvertFrom(DS.Tables(0).Rows(0).Item(0)), Image))
        ' Draw.Ink.Load(DS.Tables(0).Rows(0).Item(0), isf))
        'ConvertFrom(Context As ITypeDescriptorContext, culture As CultureInfo, value As Object)
        'With Draw
        '    .Ink.Dispose()
        '    .Ink = New Microsoft.Ink.Ink()
        '    .Ink.Load(DS.Tables(0).Rows(0).Item(0))
        '    .InkEnabled = True
        '    .Refresh()
        'End With

        ''If DS.Tables(0).Rows(0).Item(0) = Nothing Then Return
        'Dim bs() As Byte = System.Convert.FromBase64String(DS.Tables(0).Rows(0).Item(0))
        'Draw.Strokes = New StrokeCollection(MS)
        '' Draw.Ink.Strokes.Cast(Of Base64FormattingOptions) = a
    End Sub


    Private Sub Bx_Get_Pict(ByRef Lab As Lab_Type, ByRef Pict As PictureBox, PictID As Long)
        'イメージデータをパスから読み出し
        Dim DS As DataSet
        DS = DBAcs_Get_DataS(Lab, "Pbox", "Dat", "ID = " & PictID.ToString)
        If DS.Tables(0).Rows.Count = 0 Then Return
        Pict.ImageLocation = DS.Tables(0).Rows(0).Item(0)

    End Sub


    Private Sub Bx_Get_Pict2(ByRef Lab As Lab_Type, ByRef Pict As PictureBox, PictID As Long)
        Dim imgconv As New ImageConverter()
        Dim DS As DataSet
        DS = DBAcs_Get_DataS(Lab, "BDat", "Dat", "ID = " & PictID.ToString)
        If DS.Tables(0).Rows.Count = 0 Then Return
        If DS.Tables(0).Rows(0).Item(0) Is DBNull.Value Then Return
        'If DS.Tables(0).Row.isnull("Dat") Then Return
        ' Dim img As System.Drawing.Image = System.Drawing.Image.FromFile(files(0))

        ' Pict.Image = ByteArrayToImage(DS.Tables(0).Rows(0).Item(0))

        Pict.Image = CType(imgconv.ConvertFrom(DS.Tables(0).Rows(0).Item(0)), Image)
    End Sub
    Private Function Bx_get_ByteDATA(ByRef Lab As Lab_Type, DbKey As String, Act As String, cnd As String) As Byte()
        'イメージデータをDBから読み出し
        'Dim DS As DataSet
        ''Dim bData As Byte()
        ''DS = DBAcs_Get_DataS(Lab, "Pbox", "Dat", "ID = " & PictID.ToString)

        ''PICT_GET(Lab.DbPath, "Pbox", "Dat", "ID = " & PictID.ToString)



        'Dim strSQL As String
        'strSQL = " SELECT "" Pict"
        'strSQL &= vbCrLf & " FROM  TBL_Pbox2"
        'strSQL &= vbCrLf & " WHERE  ID =" & PictID


        '''    読み出し処理
        'Using con As New SqlClient.SqlConnection(Lab.DbPath & ";Initial Catalog =" & "DB_Pbox2")
        '    Dim cmd As New SqlClient.SqlCommand(strSQL, con)
        '    Dim objRs As SqlDataReader

        '    con.Open()


        '    'objRs = cmd.ExecuteReader()
        '    'bData = CType(objRs.Item(0), Byte())
        '    ''' Dim imageData As Byte() = DirectCast(cmd.ExecuteScalar(), Byte())
        '    ''Dim imagedata() As Byte = CType(cmd.ExecuteScalar(), Byte())
        '    ''' Pict.Image = ByteArrayToImage(Reader)
        '    For I As Integer = 1 To 50
        '        Try
        '            Dim imageData As Image = cmd.ExecuteScalar()


        '            Exit For
        '        Catch ex As Exception
        '            Log_Show(1, "DBエラー" & I & ex.Message)
        '            For J As Integer = 1 To 100
        '                System.Windows.Forms.Application.DoEvents()
        '            Next
        '        End Try
        '    Next
        '    con.Close()
        'End Using



    End Function

    'Private Function PICT_GET(ByRef DB_Path As String, Name As String, Action As String, ByRef Condition As String) As DataSet

    '    'DB 指定Key の　値　を格納
    '    '--------------------------
    '    Dim tmpR As Integer

    '    Dim DBCon As SqlConnection
    '    Dim TmpCmd As String
    '    tmpR = -1

    '    TmpCmd = "Select " & Action & " FROM [" & .dbname & Name & "].[dbo].[" & "TBL_" & Name & "] WHERE " & Condition
    '    ' TmpCmd = "Select * FROM [" & .DB_Name & "].[dbo].[" & .TBL_Name & "] "
    '    DBCon = New SqlConnection(DB_Path & ";Initial Catalog =" & .dbname & Name)
    '    DBCon.Open()
    '    Dim cmd As New SqlClient.SqlCommand(TmpCmd, DBCon)

    '    Dim objRs As SqlDataReader
    '    objRs = cmd.ExecuteReader()
    '    Dim bData As Byte()
    '    bData = CType(objRs.Item(0), Byte())
    'End Function


    Public Shared Function ByteArrayToImage(ByVal b As Byte()) As Image
        ' バイト配列をImageオブジェクトに変換
        Dim imgconv As New ImageConverter()
        Dim img As Image = CType(imgconv.ConvertFrom(b), Image)
        Return img
    End Function


    Public Shared Function ImageToByteArray(ByVal img As Image) As Byte()
        ' Imageオブジェクトをバイト配列に変換
        Dim imgconv As New ImageConverter()
        Dim b As Byte() = CType(imgconv.ConvertTo(img, GetType(Byte())), Byte())
        Return b
    End Function

    Private Sub ListView2_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs)

        With ListView2
            '            Dim Target_Item As ListViewItem
            '削除
            If e.KeyData = Keys.Delete Then
                ListView_Del(ListView2)　'
            End If

            '作成
            If e.KeyData = Keys.Insert Then
                Proj_Atr_New_Item(ListView2, "S")
            End If

        End With

    End Sub


    Private Sub ListView_Del(ByRef LV As ListView)
        'projectの削除
        '-----------------

        With LV
            If IsNothing(.SelectedItems) Then Return
            If .SelectedItems.Count = 0 Then Return

            '  'Histroy_start(gLab.Desk(1).history)
            For L = .SelectedItems.Count - 1 To 0 Step -1
                .SelectedItems(L).Remove()
            Next

            If Not (.FocusedItem Is Nothing) Then
                .FocusedItem.Selected = True
            End If


        End With

    End Sub



    Private Sub ListView2_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs)


        'Select Case e.Button
        '    Case Windows.Forms.MouseButtons.Right
        '        gBff_Act_Tree = Nothing
        '        gBff_Act_List = ListView2
        'End Select




    End Sub







    Private Sub ContextMenuStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

        'Dim menu As ContextMenuStrip = DirectCast(sender, ContextMenuStrip)
        ''ContextMenuStripを表示しているコントロールを取得する
        'Dim source As Control = menu.SourceControl
        'If source IsNot Nothing Then
        '    gBff_ObjName = source.Name
        'End If


    End Sub

    Private Sub ContextMenuStrip2_Opening(sender As System.Object, e As System.ComponentModel.CancelEventArgs)

    End Sub
    Private Sub LineV_MouseEnter(sender As Object, e As System.EventArgs)
        Dim Panel As Panel = CType(sender, Panel)

        Panel.BorderStyle = BorderStyle.FixedSingle
        Panel.BringToFront()

        Panel.Width = 12

        Dim Blist() As DataRow
        Dim dtRow As DataRow
        With gLab
            .LineID = Val(sender.name)
            gLab.Line = sender
            .Base = sender
            .Flg_BoxMouse = "OnV"
            Blist = .BoxList.Select("State = 'Nomal' AND BoxID = '" & .LineID.ToString & "' AND DeskID = '" & .DeskID.ToString & "'")
            If Blist.Count = 0 Then Return '異常
            dtRow = Blist(0)
            .LineNo = Val(dtRow("LineNo"))


            Dim BoxS() As Panel
            Dim cnt As Integer
            Dim Tag As String = .Line.Name
            For Each Box As Panel In .DesKTop(.DeskTopID).Controls
                If Line_Check_Parent(Box.Tag, Tag) = True Then
                    ReDim Preserve BoxS(cnt)
                    BoxS(cnt) = Box
                    cnt += 1
                End If
            Next
            .LineMax = cnt + 1
        End With

    End Sub
    Private Sub LineV_MouseLeave(sender As Object, e As System.EventArgs)

        Dim Panel As Panel = CType(sender, Panel)
        gLab.LineID = 0
        Panel.BorderStyle = BorderStyle.None

        Panel.Width = 4
        gLab.Flg_BoxMouse = ""
    End Sub
    Private Sub LineV_Refresh(ByRef lab As Lab_Type, CloseID As Integer)
        Dim sl() As Integer
        Dim BoxS() As Panel
        With lab

            Dim cnt As Integer

            Dim Tag As String = .Line.Name

            For Each Box As Panel In .DesKTop(.DeskTopID).Controls
                If Line_Check_Parent(Box.Tag, Tag) = True Then
                    ReDim Preserve BoxS(cnt)
                    BoxS(cnt) = Box
                    ReDim Preserve sl(cnt)
                    sl(cnt) = Box.Top
                    cnt += 1
                End If
            Next
            If cnt = 0 Then Return
            Array.Sort(sl, BoxS)

            Dim LW As Integer
            Dim Dp As Integer
            Dp = lab.Line.Top + lab.Line.Height
            LW = BoxS(cnt - 1).Height
            BoxS(cnt - 1).Left = lab.Line.Left + lab.Line.Width + 4
            BoxS(cnt - 1).Top = Dp - BoxS(cnt - 1).Height '- 4


            '  Box_Re_Move(gLab, Val(BoxS(sl.GetByIndex(sl.Count - 1)).Name), BoxS(sl.GetByIndex(sl.Count - 1)))
            Dim Gx As Integer
            Gx = 0
            For L = cnt - 2 To 0 Step -1
                BoxS(L).Left = lab.Line.Left + lab.Line.Width + 4

                If L < CloseID Then
                    BoxS(L).Top = BoxS(L + 1).Top - Gx
                    If LW < BoxS(L).Height + L Then
                        LW = BoxS(L).Height + L
                    End If
                    Gx += 1
                Else
                    BoxS(L).Top = BoxS(L + 1).Top - BoxS(L).Height '- 1
                    LW += BoxS(L).Height + 1

                End If
                If InStr(BoxS(L).Tag, "Line") > 0 Then
                    Line_Re_Move(gLab, Val(BoxS(L).Name), BoxS(L).Left, BoxS(L).Top)
                Else
                    Box_Re_Move(gLab, Val(BoxS(L).Name), BoxS(L))
                End If

                BoxS(L).BringToFront()
            Next
            .Line.Height = LW
            .Line.Top = Dp - LW
            Line_Re_Move(lab, .LineID, .Line.Left, .Line.Top)
            Line_Re_Size(gLab, Val(lab.Line.Name), lab.Line.Width, LW)
        End With


    End Sub


    Private Sub LineV_MouseUp(sender As Object, e As MouseEventArgs)
        '左クリックの場合

        If e.Button = MouseButtons.Left Then

            'コントロール取得
            Dim control As Control = CType(sender, Control)

            Select Case gFlag_Pull
                Case "Move" '移動

                    control.Left = control.Left + e.X - gBase_StartX
                    control.Top = control.Top + e.Y - gBase_StartY
                    Line_Re_Move(gLab, Val(control.Name), control.Left, control.Top)

                    Dim Tag As String = sender.name
                    With gLab
                        For Each Box As Panel In .DesKTop(.DeskTopID).Controls
                            If Line_Check_Parent(Box.Tag, Tag) = True Then
                                Box.Left = Box.Left + e.X - gBase_StartX
                                Box.Top = Box.Top + e.Y - gBase_StartY
                                Box_Re_Move(gLab, Val(Box.Name), Box)

                            End If
                        Next
                    End With

                    Me.Cursor = Cursors.Hand
            End Select

            gFlag_Pull = "Move"

        End If
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub Line_MouseDown(sender As Object, e As MouseEventArgs)
        '左クリックの場合
        If e.Button = MouseButtons.Left Then
            '移動前の位置を記録
            gBase_StartX = e.X
            gBase_StartY = e.Y
            gBase_W = sender.Width
            gBase_H = sender.Height
            If (sender.Width - gBase_StartX) < 8 And (sender.Height - gBase_StartY) < 8 Then
                gFlag_Pull = "Resize"
            Else
                gFlag_Pull = "Move"
            End If
            Me.Cursor = Cursors.Hand
        End If

    End Sub

    Private Sub LineV_MouseMove(sender As Object, e As MouseEventArgs)
        '左クリックの場合
        If e.Button = MouseButtons.Left Then
            Dim RX, RY As Integer
            'コントロール取得
            Dim control As Control = CType(sender, Control)

            Select Case gFlag_Pull
                Case "Move"  '移動


            End Select
            'コントロールの位置を設定
            RX = e.X - gBase_StartX
            RY = e.Y - gBase_StartY
            control.Top = control.Top + RY
            control.Left = control.Left + RX
            Dim Tag As String = sender.name
            With gLab
                For Each Box As Panel In .DesKTop(.DeskTopID).Controls
                    If Line_Check_Parent(Box.Tag, Tag) = True Then
                        Box.Left = Box.Left + RX
                        Box.Top = Box.Top + RY
                    End If
                Next
            End With

        End If
    End Sub
    Private Sub Line_DragOver(sender As Object, e As DragEventArgs)
        Log_Show(1, e.X & "|" & e.Y)
        gLab.ElementAT.View.ListView.Focus()
        e.Effect = DragDropEffects.Move

    End Sub


    Private Sub Base_MouseLeave(sender As Object, e As System.EventArgs)
        sender.BorderStyle = BorderStyle.None

        Me.Cursor = Cursors.Default


    End Sub
    Private Sub Base_MouseEnter(sender As Object, e As System.EventArgs)
        sender.BorderStyle = BorderStyle.FixedSingle
        'If gLab.BoxID <> Val(sender.name) Then
        sender.BringToFront()
        'End If
        gLab.BoxID = Val(sender.name)
        gLab.Base = sender
        gLab.Flg_BoxMouse = "OnBOX"
        'Log_Show(1, sender.tag)
    End Sub
    Private Sub LineH_MouseEnter(sender As Object, e As System.EventArgs)
        Dim Panel As Panel = CType(sender, Panel)

        Panel.BorderStyle = BorderStyle.FixedSingle
        Panel.BringToFront()

        Panel.Height = 12

        Dim Blist() As DataRow
        Dim dtRow As DataRow
        With gLab
            .LineID = Val(sender.name)
            gLab.Line = sender
            .Base = sender
            .Flg_BoxMouse = "OnH"
            Blist = .BoxList.Select("State = 'Nomal' AND BoxID = '" & .LineID.ToString & "' AND DeskID = '" & .DeskID.ToString & "'")
            If Blist.Count = 0 Then Return '異常
            dtRow = Blist(0)
            .LineNo = Val(dtRow("LineNo"))


            Dim BoxS() As Panel
            Dim cnt As Integer
            Dim Tag As String = .Line.Name
            For Each Box As Panel In .DesKTop(.DeskTopID).Controls
                If Line_Check_Parent(Box.Tag, Tag) = True Then
                    ReDim Preserve BoxS(cnt)
                    BoxS(cnt) = Box
                    cnt += 1
                End If
            Next
            .LineMax = cnt + 1
        End With

    End Sub
    Private Sub LineH_MouseLeave(sender As Object, e As System.EventArgs)

        Dim Panel As Panel = CType(sender, Panel)
        gLab.LineID = 0
        Panel.BorderStyle = BorderStyle.None

        Panel.Height = 4
        gLab.Flg_BoxMouse = ""


    End Sub
    Private Sub LineH_MouseUp(sender As Object, e As MouseEventArgs)
        '左クリックの場合
        If e.Button = MouseButtons.Left Then

            'コントロール取得
            Dim control As Control = CType(sender, Control)

            Select Case gFlag_Pull
                Case "Move" '移動

                    control.Left = control.Left + e.X - gBase_StartX
                    control.Top = control.Top + e.Y - gBase_StartY
                    Line_Re_Move(gLab, Val(control.Name), control.Left, control.Top)

                    '子供チェック
                    Dim Tag As String = sender.name
                    With gLab
                        For Each Box As Panel In .DesKTop(.DeskTopID).Controls
                            If Line_Check_Parent(Box.Tag, Tag) = True Then
                                Box.Left = Box.Left + e.X - gBase_StartX
                                Box.Top = Box.Top + e.Y - gBase_StartY
                                If InStr(Box.TabIndex, "Line") = 0 Then
                                    Box_Re_Move(gLab, Val(Box.Name), Box)
                                Else
                                    Line_Re_Move(gLab, Val(Box.Name), Box.Left, Box.Top)
                                End If
                            End If
                        Next
                    End With

                    '他のラインの上
                    Dim tLID, sLID As Integer
                    tLID = Line_Check_OnLine(gLab, control)  'LINE上か
                    Dim GX() As String
                    GX = Split(control.Tag, ".")

                    If GX.Count > 1 Then 'LINEから要素を外す場合

                        sLID = Val(GX(1))
                        control.Tag = "LineH"
                        Box_Update(gLab, gLab.DeskID, Val(control.Name), "LineID", control.Tag) 'LINE要素の作り変え

                        Line_Set_ByLID(gLab, sLID) '対象ラインにフォーカス
                        If gLab.Line Is Nothing Then Return
                        LineV_Refresh(gLab, 0) '要素を外す場合もとりあえず全展開
                    End If

                    If tLID > 0 Then
                        ' Line_Set_ByLID(gLab, tLID) '対象ラインにフォーカス
                        control.Tag = "LineH." & tLID.ToString
                        control.Left = gLab.Line.Left + gLab.Line.Width
                        Box_Update(gLab, gLab.DeskID, Val(control.Name), "LineID", control.Tag)

                        Line_Set_ByLID(gLab, tLID) '対象ラインにフォーカス
                        If gLab.Line Is Nothing Then Return
                        LineV_Refresh(gLab, 0) 'とりあえず全展開
                    End If
            End Select


            gFlag_Pull = "Move"

        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub LineH_MouseDoubleClick(sender As Object, e As MouseEventArgs)
        'LINEの開閉
        sender.BringToFront()

        With gLab
            .LineID = Val(sender.name)
            .Line = sender
            .Base = sender
            .Flg_BoxMouse = "OnH"
            If .LineNo > 0 Then
                .LineNo = 0
            Else
                .LineNo = .LineMax
            End If

            LineH_Refresh(gLab, .LineNo)
            Box_Update(gLab, .DeskID, .LineID, "LineNo", .LineNo.ToString)
        End With

    End Sub

    Private Sub LineH_MouseMove(sender As Object, e As MouseEventArgs)
        '左クリックの場合
        If e.Button = MouseButtons.Left Then
            Dim RX, RY As Integer
            'コントロール取得
            Dim control As Control = CType(sender, Control)

            Select Case gFlag_Pull
                Case "Move"  '移動


            End Select
            'コントロールの位置を設定
            RX = e.X - gBase_StartX
            RY = e.Y - gBase_StartY
            control.Top = control.Top + RY
            control.Left = control.Left + RX
            Dim Tag As String = sender.name
            With gLab
                For Each Box As Panel In .DesKTop(.DeskTopID).Controls
                    If Line_Check_Parent(Box.Tag, Tag) = True Then
                        Box.Left = Box.Left + RX
                        Box.Top = Box.Top + RY
                    End If
                Next
            End With
            'gBase_StartX = e.X
            'gBase_StartY = e.Y
            'リフレッシュ
            '  Me.Refresh()
        End If
    End Sub

    Private Sub Line_DragDrop(sender As Object, e As DragEventArgs)

        'Desk_Drops(sender, e) 'Stackぽっとん
        MsgBox("ライン上への直接ポットンはまだできてない")

    End Sub


    Private Sub TextBox8_TextChanged(sender As System.Object, e As System.EventArgs)

    End Sub



    Private Sub Proj_Set_Env(ByRef Lab As Lab_Type)
        'Projの基本パラメータを設定　　（データの読み込みはしない
        '-----------------------------
        'Log_Show(1, "Projの基本パラメータを設定")
        'With Lab

        '    ''DB作成
        '    '--------------------------
        '    DBAcs_Init2(.ProjAT.DBACS, "RepoDB", "Proj", "Lab_Key")


        'End With

    End Sub

    Private Sub Desk_Set_Env(ByRef Lab As Lab_Type)

        'Desk設定
        '--------------
        Dim DeskID As Integer
        DeskID = 0

        'With Lab.Desk(DeskID)
        '    'TreeView
        '    '--------------------------
        '    '.Tree = Me.TreeView1
        '    '.Tree.ImageList = gImg_List.Image

        '    'Listview
        '    '--------------------------
        '    '.View = ListView1
        '    'ReDim .Columns(3)
        '    'Columns_Set_Env(.Columns(0), "[Proj]")
        '    '' Columns_Set_Env(.Columns(1), "[Branch]")
        '    'Columns_Set_Listview(.View, .Columns(0))


        '    ''Desk_DB作成
        '    '--------------------------
        '    DBAcs_Init2(.DBACS, "Desk", "Node", "Desk_Node")
        '    'Dim Ds As DataSet
        '    'Dim Cnd As String
        '    'Cnd = "OwnerID = " & Lab.OwnerID.ToString

        '    'Ds = DBAcs_Get_DataS(.DBACS, "ID", Cnd)

        '    'If Ds.Tables(0).Rows.Count = 0 Then  'TBLがからっぽなら、ＤＢに、デフォルト登録

        '    '    .Node_Trash = Desk_Make_Node(Lab, "[Trash]")
        '    '    .Node_Society = Desk_Make_Node(Lab, "[Society]")
        '    '    .Node_Closed = Desk_Make_Node(Lab, "[Closed]")
        '    '    .Node_Opened = Desk_Make_Node(Lab, "[Opened]")
        '    '    .Node_Current = Desk_Make_Node(Lab, "[Current]")

        '    'End If


        '    'If DB_Get_Data_ByKey(.DBACS, .DBACS.TBL_Name, "ID", 1) = "" Then


        '    'Else
        '    '    'DBからの読みだしはもっとあとでやる


        '    'End If



        'End With

    End Sub
    'Private Sub Desk_Read_NodeData(ByRef Lab As Lab_Type)

    '    '    'Desk設定
    '    '    '--------------
    '    '    Dim Ds As DataSet
    '    '    Dim Cnd As String

    '    '    Dim DeskID As Integer
    '    '    DeskID = 1
    '    '    Dim NID As Integer

    '    '    With Lab.Desk(DeskID)
    '    '        Cnd = "OwnerID = " & Lab.OwnerID.ToString
    '    '        Ds = DBAcs_Get_DataS(.DBACS, "*", Cnd)

    '    '        If Ds.Tables(0).Rows.Count = 0 Then  'TBLがからっぽなら、ＤＢに、デフォルト登録

    '    '            Return
    '    '        Else
    '    '            'DBから（自分の）全ノードデータを読み込む "ID,NID,PID,RID,State,NodeType,Name,Owner,OwnerID"
    '    '            For L = 0 To Ds.Tables(0).Rows.Count - 1

    '    '                NID = Val(Ds.Tables(0).Rows(L).Item(1))
    '    '                If NID > .NodeMax Then
    '    '                    .NodeMax = NID
    '    '                    ReDim Preserve Lab.Desk(DeskID).Node(.NodeMax)
    '    '                End If

    '    '                With Lab.Desk(DeskID).Node(NID)
    '    '                    .ID = Val(Ds.Tables(0).Rows(L).Item(0))
    '    '                    .PID = Val(Ds.Tables(0).Rows(L).Item(2))
    '    '                    .RID = Val(Ds.Tables(0).Rows(L).Item(3))
    '    '                    .State = Trim(Ds.Tables(0).Rows(L).Item(4))
    '    '                    .NodeType = Trim(Ds.Tables(0).Rows(L).Item(5))
    '    '                    .Name = Trim(Ds.Tables(0).Rows(L).Item(6))
    '    '                    .Owner = Trim(Ds.Tables(0).Rows(L).Item(7))
    '    '                    .OwnerID = Val(Ds.Tables(0).Rows(L).Item(8))
    '    '                End With
    '    '            Next
    '    '        End If
    '    '    End With

    'End Sub

    Private Sub Element_Set_Env(ByRef Lab As Lab_Type)
        'Elmの基本パラメータを設定　　（データの読み込みはしない
        '-----------------------------
        Log_Show(1, "Elementの基本パラメータを設定")

        With Lab.ElementAT.View
            .ListView = ListView2
            Columns_Set_Env(.Columns, "[Element]")
            Columns_Set_Listview(.ListView, .Columns)
        End With

    End Sub










    'Private Sub CheckedListBox1_DragOver(sender As Object, e As System.Windows.Forms.DragEventArgs)
    '    If e.Data.GetDataPresent(DataFormats.FileDrop) = True Then
    '        e.Effect = DragDropEffects.Copy
    '    Else
    '        e.Effect = DragDropEffects.None
    '    End If
    'End Sub

    Private Sub GrouPbox1_Enter(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub SplitContainer9_Panel2_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs)

    End Sub

    Private Sub SplitContainer8_Panel2_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs)

    End Sub

    Private Sub SplitContainer13_SplitterMoved(sender As System.Object, e As System.Windows.Forms.SplitterEventArgs)

    End Sub

    Private Sub ToolStripLabel1_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed

    End Sub

    Private Sub ListView4_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form1_Invalidated(sender As Object, e As InvalidateEventArgs) Handles Me.Invalidated

    End Sub

    Private Sub ToolStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs)

    End Sub

    Private Sub ContextMenuStrip1_MouseClick(sender As Object, e As MouseEventArgs) Handles ContextMenuStrip1.MouseClick

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub ContextMenuStrip1_MouseDown(sender As Object, e As MouseEventArgs) Handles ContextMenuStrip1.MouseDown

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs)

    End Sub



    Private Sub ListView2_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView2.MouseDoubleClick

        With ListView2
            If .SelectedItems.Count = 0 Then Return
            Element_Act_Item(gLab, .SelectedItems(0).Tag)
        End With
    End Sub





    Private Sub Button2_Click_3(sender As Object, e As EventArgs)

        'OWNER プロファイル登録
        '-------------------


        'With gLab.Person(0)
        '    .Image = PictureBox1.ImageLocation
        '    .Name = TextBox4.Text
        '    .Tel = TextBox2.Text
        '    .Mail = TextBox3.Text
        '    .Prof = TextBox5.Text
        'End With

        Person_Set_NewData(gLab)
        Lab_Start(gLab)
        MsgBox("登録されました")
    End Sub



    Private Sub PictureBox1_DragEnter(sender As Object, e As DragEventArgs)
        If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub PictureBox2_MouseEnter(sender As Object, e As EventArgs)
        sender.BackColor = Color.White
        'PictureBox2.BackColor =Color.FromName(gdeco.BackColor)
    End Sub


    Private Sub PictureBox2_MouseLeave(sender As Object, e As EventArgs)
        sender.BackColor = Color.FromName(gDeco.BackColor)
    End Sub

    Private Sub PictureBox3_MouseLeave(sender As Object, e As EventArgs)
        sender.BackColor = Color.FromName(gDeco.BackColor)
    End Sub

    Private Sub PictureBox3_MouseEnter(sender As Object, e As EventArgs)
        sender.BackColor = Color.White
    End Sub

    Private Sub PictureBox4_MouseEnter(sender As Object, e As EventArgs)
        sender.BackColor = Color.White
    End Sub

    Private Sub PictureBox4_MouseLeave(sender As Object, e As EventArgs)
        sender.BackColor = Color.FromName(gDeco.BackColor)
    End Sub



    Private Sub Lab_Act(Lab As Lab_Type, Key As String)



    End Sub

    Private Sub PictureBox2_MouseClick(sender As Object, e As MouseEventArgs)


        Lab_Act(gLab, sender.name)


    End Sub

    Private Sub PictureBox3_MouseClick(sender As Object, e As MouseEventArgs)

        Lab_Act(gLab, sender.name)
    End Sub

    Private Sub PictureBox6_MouseEnter(sender As Object, e As EventArgs)
        sender.BackColor = Color.White
    End Sub

    Private Sub PictureBox6_MouseLeave(sender As Object, e As EventArgs)
        sender.BackColor = Color.FromName(gDeco.BackColor)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Proj_Search(gLab)

        '  ListView1.Visible = True

    End Sub
    Private Sub Proj_Search(lab As Lab_Type)

        Dim DS As DataSet
        Dim Cmd, Key As String
        Key = TextBox6.Text
        Key = Replace(Key, Chr(13), "")
        Key = Replace(Key, Chr(10), "")


        If Key = "" Or Key = "*" Then
            Cmd = " State = 'Nomal' "
            'Cmd &= "AND OwnerID = '" & gX_ownerID.ToString & "'"
            Proj_Show_ListByCmd(gLab, Cmd) 'オーナーのレポ表示

        Else
            '  Cmd = " State ='[Opened]' AND OwnerID <> " & gLab.OwnerID
            Cmd = " State = 'Nomal' "
            Cmd &= "AND Title LIKE '%" & Key & "%'"
            Proj_Show_ListByCmd(gLab, Cmd)
        End If



    End Sub
    ''Private Sub Proj_Cast(lab As Lab_Type)

    ''    Dim DS As DataSet
    ''    Dim Cmd, Key As String
    ''    Key = TextBox6.Text
    ''    Key = Replace(Key, Chr(13), "")
    ''    Key = Replace(Key, Chr(10), "")


    ''    If Key = "" Or Key = "*" Then
    ''        Cmd = " State > 0 "
    ''        Cmd &= "AND OwnerID = " & gX_ownerID
    ''        Proj_Show_ListByCmd(gLab, Cmd) 'オーナーのレポ表示

    ''    Else
    ''        '  Cmd = " State ='[Opened]' AND OwnerID <> " & gLab.OwnerID
    ''        Cmd = " State = 1 "
    ''        Cmd &= "AND Name LIKE '%" & Key & "%'"
    ''        Proj_Show_ListByCmd(gLab, Cmd)
    ''    End If



    'End Sub
    Private Sub Button4_Click_1(sender As Object, e As EventArgs)
        'WebView21.Visible = True
        ''WebView21.Dock = "fill"

        'ListView1.Visible = False
        '' ListView4.Dock = "non"
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs)
        Lab_Act(gLab, sender.name)
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs)
        Lab_Act(gLab, sender.name)
    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs)
        Lab_Act(gLab, sender.name)
    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs)
        Lab_Act(gLab, sender.name)
    End Sub
    Private Sub Lab_Selected_TabPage(ByRef Lab As Lab_Type, Key As String)



    End Sub



    Private Sub PictureBox6_Leave(sender As Object, e As EventArgs)
        sender.BackColor = Color.FromName(gDeco.BackColor)
    End Sub

    Private Sub PictureBox6_Enter(sender As Object, e As EventArgs)
        sender.BackColor = Color.White
    End Sub

    Private Sub PictureBox5_MouseEnter(sender As Object, e As EventArgs)
        sender.BackColor = Color.White
    End Sub

    Private Sub PictureBox5_MouseLeave(sender As Object, e As EventArgs)
        sender.BackColor = Color.FromName(gDeco.BackColor)
    End Sub

    Private Sub PictureBox5_Click_1(sender As Object, e As EventArgs)
        Lab_Act(gLab, sender.name)
    End Sub

    Private Sub ListView1_MouseDown(sender As Object, e As MouseEventArgs) Handles ListView1.MouseDown

    End Sub

    Private Sub ListView1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseDoubleClick

        With ListView1
            If .SelectedItems Is Nothing Then Return
            If .SelectedItems.Count <= 0 Then Return

            If .SelectedItems.Count > 1 Then Return

            If .SelectedItems(0).Tag Is Nothing Then Return

            gProjKey = .SelectedItems(0).Tag


        End With
        'Proj_Show_Data(gLab, gProjID)

    End Sub

    Private Async Sub InitializeAsync()
        ' Await WebView21.EnsureCoreWebView2Async(Nothing)
    End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs)


        '  WebView21.Source = New Uri("https://freecalend.com/open/mem179529/index")
        '  WebView21.Source = New Uri("C:\Users\miura\Documents\ＪＳＡライフ２０２１.pdf")
        'WebView21.Source = New Uri("file:///C:/Users/miura/Documents/XXtest_.html")
        '   WebView21.Source = New Uri("https://query.wikidata.org/sparql?query=SELECT%20DISTINCT%20%20%3Fitem%20%3FitemLabel%20%3FgenderLabel%20%3FcountryLabel%20%3FbirthdateLabel%20%3FbirthplaceLabel%20%3FjobLabel%20%3FemployerLabel%20%3FinstLabel%20WHERE%20%7B%0A%20%20SERVICE%20wikibase%3Alabel%20%7B%20bd%3AserviceParam%20wikibase%3Alanguage%20%22ja%22.%20%7D%0A%20%20%7B%0A%20%20%20%20SELECT%20DISTINCT%20%20*%20WHERE%20%7B%0A%20%20%20%20%20%20%3Fitem%20rdfs%3Alabel%20%22%E4%BD%90%E5%80%89%E7%9C%9F%E8%A1%A3%22%40ja%20.%0A%20%20%20%20%20%20%23%3Fitem%20rdfs%3Alabel%20%22%E6%A1%9C%E4%BA%95%E6%B4%8B%E5%AD%90%22%40ja%20.%0A%20%20%20%20%20%20%3Fitem%20wdt%3AP21%20%3Fgender.%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP27%20%3Fcountry.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP569%20%3Fbirthdate.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP19%20%3Fbirthplace.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP106%20%3Fjob.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP108%20%3Femployer.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP1303%20%3Finst.%7D%0A%20%20%20%20%20%0A%20%20%20%20%7D%0A%20%20%20%20LIMIT%201000%0A%20%20%7D%0A%7D")
        Dim tmpURL, TmpKey As String
        tmpURL = "https://query.wikidata.org/embed.html#SELECT%20DISTINCT%20%20%3Fitem%20%3FitemLabel%20%3FgenderLabel%20%3FcountryLabel%20%3FbirthdateLabel%20%3FbirthplaceLabel%20%3FjobLabel%20%3FemployerLabel%20%3FinstLabel%20WHERE%20%7B%0A%20%20SERVICE%20wikibase%3Alabel%20%7B%20bd%3AserviceParam%20wikibase%3Alanguage%20%22ja%22.%20%7D%0A%20%20%7B%0A%20%20%20%20SELECT%20DISTINCT%20%20*%20WHERE%20%7B%0A%20%20%20%20%20%20%3Fitem%20rdfs%3Alabel%20%22%40%40JINMEI%40%40%22%40ja%20.%0A%20%20%20%20%20%20%23%3Fitem%20rdfs%3Alabel%20%22%E6%A1%9C%E4%BA%95%E6%B4%8B%E5%AD%90%22%40ja%20.%0A%20%20%20%20%20%20%3Fitem%20wdt%3AP21%20%3Fgender.%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP27%20%3Fcountry.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP569%20%3Fbirthdate.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP19%20%3Fbirthplace.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP106%20%3Fjob.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP108%20%3Femployer.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP1303%20%3Finst.%7D%0A%20%20%20%20%20%0A%20%20%20%20%7D%0A%20%20%20%20LIMIT%201000%0A%20%20%7D%0A%7D"
        TmpKey = System.Web.HttpUtility.UrlEncode("岩倉具視")
        tmpURL = Replace(tmpURL, "%40%40JINMEI%40%40", TmpKey)
        '  WebView21.CoreWebView2.Navigate(tmpURL)
        ' WebView21.CoreWebView2.Navigate("https://query.wikidata.org/embed.html#SELECT%20DISTINCT%20%20%3Fitem%20%3FitemLabel%20%3FgenderLabel%20%3FcountryLabel%20%3FbirthdateLabel%20%3FbirthplaceLabel%20%3FjobLabel%20%3FemployerLabel%20%3FinstLabel%20WHERE%20%7B%0A%20%20SERVICE%20wikibase%3Alabel%20%7B%20bd%3AserviceParam%20wikibase%3Alanguage%20%22ja%22.%20%7D%0A%20%20%7B%0A%20%20%20%20SELECT%20DISTINCT%20%20*%20WHERE%20%7B%0A%20%20%20%20%20%20%3Fitem%20rdfs%3Alabel%20%22%E4%BD%90%E5%80%89%E7%9C%9F%E8%A1%A3%22%40ja%20.%0A%20%20%20%20%20%20%23%3Fitem%20rdfs%3Alabel%20%22%E6%A1%9C%E4%BA%95%E6%B4%8B%E5%AD%90%22%40ja%20.%0A%20%20%20%20%20%20%3Fitem%20wdt%3AP21%20%3Fgender.%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP27%20%3Fcountry.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP569%20%3Fbirthdate.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP19%20%3Fbirthplace.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP106%20%3Fjob.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP108%20%3Femployer.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP1303%20%3Finst.%7D%0A%20%20%20%20%20%0A%20%20%20%20%7D%0A%20%20%20%20LIMIT%201000%0A%20%20%7D%0A%7D")

        ' WebView21.CoreWebView2.Navigate("https://www.microsoft.com/ja-jp")

        'Dim html = Await WebView21.ExecuteScriptAsync("document.documentElement.outerHTML")

        ' html = Regex.Unescape(html)

        'TextBox7.Text = html.ToString

        'Dim sUrl As String   '// ダウンロード対象ファイルのURL
        'Dim sDir As String   '// ダウンロードファイルを保存するローカルPCのフォルダパス

        ''// URL設定
        'sUrl = = "https://query.wikidata.org/embed.html#SELECT%20DISTINCT%20%20%3Fitem%20%3FitemLabel%20%3FgenderLabel%20%3FcountryLabel%20%3FbirthdateLabel%20%3FbirthplaceLabel%20%3FjobLabel%20%3FemployerLabel%20%3FinstLabel%20WHERE%20%7B%0A%20%20SERVICE%20wikibase%3Alabel%20%7B%20bd%3AserviceParam%20wikibase%3Alanguage%20%22ja%22.%20%7D%0A%20%20%7B%0A%20%20%20%20SELECT%20DISTINCT%20%20*%20WHERE%20%7B%0A%20%20%20%20%20%20%3Fitem%20rdfs%3Alabel%20%22%E4%BD%90%E5%80%89%E7%9C%9F%E8%A1%A3%22%40ja%20.%0A%20%20%20%20%20%20%23%3Fitem%20rdfs%3Alabel%20%22%E6%A1%9C%E4%BA%95%E6%B4%8B%E5%AD%90%22%40ja%20.%0A%20%20%20%20%20%20%3Fitem%20wdt%3AP21%20%3Fgender.%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP27%20%3Fcountry.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP569%20%3Fbirthdate.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP19%20%3Fbirthplace.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP106%20%3Fjob.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP108%20%3Femployer.%7D%0A%20%20%20%20%20%20OPTIONAL%20%7B%3Fitem%20wdt%3AP1303%20%3Finst.%7D%0A%20%20%20%20%20%0A%20%20%20%20%7D%0A%20%20%20%20LIMIT%201000%0A%20%20%7D%0A%7D"
        'sDir = "C:\Users\miura\Downloads\Test"

        ''// ダウンロード
        'URLDownloadToFile(sUrl, sDir)


    End Sub




    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ListView2_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs) Handles ListView2.ItemSelectionChanged

    End Sub





    Private Sub Panel1_DragDrop(sender As Object, e As DragEventArgs)
        Desk_Drops(sender, e)



    End Sub
    Private Sub Desk_DoubleClick(sender As Object, e As MouseEventArgs)
        'BOXを全部元位置に

        Dim Dlist() As DataRow
        Dim dtRow As DataRow



        Dim SX As Integer = 10
        With gLab

            .DeskID = Val(sender.tag)
            .DeskTopID = Val(sender.name)
            Dlist = .DeskList.Select("State = 'Nomal' AND DeskID = '" & .DeskID.ToString & "'")
            If Dlist.Count = 0 Then Return
            dtRow = Dlist(0)

            With .Desk(.DeskID)
                For Each box As Panel In sender.Controls
                    Select Case box.Tag
                        Case "LINE"
                        Case "OnDESK"

                            ''拡大縮小を元に
                            'For L = 1 To Math.Abs(.ZM)
                            '    box.Left = Box_ZM_Move(box.Left, .ShiftX, .ZM, SX)
                            '    box.Top = Box_ZM_Move(box.Top, .ShiftY, .ZM, SX)
                            '    box.Width = Box_ZM(box.Width, .ZM, SX)
                            '    box.Height = Box_ZM(box.Height, .ZM, SX)
                            'Next

                            'box.Left = box.Left - .ShiftX
                            'box.Top = box.Top - .ShiftY
                            'Box_Re_Move(gLab, Val(box.Name), box)
                            'Box_Re_Size(gLab, Val(box.Name), box.Width, box.Height)

                    End Select
                Next
                .ZM = 0
            End With
            Dim Dat As String
            With .DesKTop(.DeskTopID)
                .Top = gLab.WorldX
                .Left = gLab.WorldY

            End With

            dtRow("ShiftX") = 0
            dtRow("ShiftY") = 0
            Desk_Save_Info(dtRow)
        End With
        '   Map_View(gLab, XX, YY)


    End Sub
    Private Sub Map_MouseUp(sender As Object, e As MouseEventArgs)

        '移動後の位置を記録

        Me.Cursor = Cursors.Arrow

    End Sub

    Private Sub Map_MouseDown(sender As Object, e As MouseEventArgs)
        ' Log_Show(1, "X=" & e.X & "  Y=" & e.Y & "|" & gLab.DeskSx.ToString & " |" & gLab.DeskSy.ToString)
        If e.Button = MouseButtons.Left Then
            '移動前の位置を記録
            gMap_StartX = e.X
            gMap_StartY = e.Y
        End If
        'With gLab
        '    If .DeskXX = 0 Then
        '        .DeskXX = .DesKTop(.DeskTopID).Left
        '        .DeskYY = .DesKTop(.DeskTopID).Top

        '    End If
        'End With

        'gLab.Flg_BoxMouse = "OnDESK"
        Me.Cursor = Cursors.Hand
        Log_Show(2, "X=" & e.X & "  Y=" & e.Y & "|")
    End Sub
    Private Sub Desk_MouseDown(sender As Object, e As MouseEventArgs)
        ' Log_Show(1, "X=" & e.X & "  Y=" & e.Y & "|" & gLab.DeskSx.ToString & " |" & gLab.DeskSy.ToString)
        If e.Button = MouseButtons.Left Then
            '移動前の位置を記録
            gBase_StartX = e.X
            gBase_StartY = e.Y


        End If
        With gLab
            If .DeskXX = 0 Then
                .DeskXX = .DesKTop(.DeskTopID).Left
                .DeskYY = .DesKTop(.DeskTopID).Top

            End If
        End With

        gLab.Flg_BoxMouse = "OnDESK"
        Me.Cursor = Cursors.Hand
        Log_Show(2, "X=" & e.X & "  Y=" & e.Y & "|")
    End Sub


    Private Sub Desk_DragDrop(sender As Object, e As DragEventArgs)
        Desk_Drops(sender, e) 'Stackぽっとん
    End Sub

    Private Sub Desk_DragOver(sender As Object, e As DragEventArgs)

        gLab.ElementAT.View.ListView.Focus()
        e.Effect = DragDropEffects.Move

    End Sub

    Private Sub WebView21_Click(sender As Object, e As EventArgs)

    End Sub
    Dim gBase_StartX As Integer
    Dim gBase_StartY As Integer
    Dim gMap_StartX As Integer
    Dim gMap_StartY As Integer
    Dim gFlag_Pull As String
    Dim gBase_W As Integer
    Dim gBase_H As Integer

    Private Sub Base_DoubleClick(sender As Object, e As MouseEventArgs)
        If sender.dock = DockStyle.None Then
            sender.dock = DockStyle.Fill
        Else
            sender.dock = DockStyle.None
        End If
    End Sub

    Private Sub Base_MouseDown(sender As Object, e As MouseEventArgs)
        '左クリックの場合
        If e.Button = MouseButtons.Left Then
            '移動前の位置を記録
            gBase_StartX = e.X
            gBase_StartY = e.Y
            gBase_W = sender.Width
            gBase_H = sender.Height
            Dim tmpP As String
            tmpP = ""
            tmpP &= gBase_StartX & "|"
            tmpP &= gBase_StartY & "|"
            tmpP &= sender.left & "|"
            tmpP &= sender.top & "|"
            tmpP &= gBase_W & "|"
            tmpP &= gBase_H & "|"
            Log_Show(1, tmpP)

            gFlag_Pull = ""

            If (sender.Width - gBase_StartX) < 10 And (sender.Height - gBase_StartY) < 10 Then
                gFlag_Pull = "Resize"
                Me.Cursor = Cursors.SizeNWSE
            End If

            If (gBase_StartX) < 10 And (sender.Height - gBase_StartY) < 10 Then
                gFlag_Pull = "Resize"
                Me.Cursor = Cursors.SizeNESW
            End If

            If (sender.Width - gBase_StartX) < 10 And Math.Abs((sender.Height) / 2 - gBase_StartY) < 10 Then
                gFlag_Pull = "Resize"
                Me.Cursor = Cursors.SizeWE
            End If

            If Math.Abs((sender.Width) / 2 - gBase_StartX) < 10 And (sender.Height - gBase_StartY) < 10 Then
                gFlag_Pull = "Resize"
                Me.Cursor = Cursors.SizeNS
            End If


            If (gBase_StartX) < 10 And Math.Abs((sender.Height) / 2 - gBase_StartY) < 10 Then
                gFlag_Pull = "Resize"

                Me.Cursor = Cursors.SizeWE
            End If


            If gFlag_Pull = "" Then
                gFlag_Pull = "Move"
                Me.Cursor = Cursors.Hand
            End If
        End If
    End Sub
    Private Sub Desk_MouseUp(sender As Object, e As MouseEventArgs)
        Log_Show(1, sender.tag & sender.name)

        Dim Dlist() As DataRow
        Dim dtRow As DataRow
        With gLab
            .DeskID = Val(sender.tag)
            .DeskTopID = Val(sender.name)
            Dlist = .DeskList.Select("State = 'Nomal' AND DeskID = '" & .DeskID.ToString & "'")
            If Dlist.Count = 0 Then Return
            dtRow = Dlist(0)
            dtRow("ShiftX") = .DesKTop(.DeskTopID).Left - .WorldX
            dtRow("ShiftY") = .DesKTop(.DeskTopID).Top - .WorldY
            Desk_Save_Info(dtRow) 'ファイル書き込み

            ' Map_Show_Direction(.DeskXX, .DeskYY)
            .MvX = .DeskXX - .DesKTop(.DeskTopID).Left
            .MvY = .DeskYY - .DesKTop(.DeskTopID).Top
            Log_Show(2, "[VECT]" & gLab.MvX & "|" & gLab.MvY)
            .DeskXX = 0


        End With




        gLab.Flg_DeskMove = 0
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub Desk_MouseMove(sender As Object, e As MouseEventArgs)
        Dim XX, YY As Long
        ''左クリックの場合
        'With gLab
        '    Log_Show(1, .DeskTopID.ToString & "|" & e.X & "|" & e.Y & "|MAP" & .Desk(.DeskID).DTop.Left & "|" & .Desk(.DeskID).DTop.Top & "|SHIFT" & .Desk(.DeskID).DTop.Left & "|" & .Desk(.DeskID).DTop.Top)
        'End With


        If e.Button = MouseButtons.Left Then
            With gLab


                '微小な変化を削除
                XX = e.X - gBase_StartX
                YY = e.Y - gBase_StartY
                If Math.Abs(XX) < 1 And Math.Abs(YY) < 1 Then Return

                .Flg_DeskMove = 1
                With .DesKTop(.DeskTopID)
                    .Left = .Left + XX
                    .Top = .Top + YY

                    '  Map_View(gLab, .Left, .Top)

                End With



            End With



        End If
    End Sub
    Private Sub Map_View(ByRef lab As Lab_Type, SX As Integer, SY As Integer) ', ZM As Integer)

        With gMap
            .Left = Math.Abs(Int(gLab.WorldX - SX) * gWorkPanel(0).Width * 4 / gLab.WorldW)
            .Top = Math.Abs(Int(gLab.WorldY - SY) * gWorkPanel(0).Height * 4 / gLab.WorldH)
            'Select Case ZM
            '    Case 0
            '        .Width = W
            '        .Height = H
            '    Case > 0
            '        .Width = W * (1 + ZM / BX)
            '        .Height = H * (1 + ZM / BX)
            '    Case < 0
            '        .Width = W * (1 + ZM / BX)
            '        .Height = H * (1 + ZM / BX)
            'End Select
            'Log_Show(2, .Left.ToString & "|" & .Top.ToString & "|" & .Width.ToString & "|" & .Height.ToString)

        End With


    End Sub
    Private Sub ColorGrid_MouseMove(sender As Object, e As MouseEventArgs)


        If e.Button <> MouseButtons.Left Then Return '左クリックの場合
        If gColorName = "" Then Return

        Dim effect As DragDropEffects = sender.DoDragDrop("[Col]" & gColorName, DragDropEffects.Copy Or DragDropEffects.Move)

    End Sub

    Private Sub Base_MouseMove(sender As Object, e As MouseEventArgs)

        Log_Show(2, "|" & sender.left & "|" & sender.top & "|" & e.X & "|" & e.Y)
        '左クリックの場合
        If e.Button = MouseButtons.Left Then
            Dim Wd, He As Integer
            'コントロール取得
            Dim control As Control = CType(sender, Control)

            Select Case gFlag_Pull
                Case "Move" '移動
                    'コントロールの位置を設定
                    control.Left = control.Left + e.X - gBase_StartX
                    control.Top = control.Top + e.Y - gBase_StartY
                    Me.Cursor = Cursors.Hand
                  '  Box_Re_Move(gLab, control.name, control.Left, control.Top)
                Case "Resize" '1 'リサイズ
                    Wd = gBase_W + e.X - gBase_StartX
                    He = gBase_H + e.Y - gBase_StartY
                    If Wd < 10 Then Wd = 10
                    If He < 10 Then He = 10
                    control.Width = Wd
                    control.Height = He
                    ' Me.Cursor = Cursors.SizeNESW
                    ' Box_Re_Size(gLab, control.name, control.Left, control.Top)
            End Select

            'リフレッシュ
            ' Me.Refresh()
        Else
            'マウスカーソルを変化させるだけ
            Dim fflag As Integer
            fflag = 0
            If (sender.Width - e.X) < 10 And (sender.Height - e.Y) < 10 Then
                fflag = 1
                Me.Cursor = Cursors.SizeNWSE
            End If

            If (e.X) < 10 And (sender.Height - e.Y) < 10 Then
                fflag = 1
                Me.Cursor = Cursors.SizeNESW
            End If

            If (sender.Width - e.X) < 10 And Math.Abs((sender.Height) / 2 - e.Y) < 10 Then
                fflag = 1
                Me.Cursor = Cursors.SizeWE
            End If

            If Math.Abs((sender.Width) / 2 - e.X) < 10 And (sender.Height - e.Y) < 10 Then
                fflag = 1
                Me.Cursor = Cursors.SizeNS
            End If


            If (e.X) < 10 And Math.Abs((sender.Height) / 2 - e.Y) < 10 Then
                fflag = 1

                Me.Cursor = Cursors.SizeWE
            End If


            If fflag = 0 Then

                Me.Cursor = Cursors.Hand
            End If



        End If
    End Sub
    Private Sub Map_MouseMove(sender As Object, e As MouseEventArgs)

        Log_Show(2, "|" & sender.left & "|" & sender.top & "|" & e.X & "|" & e.Y)


        '左クリックの場合
        If e.Button = MouseButtons.Left Then
            Dim MvX As Integer = e.X - gMap_StartX
            Dim MvY As Integer = e.Y - gMap_StartY
            'コントロール取得
            Dim control As Control = CType(sender, Control)
            'コントロールの位置を設定


            control.Left = control.Left + MvX
            control.Top = control.Top + MvY
            Me.Cursor = Cursors.Hand

            With gLab.DesKTop(gLab.DeskTopID)
                .Left = .Left - MvX * 10
                .Top = .Top - MvY * 10
            End With



        End If


    End Sub

    Private Sub Line_Re_Move(ByRef lab As Lab_Type, BoxID As Long, X As Integer, Y As Integer)
        'DB記録
        'LineH@LineH|0|DESK|X|Y|1|640|5|blue   
        Try
            Dim BluePrint As String
            Dim STACKS() As String
            Dim STK_PRT() As String

            BluePrint = Box_Get_RowDat(lab.BoxList, lab.DeskID, BoxID, "BluePrint")

            STACKS = Split(BluePrint, "@")
            STK_PRT = Split(STACKS(1), "|")

            STK_PRT(3) = (X + lab.WorldX).ToString
            STK_PRT(4) = (Y + lab.WorldY).ToString


            STACKS(1) = String.Join("|", STK_PRT)
            BluePrint = String.Join("@", STACKS)
            Box_Update(lab, lab.DeskID, BoxID, "BluePrint", BluePrint)

        Catch ex As Exception
            MsgBox(ex.Message)
            With lab.BoxList
                For L = 0 To .Rows.Count - 1
                    For K = 0 To .Columns.Count - 1
                        Debug.Print(.Columns(K).Caption & "=" & .Rows(L).Item(K))
                    Next
                Next
            End With

        End Try


    End Sub

    Private Sub Line_Re_Size(ByRef lab As Lab_Type, BoxID As Long, W As Integer, H As Integer)
        '設計図"@LineH|0|DESK|[BX]|[BY]|[BZ]|W|H|Col"
        'DB記録
        Try


            Dim BluePrint As String
            Dim STACKS() As String
            Dim STK_PRT() As String

            BluePrint = Box_Get_RowDat(lab.BoxList, lab.DeskID, BoxID, "BluePrint")

            STACKS = Split(BluePrint, "@")
            STK_PRT = Split(STACKS(1), "|")

            STK_PRT(6) = W.ToString
            STK_PRT(7) = H.ToString

            STACKS(1) = String.Join("|", STK_PRT)
            BluePrint = String.Join("@", STACKS)
            Box_Update(lab, lab.DeskID, BoxID, "BluePrint", BluePrint)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub


    Private Sub Box_Re_Move(ByRef lab As Lab_Type, BoxID As Long, ByRef box As Panel)
        'DB記録
        Dim BluePrint As String
        Dim STACKS() As String
        Dim STK_PRT() As String
        Dim Blist() As DataRow
        Dim dtRow As DataRow

        With lab
            Blist = .BoxList.Select("State = 'Nomal' AND BoxID = '" & BoxID.ToString & "' AND DeskID = '" & lab.DeskID.ToString & "'")
            If Blist.Count = 0 Then Return '異常
            dtRow = Blist(0)
            STACKS = Split(dtRow("BluePrint"), "@")
            STK_PRT = Split(STACKS(1), "|")

            STK_PRT(4) = (box.Left + lab.WorldX).ToString
            STK_PRT(5) = (box.Top + lab.WorldY).ToString
            STK_PRT(6) = ""
            STK_PRT(7) = box.Width.ToString
            STK_PRT(8) = box.Height.ToString


            STACKS(1) = String.Join("|", STK_PRT)
            BluePrint = String.Join("@", STACKS)
        End With

        Box_Update(lab, lab.DeskID, BoxID, "BluePrint", BluePrint)

        '  DBAcs_Update(lab, "Box", "BluePrint = '" & BluePrint & "'", "ID = " & BoxID.ToString)

        'パラメータ一覧
        '@BASE|0|DESK|Dock|X|Y|Z|W|H|"
        '@PICT|0|BASE|Dock|C|DAT|"
        '@TEXT|0|BASE|Dock|Mul|Bc|Fc|FName|FSize|Text|"
        '@WebView|0|BASE|Dock|URL|"
        '@TAB|0|BASE|Dock|Alignment|ItemSizeX|ItemSizeY|"
        '@TABp|0|TAB|Text|BackColor|"
        '@CogT|0|NEXT|Dock|TEXT|"
        '@Draw|0|NEXT|Dock|DAT|
        '@Split|0|NEXT|Dock|Orientation|SplitterDistance|


    End Sub
    Private Sub Box_Re_Size(ByRef lab As Lab_Type, BoxID As Long, W As Integer, H As Integer)
        'パラメータ一覧
        '@BASE|0|DESK|Dock|X|Y|Z|W|H|"
        'DB記録
        Dim BluePrint As String
        Dim STACKS() As String
        Dim STK_PRT() As String
        Dim Blist() As DataRow
        Dim dtRow As DataRow

        With lab
            Blist = .BoxList.Select("State = 'Nomal' AND BoxID = '" & BoxID.ToString & "' AND DeskID = '" & lab.DeskID.ToString & "'")
            If Blist.Count = 0 Then Return '異常
            dtRow = Blist(0)

            STACKS = Split(dtRow("BluePrint"), "@")
            STK_PRT = Split(STACKS(1), "|")

            STK_PRT(7) = W.ToString
            STK_PRT(8) = H.ToString

            STACKS(1) = String.Join("|", STK_PRT)
            BluePrint = String.Join("@", STACKS)
        End With

        Box_Update(lab, lab.DeskID, lab.BoxID, "BluePrint", BluePrint)




    End Sub
    'Private Function Box_Re_Text(ByRef lab As Lab_Type, BoxID As Long, StackID As Integer, Text As String) As String
    '    '設計図上のテキストをDBへのアクセスに書き換え
    '    '@TEXT|0|BASE|Dock|Mul|Bc|Fc|FName|FSize|Text|"
    '    'DB記録
    '    'Dim BluePrint As String
    '    'Dim STACKS() As String
    '    'Dim STK_PRT() As String
    '    'Dim Ds As DataSet
    '    'Ds = DBAcs_Get_DataS(lab, "Box", "BluePrint", "ID = " & BoxID.ToString)

    '    'If Ds.Tables(0).Rows.Count = 0 Then Return "" '異常

    '    'STACKS = Split(Ds.Tables(0).Rows(0).Item(0), "@")

    '    'STK_PRT = Split(STACKS(StackID), "|")

    '    'STK_PRT(13) = Text

    '    'STACKS(StackID) = String.Join("|", STK_PRT)
    '    'BluePrint = String.Join("@", STACKS)
    '    'DBAcs_Update(lab, "Box", "BluePrint = '" & BluePrint & "'", "ID = " & BoxID.ToString)
    '    Return BluePrint
    'End Function
    'Private Sub Box_Re_Draw(ByRef lab As Lab_Type, BoxID As Long, StackID As Integer, DBID As String)
    '    '設計図上のデータをDBへのアクセスに書き換え
    '    '@Draw|0|NEXT|Dock|DaT| 
    '    'DB記録
    '    Dim BluePrint As String
    '    Dim STACKS() As String
    '    Dim STK_PRT() As String
    '    Dim Ds As DataSet
    '    Ds = DBAcs_Get_DataS(lab, "Box", "BluePrint", "ID = " & BoxID.ToString)

    '    If Ds.Tables(0).Rows.Count = 0 Then Return '異常

    '    STACKS = Split(Ds.Tables(0).Rows(0).Item(0), "@")

    '    STK_PRT = Split(STACKS(StackID), "|")

    '    STK_PRT(4) = DBID

    '    STACKS(StackID) = String.Join("|", STK_PRT)
    '    BluePrint = String.Join("@", STACKS)
    '    DBAcs_Update(lab, "Box", "BluePrint = '" & BluePrint & "'", "ID = " & BoxID.ToString)

    'End Sub
    Private Sub Box_Re_Pict(ByRef lab As Lab_Type, BoxID As Long, StackID As Integer, DBID As String)
        '設計図上のデータをDBへのアクセスに書き換え
        '@PICT|0|BASE|Dock|C|DAT|"
        'DB記録
        Dim BluePrint As String
        Dim STACKS() As String
        Dim STK_PRT() As String
        Dim Ds As DataSet
        Ds = DBAcs_Get_DataS(lab, "Box", "BluePrint", "ID = " & BoxID.ToString)

        If Ds.Tables(0).Rows.Count = 0 Then Return '異常

        STACKS = Split(Ds.Tables(0).Rows(0).Item(0), "@")

        STK_PRT = Split(STACKS(StackID), "|")

        STK_PRT(5) = DBID

        STACKS(StackID) = String.Join("|", STK_PRT)
        BluePrint = String.Join("@", STACKS)
        '   DBAcs_Update(lab, "Box", "BluePrint = '" & BluePrint & "'", "ID = " & BoxID.ToString)

    End Sub
    Private Function Base_Check_OnLine(ByRef Lab As Lab_Type, top As Integer, left As Integer) As Long
        'LINE 上かのチェック
        '----------------------

        With Lab
            For Each Line As Panel In .DesKTop(.DeskTopID).Controls

                If InStr(Line.Tag, "LineH") > 0 Then

                    If Math.Abs(Line.Top - top) < 10 And Line.Left <= left And Line.Left + Line.Width >= left Then
                        Lab.Line = Line
                        Lab.LineID = Val(Line.Name)
                        Return Val(Line.Name)
                    End If
                End If
                If InStr(Line.Tag, "LineV") > 0 Then

                    If Math.Abs(Line.Left - left) < 10 And Line.Top <= top And Line.Top + Line.Height >= top Then
                        Lab.Line = Line
                        Lab.LineID = Val(Line.Name)
                        Return Val(Line.Name)
                    End If
                End If


            Next
        End With

        Return 0
    End Function

    Private Function Line_Check_OnLine(ByRef Lab As Lab_Type, sLine As Panel) As Long
        'LINE 上かのチェック
        '----------------------

        With Lab
            For Each tLine As Panel In .DesKTop(.DeskTopID).Controls
                If InStr(sLine.Tag, "LineH") > 0 Then

                    If InStr(tLine.Tag, "LineV") > 0 Then
                        If Math.Abs(tLine.Left - sLine.Left) < 10 And tLine.Top <= sLine.Top And tLine.Top + tLine.Height >= sLine.Top Then
                            Lab.Line = tLine
                            Lab.LineID = Val(tLine.Name)
                            Return Val(tLine.Name)
                        End If
                    End If
                End If
                If InStr(sLine.Tag, "LineV") > 0 Then
                    If InStr(tLine.Tag, "LineH") > 0 Then
                        If Math.Abs(tLine.Top - sLine.Top) < 10 And tLine.Left <= sLine.Left And tLine.Left + tLine.Width >= sLine.Left Then
                            Lab.Line = tLine
                            Lab.LineID = Val(tLine.Name)
                            Return Val(tLine.Name)
                        End If
                    End If
                End If

            Next
        End With

        Return 0
    End Function

    Private Function Line_Check_Parent(ByRef Key As String, ID As String) As Boolean
        'LINE 上かのチェック
        '----------------------
        Dim Kx() As String
        Kx = Split(Key, ".")
        If Kx Is Nothing Then Return False
        If Kx.Count < 2 Then Return False
        If Kx(1) = ID Then Return True
        Return False
    End Function


    Private Sub Line_Set_ByLID(ByRef Lab As Lab_Type, LID As Integer)
        'カレントLINE 設定
        '----------------------

        With Lab
            For Each Line As Panel In .DesKTop(.DeskTopID).Controls
                If Line.Name = LID.ToString Then

                    Lab.Line = Line
                    Lab.LineID = LID

                    Return

                End If
            Next
        End With
        Lab.Line = Nothing

    End Sub
    Private Sub Base_MouseUp(sender As Object, e As MouseEventArgs)
        '左クリックの場合

        Dim control As Control = CType(sender, Control)
        gLab.BoxID = Val(control.Name)
        If e.Button = MouseButtons.Left Then

            'コントロール取得

            Select Case gFlag_Pull
                Case "Move" '移動
                    'コントロールの位置を設定
                    control.Left = control.Left + e.X - gBase_StartX
                    control.Top = control.Top + e.Y - gBase_StartY
                    Me.Cursor = Cursors.Hand
                Case "Resize" '1 'リサイズ
                    Box_Re_Size(gLab, Val(control.Name), control.Width, control.Height)
            End Select

            gFlag_Pull = "Move"
        End If


        Dim tLID, sLID As Integer
        tLID = Base_Check_OnLine(gLab, control.Top, control.Left)  'LINE上か

        If InStr(control.Tag, "OnH") > 0 Then 'LINEから要素を外す場合

            sLID = Val(Replace(control.Tag, "OnH.", ""))
            Line_Set_ByLID(gLab, sLID) '対象ラインにフォーカス
            If gLab.Line Is Nothing Then Return

            control.Tag = "OnDESK"
            Box_Update(gLab, gLab.DeskID, gLab.BoxID, "LineID", "OnDESK") 'LINE要素の作り変え
            LineH_Refresh(gLab, 0) '要素を外す場合もとりあえず全展開
        End If

        If InStr(control.Tag, "OnV") > 0 Then
            sLID = Val(Replace(control.Tag, "OnV.", ""))
            Line_Set_ByLID(gLab, sLID) '対象ラインにフォーカス
            If gLab.Line Is Nothing Then Return

            control.Tag = "OnDESK"
            Box_Update(gLab, gLab.DeskID, gLab.BoxID, "LineID", "OnDESK") 'LINE要素の作り変え
            LineV_Refresh(gLab, 0) '要素を外す場合もとりあえず全展開
        End If


        If tLID <> 0 Then
            Line_Set_ByLID(gLab, tLID) '対象ラインにフォーカス
            If InStr(gLab.Line.Tag, "LineH") > 0 Then
                control.Tag = "OnH." & tLID.ToString
                control.Top = gLab.Line.Top + 10
                LineH_Refresh(gLab, 0) 'とりあえず全展開
            End If
            If InStr(gLab.Line.Tag, "LineV") > 0 Then
                control.Tag = "OnV." & tLID.ToString
                control.Left = gLab.Line.Left + 10
                LineV_Refresh(gLab, 0) 'とりあえず全展開
            End If

            Box_Update(gLab, gLab.DeskID, gLab.BoxID, "LineID", control.Tag)

            End If

            Box_Re_Move(gLab, Val(control.Name), control)


        Me.Cursor = Cursors.Default
    End Sub
    'Private Sub Base_MouseUp(sender As Object, e As MouseEventArgs)
    '    '左クリックの場合

    '    Dim control As Control = CType(sender, Control)
    '    gLab.BoxID = Val(control.Name)
    '    If e.Button = MouseButtons.Left Then

    '        'コントロール取得

    '        Select Case gFlag_Pull
    '            Case "Move" '移動
    '                'コントロールの位置を設定
    '                control.Left = control.Left + e.X - gBase_StartX
    '                control.Top = control.Top + e.Y - gBase_StartY
    '                Me.Cursor = Cursors.Hand
    '            Case "Resize" '1 'リサイズ
    '                Box_Re_Size(gLab, Val(control.Name), control.Width, control.Height)
    '        End Select

    '        gFlag_Pull = "Move"
    '        Log_Show(1, "BID[" & control.Name & "]" & control.Left & "|" & control.Top & "|" & control.Width & "|" & control.Height)

    '    End If


    '    Dim LID As Long
    '    LID = Base_Check_OnLine(gLab, control.Top, control.Left)  'LINE上か

    '    If LID <> 0 Then

    '        Select Case gLab.Line.Tag
    '            Case "LineH"
    '                control.Tag = "OnH." & LID.ToString
    '                control.Top = gLab.Line.Top + 10
    '                LineH_Refresh(gLab, 0) 'とりあえず全展開
    '            Case "LineV"
    '                control.Tag = "OnV." & LID.ToString
    '                control.Left = gLab.Line.Left + 10
    '                LineV_Refresh(gLab, 0) 'とりあえず全展開
    '        End Select
    '        Box_Update(gLab, gLab.DeskID, gLab.BoxID, "LineID", control.Tag)


    '    Else
    '        If InStr(control.Tag, "OnH") > 0 Then 'LINEから要素を外す場合

    '            LID = Val(Replace(control.Tag, "OnH.", ""))
    '            Line_Set_ByLID(gLab, LID)　'対象ラインにフォーカス
    '            If gLab.Line Is Nothing Then Return

    '            control.Tag = "OnDESK"
    '            Box_Update(gLab, gLab.DeskID, gLab.BoxID, "LineID", "OnDESK")　'LINE要素の作り変え
    '            LineH_Refresh(gLab, 0) '要素を外す場合もとりあえず全展開
    '        End If

    '        If InStr(control.Tag, "OnV") > 0 Then
    '            LID = Val(Replace(control.Tag, "OnV.", ""))
    '            Line_Set_ByLID(gLab, LID)　'対象ラインにフォーカス
    '            If gLab.Line Is Nothing Then Return

    '            control.Tag = "OnDESK"
    '            Box_Update(gLab, gLab.DeskID, gLab.BoxID, "LineID", "OnDESK")　'LINE要素の作り変え
    '            LineV_Refresh(gLab, 0) '要素を外す場合もとりあえず全展開
    '        End If

    '    End If

    '    Box_Re_Move(gLab, Val(control.Name), control)

    '    'リフレッシュ
    '    ' Me.Refresh()
    '    Me.Cursor = Cursors.Default
    'End Sub
    Private Sub LineH_Refresh(ByRef lab As Lab_Type, CloseID As Integer)
        Dim sl() As Integer
        Dim BoxS() As Panel
        With lab

            Dim cnt As Integer

            Dim Tag As String = .Line.Name

            For Each Box As Panel In .DesKTop(.DeskTopID).Controls
                If Line_Check_Parent(Box.Tag, Tag) = True Then
                    ReDim Preserve BoxS(cnt)
                    BoxS(cnt) = Box
                    ReDim Preserve sl(cnt)
                    sl(cnt) = Box.Left
                    cnt += 1
                End If
            Next

            If cnt = 0 Then Return
            Array.Sort(sl, BoxS)

            Dim LW As Integer
            LW = BoxS(0).Width
            BoxS(0).Left = lab.Line.Left + 2
            Box_Re_Move(gLab, Val(BoxS(0).Name), BoxS(0))

            For L = 1 To sl.Count - 1
                If L <= CloseID Then
                    BoxS(L).Left = BoxS(L - 1).Left + L
                    If LW < BoxS(L).Width + L Then
                        LW = BoxS(L).Width + L
                    End If
                Else
                    BoxS(L).Left = BoxS(L - 1).Left + BoxS(L - 1).Width
                    LW += BoxS(L).Width + 1

                End If
                'Box_Re_Move(gLab, Val(BoxS(L).Name), BoxS(L))
                If InStr(BoxS(L).Tag, "Line") > 0 Then
                    Line_Re_Move(gLab, Val(BoxS(L).Name), BoxS(L).Left, BoxS(L).Top)
                Else
                    Box_Re_Move(gLab, Val(BoxS(L).Name), BoxS(L))
                End If
                BoxS(L).BringToFront()
            Next
            lab.Line.Width = LW
            Line_Re_Size(gLab, Val(lab.Line.Name), LW, lab.Line.Height)
        End With

    End Sub
    Private Function LineH_Get_BoxS(ByRef lab As Lab_Type) As System.Collections.SortedList
        'LineH上のBOXを左から順にリスト化する
        Dim sl As New System.Collections.SortedList() 'SortedListを作成
        With lab
            Dim BoxS() As Panel
            Dim cnt As Integer

            Dim Tag As String = .Line.Name

            For Each Box As Panel In .DesKTop(.DeskTopID).Controls
                If Line_Check_Parent(Box.Tag, Tag) = True Then
                    ReDim Preserve BoxS(cnt)
                    BoxS(cnt) = Box
                    sl.Add(Box.Left, cnt)
                    cnt += 1
                End If
            Next
        End With
        Return sl
    End Function


    Private Sub LineV_MouseDoubleClick(sender As Object, e As MouseEventArgs)
        'LINEの開閉
        sender.BringToFront()

        With gLab
            .LineID = Val(sender.name)
            .Line = sender
            .Base = sender
            .Flg_BoxMouse = "OnV"
            If .LineNo > 0 Then
                .LineNo = 0
            Else
                .LineNo = .LineMax
            End If

            LineV_Refresh(gLab, .LineNo)
            Box_Update(gLab, .DeskID, .LineID, "LineNo", .LineNo.ToString)
        End With

    End Sub


    Private Sub BxPict_Drop(sender As Object, e As DragEventArgs)
        Dim files() As String = DirectCast(e.Data.GetData(DataFormats.FileDrop, False), String())
        '取得したファイルのパスを元にピクチャーボックスに画像を表示

        ''DBへ画像のパスを保存
        'sender.ImageLocation = files(0)
        'Bx_Re_Pict(gLab, sender.name, files(0))

        'DBへ画像のデータを保存
        Dim img As System.Drawing.Image = System.Drawing.Image.FromFile(files(0))
        sender.Image = img


        Bx_Re_Pict2(gLab, sender.name, files(0))

    End Sub

    Private Sub BxPict_DragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then

            ' ドラッグ中のファイルやディレクトリの取得
            Dim drags() As String =
          CType(e.Data.GetData(DataFormats.FileDrop), String())

            For Each d As String In drags
                If Not System.IO.File.Exists(d) Then
                    ' ファイル以外であればイベント・ハンドラを抜ける
                    Return
                End If
            Next

            e.Effect = DragDropEffects.Copy
        End If

    End Sub

    Private Sub WebView21_NavigationCompleted(sender As Object, e As CoreWebView2NavigationCompletedEventArgs)
        Dim mx As String
        mx = e.HttpStatusCode

        GGGX()

    End Sub



    Private Async Sub GGGX()
        'Dim html = Await WebView21.ExecuteScriptAsync("document.documentElement.outerHTML")
        'html = Regex.Unescape(html)
        'html = html.Remove(0, 1)
        'html = html.Remove(html.Length - 1, 1)
        'Debug.Print(html)

    End Sub




    Private Function Desk_Make_New(ByRef Lab As Lab_Type) As String
        'DESK 新設 &DB登録
        '-------------------
        Dim DID As Integer

        'Dim key, tmpD As String
        'Dim NewPage As Integer

        With Lab
            If .DesKTop Is Nothing Then
                .DeskTopID = 0
            Else
                If .DesKTop.Count = 1 Then
                    If .DesKTop(0) Is Nothing Then
                        .DeskTopID = 0
                    Else
                        .DeskTopID = .DesKTop.Count
                    End If
                Else
                    .DeskTopID = .DesKTop.Count
                End If

            End If

        End With

        With Lab

            If .DeskList Is Nothing Then
                .DeskList = New DataTable("Desk")
                Dim Ent() As String
                Ent = Split("Name,DeskID,DeskTopID,ProjKey,OwnerID,State,Type,Color,ShiftX,ShiftY,ZM,Path", ",")

                DID = 1
            Else

                DID = .DeskList.Rows.Count + 1
            End If
            .DeskID = DID
        End With

        Dim dtRow As DataRow
        With Lab.DeskList
            dtRow = .NewRow
            dtRow("DeskID") = DID.ToString
            dtRow("DeskTopID") = Lab.DeskTopID
            dtRow("Name") = "DESK"
            dtRow("ProjKey") = Lab.ProjKey
            dtRow("OwnerID") = gX_ownerID.ToString
            dtRow("Path") = gENV_Path & "\PRJ\" & Lab.ProjOwnerID.ToString & "\" & Lab.ProjID.ToString & "\" & DID.ToString & "\" & "Desk.txt"
            dtRow("State") = "Nomal"
            dtRow("Type") = ""
            dtRow("Color") = "AliceBlue"
            dtRow("ShiftX") = "0"
            dtRow("ShiftY") = "0"
            dtRow("ZM") = "0"
            .Rows.Add(dtRow)
        End With

        Desk_Save_Info(dtRow) 'ファイル書き込み

        Desk_Show(Lab, dtRow)
        Desk_Show_Name(Lab, dtRow)

    End Function
    Private Sub Desk_Del(ByRef Lab As Lab_Type)
        'DESKの削除
        '-------------------------------------------

        With DataGridView3
            If .CurrentCell Is Nothing Then Return '選択されたセルがない
            Lab.DeskTopID = Val(.CurrentCell.Tag)
        End With

        Dim Dlist() As DataRow
        Dim dtRow As DataRow
        With Lab
            .DeskID = .DesKTop(.DeskTopID).Tag
            Dlist = .DeskList.Select("DeskID ='" & .DeskID & "'")
            dtRow = Dlist(0)
            If dtRow("OwnerID") <> gX_ownerID.ToString Then
                MsgBox（"オーナーではありません"）
                Return
            End If

            dtRow("State") = "Del"
            Desk_Save_Info(dtRow) 'ファイル書き込み
        End With
        '全再描画で手抜きなの
        Proj_Set(Lab)


    End Sub
    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs)

    End Sub
    'Private Sub DB_CreateTBl(ByRef lab As Lab_Type)
    '    ' コネクション作成
    '    Dim con As SQLiteConnection
    '    con = New SQLiteConnection(GetConnectionString(lab.DbPath))
    '    con.Open()
    '    Using cmd = con.CreateCommand()
    '        ' テーブル作成SQL
    '        cmd.CommandText = "CREATE TABLE users (" &
    '                              "id INTEGER PRIMARY KEY," &
    '                              "name TEXT NOT NULL," &
    '                              "age INTEGER," &
    '                              "email TEXT NOT NULL UNIQUE" &
    '                              ")"
    '        cmd.ExecuteNonQuery()
    '    End Using
    '    con.Close()
    'End Sub


    Private Function GetConnectionString(ByRef Dbpath) As String
        ' 
        Dim builder As SQLiteConnectionStringBuilder = New SQLiteConnectionStringBuilder()
        builder.DataSource = Dbpath

        Return builder.ConnectionString
    End Function



    Private Sub ListView1_BeforeLabelEdit(sender As Object, e As LabelEditEventArgs) Handles ListView1.BeforeLabelEdit
        If gX_ownerID <> gLab.ProjOwnerID Then e.CancelEdit = True
    End Sub

    Private Sub ListView1_AfterLabelEdit(sender As Object, e As LabelEditEventArgs) Handles ListView1.AfterLabelEdit
        If e.Label = Nothing Then
            Return
        End If
        Proj_Rename(gLab, e.Label)

    End Sub

    Private Sub ListView3_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView3.MouseDoubleClick
        With ListView3
            If .SelectedItems.Count = 0 Then Return
            Element_Act_Item(gLab, .SelectedItems(0).Tag)
        End With
    End Sub

    'Private Sub TabControl3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl3.SelectedIndexChanged

    '    If sender.SelectedTab Is Nothing Then Return
    '    With gLab
    '        .DeskID = Val(sender.SelectedTab.name)
    '        .DeskTopID= sender.Selectedindex

    '        .DeskSx = .Desk(.DeskID).ShiftX
    '        .DeskSy = .Desk(.DeskID).ShiftY
    '        .DeskZm = .Desk(.DeskID).ZM
    '    End With




    '    Log_Show(1, "DESK --" & gLab.DeskID.ToString)

    'End Sub

    Private Sub TextBox1_MouseLeave(sender As Object, e As EventArgs) Handles TextBox1.MouseLeave

    End Sub

    Private Sub BxText_MouseLeave(sender As Object, e As EventArgs)
        Dim Tbox As TextBox
        Tbox = CType(sender, TextBox)
        System.IO.File.WriteAllText(Tbox.Tag, Tbox.Text)

    End Sub

    Private Sub BxDraw_MouseLeave(sender As Object, e As EventArgs)

        Dim DrID As Integer = Val(sender.Tag)

        Dim TEGAKI() As Byte
        TEGAKI = sender.Ink.Save()
        My.Computer.FileSystem.WriteAllBytes(gWORK_Path & "\DDX" & DrID.ToString, TEGAKI, True)
    End Sub

    Private Sub ContextMenuStrip2_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ContextMenuStrip2.ItemClicked
        Select Case e.ClickedItem.Tag
            Case "NEW"
                Desk_Make_New(gLab)
            Case "RENAME"

                Desk_Check_Rename(gLab)


            Case "DEL"
                Desk_Del(gLab)

            Case Else
        End Select
    End Sub


    Private Sub Desk_ReSize(sender As Object, e As EventArgs)
        'デスクリサイズ
        '--------------------
        Dim desk As TabPage
        desk = CType(sender, Control)
        For Each box As Panel In sender.Controls
            If box.Tag = "LINE" Then
                box.Width = desk.Width
                '   Line_ReSize(gLab, Val(box.Name), desk.Width)
            End If

        Next


    End Sub
    Private Sub ContextMenuStrip3_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ContextMenuStrip3.ItemClicked
        Select Case e.ClickedItem.Tag
            Case "COPY"
                MsgBox("いるかな？")
            Case "PASTE"
                MsgBox("いるかな？")
            Case "DEL"
                Box_Del(gLab)

            Case Else
        End Select
    End Sub
    Private Sub Proj_Del(ByRef Lab As Lab_Type)
        'Projの削除(DB修正)
        '-------------------------------------------
        Dim cmd As String
        If Lab.ProjOwnerID <> gX_ownerID Then
            MsgBox（"オーナーではありません"）
            Return
        End If

        Dim Plist() As DataRow
        Dim dtRow As DataRow
        With gLab
            Plist = .ProjList.Select("ProjKey ='" & .ProjKey & "'")

            dtRow = Plist(0)
            dtRow("State") = "Del"
            Proj_Save_Info(dtRow) 'ファイル書き込み
        End With
        For Each item As ListViewItem In ListView1.SelectedItems
            ListView1.Items.Remove(item)
        Next

    End Sub

    Private Sub Element_Del(ByRef Lab As Lab_Type)

    End Sub
    Private Sub Box_Del(ByRef Lab As Lab_Type)
        'BOX削除
        '------------------------

        '権限チェック

        'Dim cmd As String
        'If Lab.OwnerID <> gX_ownerID Then
        '    MsgBox（"オーナーでありません"）
        '    Return
        'End If

        gLab.Base.Visible = False
        Box_Update(Lab, Lab.DeskID, Lab.BoxID, "State", "Del")  'ステート変更



    End Sub
    Private Function Box_ZM(ByVal WD As Integer, Sx As Integer, SZ As Integer) As Integer
        'BOXのリサイズ
        '----------------------

        If Sx > 0 Then
            WD = WD - (WD * SZ) / 100
            If WD < 10 Then WD = 10
        End If
        If Sx < 0 Then
            WD = WD + (WD * SZ) / 100
        End If
        Return WD
    End Function


    Private Sub Box_UpDown(ByRef Base As Panel, SZ As Integer, Sx As Integer)
        'BOXのZ軸移動
        '----------------------
        ' Dim Wd, He As Integer
        With Base

            If Sx > 0 Then
                .SendToBack()
            End If
            If Sx < 0 Then
                .BringToFront()
            End If

        End With

    End Sub
    Private Function Box_ZM_Move(BoxP As Integer, Sp As Integer, Sx As Integer, Sz As Integer) As Integer
        'BOXのリサイズ
        '----------------------


        Select Case BoxP - Sp
            Case 0
                Return BoxP
            Case > 0
                Return Sp + Box_ZM(Math.Abs(BoxP - Sp), Sx, Sz)
            Case < 0
                Return Sp - Box_ZM(Math.Abs(BoxP - Sp), Sx, Sz)
        End Select

    End Function

    'Private Sub Box_MoveY(ByRef Base As Panel, Dx As Integer, Sx As Integer, ex As Integer)
    '    'BOXのリサイズ
    '    '----------------------
    '    Dim Y, wx As Integer
    '    With Base
    '        Y = .Top
    '        wx = Dx - Y
    '        If Math.Abs(wx) < 10 Then Return

    '        If ex > 0 Then
    '            Y = Y + (wx * Sx) / 100
    '        End If
    '        If ex < 0 Then
    '            Y = Y - (wx * Sx) / 100
    '        End If
    '        .Top = Y
    '    End With

    'End Sub

    Private Sub Form1_MouseWheel(sender As Object, e As MouseEventArgs) Handles Me.MouseWheel
        Dim Wd, He As Integer
        Select Case gLab.Flg_BoxMouse
            Case "OnBOX"
                '  Box_ZM(gLab.Box, 10, e.Delta)
                Box_UpDown(gLab.Base, 10, e.Delta)
            Case "OnFont"

                Select Case e.Delta
                    Case > 0
                        gFontSize += 1
                    Case = 0
                        Return
                    Case < 0
                        gFontSize -= 1
                        If gFontSize < 8 Then gFontSize = 8
                End Select
                FontView_Show(gFontName, gFontSize)

            Case "OnDESK", ""
                'If (Control.MouseButtons And MouseButtons.Right) = MouseButtons.Right Then
                '    With TabControl3
                '        For Each box As Panel In .TabPages(.SelectedIndex).Controls
                '            Box_ZM_Move(box, 10, e.Delta)
                '        Next
                '    End With
                '    Return
                'End If
                '拡大縮小
                With gLab
                    If (Control.MouseButtons And MouseButtons.Left) = MouseButtons.Left Then
                        If e.Delta > 0 Then Desk_Re_Size(gLab, .DeskID, 1)
                        If e.Delta < 0 Then Desk_Re_Size(gLab, .DeskID, -1)
                    End If
                End With


                '移動
                'Dim Ax, BX, BY As Integer
                'If e.Delta = 0 Then Return
                'If e.Delta < 0 Then Ax = -1
                'If e.Delta > 0 Then Ax = 1
                'If gLab.MvX = 0 Then
                '    If gLab.MvY = 0 Then Return
                '    BX = 0
                '    BY = 20 * Ax
                'Else

                '    BX = 20 * Ax
                '    BY = 20 * Ax * gLab.MvY / gLab.MvX
                'End If
                '.Left = .Left + BX
                '.Top = .Top + BY



            Case "OnH"
                'サクッとスライド
                'If (Control.MouseButtons And MouseButtons.Left) = MouseButtons.Left Then
                'End If
                Dim Ax As Integer
                If e.Delta = 0 Then Return
                If e.Delta < 0 Then Ax = -1
                If e.Delta > 0 Then Ax = 1

                With gLab
                    If .LineID = 0 Then Return
                    Select Case Ax
                        Case -1
                            If .LineNo < 0 Then Return
                            .LineNo = .LineNo - 1
                        Case 1
                            If .LineNo > .LineMax Then Return
                            .LineNo = .LineNo + 1
                    End Select
                    LineH_Refresh(gLab, .LineNo)
                    Box_Update(gLab, .DeskID, .LineID, "LineNo", .LineNo.ToString)
                End With
            Case "OnV"
                'サクッとスライド
                'If (Control.MouseButtons And MouseButtons.Left) = MouseButtons.Left Then
                'End If
                Dim Ax As Integer
                If e.Delta = 0 Then Return
                If e.Delta < 0 Then Ax = -1
                If e.Delta > 0 Then Ax = 1

                With gLab
                    If .LineID = 0 Then Return
                    Select Case Ax
                        Case -1
                            If .LineNo < 0 Then Return
                            .LineNo = .LineNo - 1
                        Case 1
                            If .LineNo > .LineMax Then Return
                            .LineNo = .LineNo + 1
                    End Select
                    LineV_Refresh(gLab, .LineNo)
                    '  Box_Update(gLab, .DeskID, .LineID, "LineNo", .LineNo.ToString)
                End With
        End Select


    End Sub
    Private Sub Desk_Re_Size(ByRef lab As Lab_Type, ByRef DeskID As Integer, Flg As Integer)
        Dim Sx As Integer '    拡大率
        Sx = 10
        Dim Key(), Dat() As String
        ReDim Key(1), Dat(1)

        With lab.Desk(DeskID)
            .ZM = .ZM + Flg
            Key(0) = "ZM" : Dat(0) = .ZM.ToString
            DeskList_Update(lab, DeskID, Key, Dat)

        End With

        Dim sp As System.Drawing.Point = System.Windows.Forms.Cursor.Position
        Dim cp As System.Drawing.Point = lab.DesKTop(lab.DeskTopID).PointToClient(sp)
        ' Dim cp As System.Drawing.Point = Me.PointToClient(sp) '画面座標をクライアント座標に変換する
        Dim Dx As Integer = cp.X
        Dim Dy As Integer = cp.Y

        For Each box As Panel In lab.DesKTop(lab.DeskTopID).Controls

            Select Case box.Tag
                Case "LineH"

                    box.Left = Box_ZM_Move(box.Left, Dx, Flg, Sx)
                    box.Top = Box_ZM_Move(box.Top, Dy, Flg, Sx)
                Case "LineV"
                    box.Left = Box_ZM_Move(box.Left, Dx, Flg, Sx)
                    box.Top = Box_ZM_Move(box.Top, Dy, Flg, Sx)
                Case "OnDESK"
                    box.Width = Box_ZM(box.Width, Flg, Sx)
                    box.Height = Box_ZM(box.Height, Flg, Sx)


                    box.Left = Box_ZM_Move(box.Left, Dx, Flg, Sx)
                    box.Top = Box_ZM_Move(box.Top, Dy, Flg, Sx)
                Case Else
                    If InStr(box.Tag, "OnH") > 0 Then
                        box.Left = Box_ZM_Move(box.Left, Dx, Flg, Sx)
                        box.Top = Box_ZM_Move(box.Top, Dy, Flg, Sx)
                    End If
                    If InStr(box.Tag, "OnV") > 0 Then
                        box.Left = Box_ZM_Move(box.Left, Dx, Flg, Sx)
                        box.Top = Box_ZM_Move(box.Top, Dy, Flg, Sx)
                    End If

                    ' Box_Re_Size(gLab, Val(box.Name), box.Width, box.Height)
            End Select

        Next
        ''End With

        'Return

    End Sub
    Private Sub BxPic_DoubleClick(sender As Object, e As EventArgs)
        If sender.tag = "" Then Return

        If InStr(sender.tag, "★"） > 0 Then Return
        System.Diagnostics.Process.Start(sender.tag)


    End Sub

    Private Sub Form1_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        gLab.Flg_BoxMouse = ""
    End Sub

    Private Sub Form1_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        gLab.Flg_BoxMouse = "OnDESK"
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ListView1_HandleCreated(sender As Object, e As EventArgs) Handles ListView1.HandleCreated

    End Sub



    Private Sub TabControl1_Resize(sender As Object, e As EventArgs)

    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress

    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Dim cp As Point


        If e.KeyData = (Keys.Control Or Keys.V) Then


            With System.Windows.Forms.Cursor.Position
                cp = New Point(.X, .Y)
            End With
            cp = PointToClient(cp)
            cp.X = cp.X - SplitContainer1.Panel1.Left
            cp.Y = cp.Y - SplitContainer1.Panel1.Top
            Desk_Paste(cp)
        End If

    End Sub

    Private Sub TabPage11_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TabPage11_KeyDown(sender As Object, e As KeyEventArgs)

    End Sub
    Private Sub Desk_KeyDown(sender As Object, e As KeyEventArgs)

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click

        Proj_Sync(gLab)




    End Sub

    Sub Proj_Sync(ByRef lab As Lab_Type)
        Try


            Dim process As New Process()
            Dim output As String
            With process
                With .StartInfo
                    .FileName = "C:\Program Files\Git\bin\bash.exe"
                    .Arguments = "-c ""cd ENV;git gc; bash"""
                    .UseShellExecute = True
                    ' .RedirectStandardOutput = True
                    .CreateNoWindow = False
                End With
                .Start()

                'output = .StandardOutput.ReadToEnd()
                .WaitForExit()

                With .StartInfo
                    .FileName = "C:\Program Files\Git\bin\bash.exe"
                    .Arguments = "-c ""cd ENV; git add .; git commit -m ""comment""; bash"""
                    .UseShellExecute = True
                    ' .RedirectStandardOutput = True
                    .CreateNoWindow = False
                End With
                .Start()
                ' output = .StandardOutput.ReadToEnd()
                .WaitForExit()

                With .StartInfo
                    .FileName = "C:\Program Files\Git\bin\bash.exe"
                    .Arguments = "-c ""cd ENV; git pull --rebase origin main; bash"""
                    .UseShellExecute = True
                    ' .RedirectStandardOutput = True
                    .CreateNoWindow = False
                End With

                .Start()
                'output = .StandardOutput.ReadToEnd()
                .WaitForExit()
            End With


            ' Log_Show(2, output)
            Return

        Catch ex As Exception

            MsgBox(ex.Message)

        End Try


        'Console.WriteLine(output)
    End Sub
    'Private Sub GitCommand(command As String, Path As String)
    '    Try
    '        Dim Git_log As String = gLOG_Path & "\GitCmd.log"
    '        Dim Git_Log2 = Replace(Git_log, "\", "/")
    '        System.IO.File.AppendAllText(Git_log, "----------------" & vbCrLf)
    '        System.IO.File.AppendAllText(Git_log, System.DateTime.Now & vbCrLf)

    '        System.IO.File.AppendAllText(Git_log, command & vbCrLf)
    '        System.IO.File.AppendAllText(Git_log, "----------------" & vbCrLf)

    '        Dim process As New Process()

    '        With process

    '            With .StartInfo
    '                .FileName = "C:\Program Files\Git\bin\bash.exe"
    '                .Arguments = "-c ""cd " & Path & "&>>" & Git_Log2 & ";" & command & " &>>" & Git_Log2 & """"
    '                .UseShellExecute = True
    '                .RedirectStandardOutput = False
    '                .CreateNoWindow = True
    '            End With
    '            .Start()
    '            ' output = .StandardOutput.ReadToEnd()
    '            .WaitForExit()

    '            TextBox2.Text &= System.IO.File.ReadAllText(Git_log)
    '        End With
    '        Return
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub

    Private Sub GitCommand(GitActSh As String)
        Try
            Dim Git_log As String = gLOG_Path & "\GitCmd.log"
            Dim Git_Log2 = Replace(Git_log, "\", "/")

            System.IO.File.AppendAllText(Git_log, "----------------" & vbCrLf)
            System.IO.File.AppendAllText(Git_log, System.DateTime.Now & vbCrLf)
            System.IO.File.AppendAllText(Git_log, Command() & vbCrLf)
            System.IO.File.AppendAllText(Git_log, "----------------" & vbCrLf)

            Dim process As New Process()
            Dim Output As String
            Dim Errout As String

            With process
                With .StartInfo
                    .FileName = "C:\Program Files\Git\bin\bash.exe"
                    .Arguments = GitActSh
                    .RedirectStandardOutput = True
                    .RedirectStandardError = True
                    .UseShellExecute = False
                End With
                .Start()
                output = process.StandardOutput.ReadToEnd()
                Errout = process.StandardError.ReadToEnd()
                process.WaitForExit()
                .WaitForExit()
            End With
            'MsgBox(output)

            System.IO.File.AppendAllText(Git_log, Output)
            System.IO.File.AppendAllText(Git_log, Errout)
            TextBox2.Text &= System.IO.File.ReadAllText(Git_log)

            Return
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Function Git_Make_Shell(ByRef lab As Lab_Type, ShFile As String) As String
        Dim InFile As String = gENV_Path & "\" & ShFile
        Dim OutFile As String = gENV_Path & "\Act_" & ShFile

        Dim GitProjName As String = GetHashValue(lab.ProjName)
        Dim ProjPath As String = Replace(lab.ProjPath, "\", "/")
        Dim GitPath As String = lab.ProjPath & "\.git"

        Dim TmpSh = System.IO.File.ReadAllText(InFile)
        TmpSh = Replace(TmpSh, "{{username}}", gX_owner)
        TmpSh = Replace(TmpSh, "{{reponame}}", GitProjName)
        '  TmpSh = Replace(TmpSh, "{{token}}", "glpat-ZT6mMixsGerz_2SMac-y")
        TmpSh = Replace(TmpSh, "{{token}}", gX_Token)
        TmpSh = Replace(TmpSh, "{{projpath}}", ProjPath)

        System.IO.File.WriteAllText(OutFile, TmpSh)

        Return OutFile

    End Function


    Private Sub Proj_Cast(lab As Lab_Type)
        Try
            If lab.ProjPath = "" Then
                MsgBox("プロジェクトを選択してください。")
                Return
            End If
            TextBox2.Text = "" 'log_view  初期化

            Dim GitPath As String = lab.ProjPath & "\.git"


            '    GitCommand("git --version", ProjPath)
            If System.IO.Directory.Exists(GitPath) = False Then
                GitCommand(Git_Make_Shell(lab, "Init.sh"))
            End If
            GitCommand(Git_Make_Shell(lab, "Cast.sh"))
            '    Return
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Public Shared Function GetHashValue(str1 As String) As String

        Dim ar1() As Byte = Encoding.UTF8.GetBytes(str1)
        Dim ar2() As Byte =
        New MD5CryptoServiceProvider().ComputeHash(ar1)

        Dim str2 As StringBuilder = New StringBuilder()
        For Each b1 As Byte In ar2
            str2.Append(b1.ToString("x2"))
        Next
        Return str2.ToString
    End Function
    'Private Sub Proj_Save(dtRow As DataRow)
    '    'Proj_Save
    '    '----------------

    '    Proj_Save_Info(gLab, ProjPath)
    '    'Proj_Save_DeskAll(lab, ProjPath)
    '    'Proj_Save_BoxAll(lab, ProjPath)

    'End Sub


    Private Sub Proj_Save_Info(dtRow As DataRow)
        'Proj_Save
        '----------------
        Dim Ent() As String
        Ent = Split("ProjID|Title|OwnerID|State|Summary|Hypothesis|Purpose|Problem", "|")

        Dir_Check(System.IO.Path.GetDirectoryName(dtRow("Path")))
        Dim Dat As String

        Dat = ""
        For L = 0 To Ent.Count - 1
            Dat &= Ent(L) & vbTab & dtRow(Ent(L)) & vbCrLf
        Next
        System.IO.File.WriteAllText(dtRow("Path"), Dat)

    End Sub
    Private Sub Desk_Save_Info(dtRow As DataRow)
        'Proj_Save
        '----------------
        Dim Ent() As String
        Ent = Split("Name,OwnerID,State,Color,ShiftX,ShiftY,ZM", ",")
        Dir_Check(System.IO.Path.GetDirectoryName(dtRow("Path")))
        Dim Dat As String

        Dat = ""
        For L = 0 To Ent.Count - 1
            Dat &= Ent(L) & vbTab & dtRow(Ent(L)) & vbCrLf
        Next
        System.IO.File.WriteAllText(dtRow("Path"), Dat)

    End Sub

    Private Sub Box_Save_Info(dtRow As DataRow)
        'Proj_Save
        '----------------
        Dim Ent() As String
        Ent = Split("OwnerID,State,Type,ParentID,LineID,LineNo,BluePrint", ",")
        Dir_Check(System.IO.Path.GetDirectoryName(dtRow("Path")))
        Dim Dat As String

        Dat = ""
        For L = 0 To Ent.Count - 1
            Dat &= Ent(L) & vbTab & dtRow(Ent(L)) & vbCrLf
        Next
        System.IO.File.WriteAllText(dtRow("Path"), Dat)

    End Sub


    'Private Sub Proj_Save_BoxAll(lab As Lab_Type, ProjPath As String)
    '    'Proj_Save
    '    '----------------
    '    Dim BoxPath As String
    '    With lab
    '        For L = 0 To .Desk.Count - 1
    '            With .Desk(L)

    '                For K = 0 To .Box.Count - 1
    '                    BoxPath = ProjPath & "\" & .DeskID.ToString & .Box(K).BoxID.ToString
    '                    Proj_Save_Box(lab, BoxPath, K, L)
    '                Next

    '            End With
    '        Next
    '    End With

    'End Sub



    Private Sub Proj_Save_Box(ByRef lab As Lab_Type, ByVal BoxPath As String, BID As Integer, DID As Integer)
        'Proj_Save
        '----------------
        Dim Dat As String
        Dim DepPath As String

        Dir_Check(BoxPath)

        With lab.Desk(BID).Box(DID)
            If InStr(.State, "Depend") > 0 Then
                DepPath = BoxPath & "\Dep"
                Proj_Save_Depend(lab, DepPath, .Depend)
                .State = Replace(.State, "Depend", "")
            End If
            Dat = ""
            Dat &= "ID" & vbTab & .ID.ToString & vbCr

            Dat &= "OwnerID" & vbTab & lab.OwnerID.ToString & vbCr
            Dat &= "State" & vbTab & .State & vbCr
            Dat &= "Type" & vbTab & .Type.ToString & vbCr

            Dat &= "ParentID" & vbTab & .ParentID.ToString & vbCr
            Dat &= "LineID" & vbTab & .LineID.ToString & vbCr
            Dat &= "LineNo" & vbTab & .LineNo.ToString & vbCr
            Dat &= "BluePrint" & vbTab & .BluePrint & vbCr
            'Dat &="Base" & vbTab & Panel

            ''位置サイズ情報

            Dat &= "X" & vbTab & .Base.Left.ToString & vbCr
            Dat &= "Y" & vbTab & .Base.Top.ToString & vbCr
            Dat &= "W" & vbTab & .Base.Width.ToString & vbCr
            Dat &= "H" & vbTab & .Base.Height.ToString & vbCr
            'Dat &="LV" & vbTab & Integer

            'Dat &="PathA()" & vbTab & String
            'Dat &="PathB()" & vbTab & String

            'System.IO.File.WriteAllText(BoxPath & "\" & .ID.ToString
        End With



    End Sub

    Private Sub Proj_Save_Depend(ByRef lab As Lab_Type, ByVal DepPath As String, Depend() As Depend_Type)
        'Proj_Save
        '----------------
        Dim Deps(3) As String
        Dim Dat As String


        Dir_Check(DepPath)

        For L = 0 To Depend.Count - 1
            With Depend(L)
                Deps(0) = DepPath & "\" & L.ToString
                Deps(1) = Deps(0) & "\C"
                Deps(2) = Deps(0) & "\V"
                For K = 0 To 3
                    Dir_Check(Deps(K))
                Next

                Dat = ""

                Dat &= "Type" & vbTab & .Type & vbCr
                Dat &= "State" & vbTab & .State & vbCr
                Dat &= "Name" & vbTab & .Name & vbCr
                Dat &= "PathO" & vbTab & .PathO & vbCr


                System.IO.File.WriteAllText(Deps(0) & "\Dep.txt", Dat)
            End With
        Next


    End Sub

    'Private Sub Proj_Load(lab As Lab_Type)
    '    'Proj_Load
    '    '----------------
    '    Dim ProjPath As String
    '    ProjPath = gApp_Path & "\ENV\PRJ\" & lab.ProjID.ToString

    '    Proj_Load_Info(lab, ProjPath)
    '    Proj_Load_DeskAll(lab, ProjPath)
    '    Proj_Load_BoxAll(lab, ProjPath)

    'End Sub


    'Private Sub Proj_Load_Info(lab As Lab_Type, ProjPath As String)
    '    'Proj_Load
    '    '----------------
    '    If System.IO.Directory.Exists(ProjPath) = False Then System.IO.Directory.CreateDirectory(ProjPath)

    '    Dim Dat As String
    '    With lab
    '        Dat = ""
    '        Dat &= "ID" & vbTab & .ProjID & vbCr
    '        Dat &= "Name" & vbTab & .ProjName & vbCr
    '        Dat &= "Owner" & vbTab & .Owner & vbCr
    '        Dat &= "OwnerID" & vbTab & .OwnerID.ToString & vbCr

    '    End With

    '    System.IO.File.WriteAllText(ProjPath & "\Proj.info", Dat)

    'End Sub

    'Private Sub Proj_Load_DeskAll(lab As Lab_Type, ProjPath As String)
    '    'Proj_Load
    '    '----------------
    '    With lab
    '        For L = 0 To .Desk.Count - 1
    '            With .Desk(L)
    '                Proj_Load_Desk(lab, ProjPath & "\Desk\" & .DeskID.ToString, L)
    '            End With
    '        Next
    '    End With

    'End Sub
    'Private Sub Proj_Load_Desk(ByRef lab As Lab_Type, ByVal DeskPath As String, Page As Integer)
    '    'Proj_Load
    '    '----------------
    '    Dim Dat As String

    '    If System.IO.Directory.Exists(DeskPath) = False Then System.IO.Directory.CreateDirectory(DeskPath)

    '    With lab.Desk(Page)
    '        Dat = ""
    '        Dat &= "ID" & vbTab & .DeskID.ToString & vbCr
    '        Dat &= "ProjID" & vbTab & lab.ProjID.ToString & vbCr
    '        Dat &= "State" & vbTab & .State & vbCr
    '        Dat &= "Type" & vbTab & .Type & vbCr
    '        Dat &= "Name" & vbTab & .Name & vbCr
    '        Dat &= "Color" & vbTab & .Color & vbCr
    '        Dat &= "ShiftX" & vbTab & .ShiftX.ToString & vbCr
    '        Dat &= "ShiftY" & vbTab & .ShiftY.ToString & vbCr
    '        Dat &= "ZM" & vbTab & .ZM.ToString & vbCr
    '        Dat &= "BMax" & vbTab & .BMax.ToString & vbCr

    '    End With

    '    System.IO.File.WriteAllText(DeskPath & "\Desk.info", Dat)

    'End Sub

    'Private Sub Proj_Load_BoxAll(lab As Lab_Type, ProjPath As String)
    '    'Proj_Load
    '    '----------------
    '    Dim BoxPath As String
    '    With lab
    '        For L = 0 To .Desk.Count - 1
    '            With .Desk(L)
    '                BoxPath = ProjPath & "\Desk\" & .DeskID.ToString & "\Box"
    '                For K = 0 To .Box.Count - 1
    '                    Proj_Load_Box(lab, BoxPath, K, L)
    '                Next

    '            End With
    '        Next
    '    End With

    'End Sub



    'Private Sub Proj_Load_Box(ByRef lab As Lab_Type, ByVal BoxPath As String, BID As Integer, DID As Integer)
    '    'Proj_Load
    '    '----------------
    '    Dim Dat As String
    '    Dim DepPath As String
    '    If System.IO.Directory.Exists(BoxPath) = False Then System.IO.Directory.CreateDirectory(BoxPath)

    '    With lab.Desk(BID).Box(DID)
    '        If InStr(.State, "Depend") > 0 Then
    '            DepPath = BoxPath & "\Dep"
    '            Proj_Load_Depend(lab, DepPath, .Depend)
    '            .State = Replace(.State, "Depend", "")
    '        End If
    '        Dat = ""
    '        Dat &= "ID" & vbTab & .ID.ToString & vbCr
    '        Dat &= "DeskID" & vbTab & lab.DeskID.ToString & vbCr
    '        Dat &= "ProjID" & vbTab & lab.ProjID.ToString & vbCr
    '        Dat &= "OwnerID" & vbTab & lab.OwnerID.ToString & vbCr
    '        Dat &= "State" & vbTab & .State & vbCr
    '        Dat &= "Type" & vbTab & .Type.ToString & vbCr

    '        Dat &= "ParentID" & vbTab & .ParentID.ToString & vbCr
    '        Dat &= "LineID" & vbTab & .LineID.ToString & vbCr
    '        Dat &= "LineNo" & vbTab & .LineNo.ToString & vbCr
    '        Dat &= "BluePrint" & vbTab & .BluePrint & vbCr
    '        'Dat &="Base" & vbTab & Panel

    '        ''位置サイズ情報

    '        Dat &= "X" & vbTab & .Base.Left.ToString & vbCr
    '        Dat &= "Y" & vbTab & .Base.Top.ToString & vbCr
    '        Dat &= "W" & vbTab & .Base.Width.ToString & vbCr
    '        Dat &= "H" & vbTab & .Base.Height.ToString & vbCr
    '        'Dat &="LV" & vbTab & Integer

    '        'Dat &="PathA()" & vbTab & String
    '        'Dat &="PathB()" & vbTab & String

    '        System.IO.File.WriteAllText(BoxPath & "\" & .ID.ToString, Dat)
    '    End With



    'End Sub

    'Private Sub Proj_Load_Depend(ByRef lab As Lab_Type, ByVal DepPath As String, Depend() As Depend_Type)
    '    'Proj_Load
    '    '----------------
    '    Dim Deps(3) As String
    '    Dim Dat As String

    '    If System.IO.Directory.Exists(DepPath) = False Then System.IO.Directory.CreateDirectory(DepPath)

    '    For L = 0 To Depend.Count - 1
    '        With Depend(L)
    '            Deps(0) = DepPath & "\" & L.ToString
    '            Deps(1) = Deps(0) & "\C"
    '            Deps(2) = Deps(0) & "\V"
    '            For K = 0 To 3
    '                If System.IO.Directory.Exists(Deps(K)) = False Then System.IO.Directory.CreateDirectory(Deps(K))
    '            Next

    '            Dat = ""
    '            Dat &= "ID" & vbTab & L.ToString & vbCr
    '            Dat &= "Type" & vbTab & .Type & vbCr
    '            Dat &= "State" & vbTab & .State & vbCr
    '            Dat &= "Name" & vbTab & .Name & vbCr
    '            Dat &= "PathO" & vbTab & .PathO & vbCr
    '            Dat &= "PathC " & vbTab & Deps(1)
    '            Dat &= "PathV" & vbTab & Deps(2)

    '            System.IO.File.WriteAllText(Deps(0) & "\Info.txt", Dat)
    '        End With
    '    Next


    'End Sub

    Private Sub TextBox6_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress

    End Sub

    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyData = Keys.Return Then
            Proj_Search(gLab)
        End If
    End Sub

    Private Sub Button14_Click_1(sender As Object, e As EventArgs) Handles Button14.Click
        Proj_Cast(gLab)
    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click
        With gLab
            If DataGridView3.ColumnCount - .DeskIDShift > 0 Then Return
            .DeskIDShift = .DeskIDShift + 1

        End With

        ' Desk_Show_Name(gLab)
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        With gLab
            If .DeskIDShift = 0 Then Return
            .DeskIDShift = .DeskIDShift - 1
        End With
        ' Desk_Show_Name(gLab)
    End Sub
    Private Sub Desk_Show_Name(ByRef Lab As Lab_Type, ByRef Row As DataRow)

        '  Ent = Split("Name,DeskID,ProjKey,OwnerID,State,Type,Color,ShiftX,ShiftY,ZM,Path", ",")
        Dim ViewID As Integer


        ViewID = Lab.DeskTopID - Lab.DeskIDShift


        With DataGridView3
            If ViewID >= 0 And ViewID < .ColumnCount Then
                .Rows(0).Cells(ViewID).Value = Row("Name")
                .Rows(0).Cells(ViewID).Tag = Lab.DeskTopID.ToString

            End If
        End With

    End Sub
    Private Sub Desk_Clear_Name(ByRef Lab As Lab_Type)
        Lab.DeskIDShift = 0
        With DataGridView3.Rows(0)

            For Each N In .Cells
                N.Value = ""
                N.Tag = ""
            Next
        End With
    End Sub

    Private Sub Desk_Clear_Desk(ByRef Lab As Lab_Type)

        With Lab
            .DeskTopID = 0
            If .DesKTop Is Nothing Then Return
            For Each N In .DesKTop
                If N IsNot Nothing Then
                    N.Visible = False
                    N.Dispose()
                End If

            Next
            ReDim .DesKTop(0)


        End With
    End Sub
    'Private Sub Desk_Clear(ByRef Lab As Lab_Type)
    '    If SplitContainer1.Panel1.Controls Is Nothing Then Return
    '    For Each Desk As Object In SplitContainer1.Panel1.Controls
    '        If Desk.GetType().Name = NameOf(Panel) Then
    '            Desk = CType(Desk, Panel)
    '            Desk.visible = False
    '            Desk.Dispose()
    '        End If
    '    Next


    'End Sub


    Private Sub DataGridView3_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView3.CellMouseDown


    End Sub

    Private Sub DataGridView3_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView3.SelectionChanged
        If sender.SelectedCells.count = 0 Then Return

        With DataGridView3.CurrentCell
            Log_Show(1, "グリッド名＝" & .Value & "|グリッドのTAG(DeskTopID)" & .Tag)
            If gLab.DesKTop IsNot Nothing Then
                Log_Show(2, "deskTop(DeskTopID).NAME＝" & gLab.DesKTop(Val(.Tag)).Name & "|deskTop(DeskTopID).Tag＝" & gLab.DesKTop(Val(.Tag)).Tag)

                gLab.DeskID = Val(gLab.DesKTop(Val(.Tag)).Tag)
            End If
        End With


        For Each c In sender.SelectedCells
            ' Log_Show(1, c.tag)
            If c.tag IsNot Nothing Then
                ' gLab.DeskTopID = Val(c.tag)
                gLab.DeskTopID = Val(c.tag)
                Exit For
            End If
        Next

        With gLab
            If .DesKTop Is Nothing Then Return

            .DesKTop(.DeskTopID).Visible = True
            .DesKTop(.DeskTopID).BringToFront()


        End With




    End Sub

    Private Sub ListView1_MouseClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseClick

    End Sub

    Private Sub DataGridView3_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellValueChanged
        If gLab.Flg_DeskRename <> True Then Return
        gLab.Flg_DeskRename = False
        With DataGridView3
            If .CurrentCell Is Nothing Then Return '選択されたセルがない
            With .CurrentCell

                If .Value = "" Then
                    ' .RejectChanges
                    Return
                End If

                If Desk_Rename(gLab, gLab.DesKTop(.Tag).Tag, .Value) <> True Then
                    '  .RejectChanges
                    Return

                End If
            End With

        End With

    End Sub


    Private Sub DataGridView3_CellLeave(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellLeave
        gLab.Flg_DeskRename = False
    End Sub

    Private Sub DataGridView3_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView3.CellMouseDoubleClick

        Desk_Check_Rename(gLab)

    End Sub

    Private Sub DataGridView3_MouseDown(sender As Object, e As MouseEventArgs) Handles DataGridView3.MouseDown

        'With gLab
        '    .DeskID = .DesKTop(Val(sender.tag)).Name
        'End With



    End Sub


End Class