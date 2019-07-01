<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.idfind_txt = New System.Windows.Forms.TextBox()
        Me.Search_Lst = New System.Windows.Forms.ListView()
        Me.id = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.en_name = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.kr_name = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.add_bookmark = New System.Windows.Forms.ToolStripMenuItem()
        Me.remove_bookmark = New System.Windows.Forms.ToolStripMenuItem()
        Me.copy_id = New System.Windows.Forms.ToolStripMenuItem()
        Me.copy_en = New System.Windows.Forms.ToolStripMenuItem()
        Me.copy_kr = New System.Windows.Forms.ToolStripMenuItem()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Bookmark_Btn = New System.Windows.Forms.Button()
        Me.Search_Btn = New System.Windows.Forms.Button()
        Me.offbtn = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ContextMenuStrip1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'idfind_txt
        '
        Me.idfind_txt.Location = New System.Drawing.Point(13, 68)
        Me.idfind_txt.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.idfind_txt.Name = "idfind_txt"
        Me.idfind_txt.Size = New System.Drawing.Size(400, 23)
        Me.idfind_txt.TabIndex = 5
        Me.idfind_txt.Text = "이곳에 검색할 아이템의 ID/이름을 입력하세요."
        '
        'Search_Lst
        '
        Me.Search_Lst.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.id, Me.en_name, Me.kr_name})
        Me.Search_Lst.ContextMenuStrip = Me.ContextMenuStrip1
        Me.Search_Lst.FullRowSelect = True
        Me.Search_Lst.GridLines = True
        Me.Search_Lst.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.Search_Lst.Location = New System.Drawing.Point(14, 143)
        Me.Search_Lst.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Search_Lst.MultiSelect = False
        Me.Search_Lst.Name = "Search_Lst"
        Me.Search_Lst.Size = New System.Drawing.Size(430, 222)
        Me.Search_Lst.TabIndex = 12
        Me.Search_Lst.UseCompatibleStateImageBehavior = False
        Me.Search_Lst.View = System.Windows.Forms.View.Details
        '
        'id
        '
        Me.id.Text = "ID"
        Me.id.Width = 132
        '
        'en_name
        '
        Me.en_name.Text = "영문 명칭"
        Me.en_name.Width = 150
        '
        'kr_name
        '
        Me.kr_name.Text = "한글 명칭"
        Me.kr_name.Width = 130
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.add_bookmark, Me.remove_bookmark, Me.copy_id, Me.copy_en, Me.copy_kr})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(255, 114)
        '
        'add_bookmark
        '
        Me.add_bookmark.Name = "add_bookmark"
        Me.add_bookmark.Size = New System.Drawing.Size(254, 22)
        Me.add_bookmark.Text = "해당 아이템을 즐겨찾기에 추가"
        '
        'remove_bookmark
        '
        Me.remove_bookmark.Name = "remove_bookmark"
        Me.remove_bookmark.Size = New System.Drawing.Size(254, 22)
        Me.remove_bookmark.Text = "해당 아이템을 즐겨찾기에서 제거"
        '
        'copy_id
        '
        Me.copy_id.Name = "copy_id"
        Me.copy_id.Size = New System.Drawing.Size(254, 22)
        Me.copy_id.Text = "해당 아이템 ID 복사"
        '
        'copy_en
        '
        Me.copy_en.Name = "copy_en"
        Me.copy_en.Size = New System.Drawing.Size(254, 22)
        Me.copy_en.Text = "해당 아이템 영문 명칭 복사"
        '
        'copy_kr
        '
        Me.copy_kr.Name = "copy_kr"
        Me.copy_kr.Size = New System.Drawing.Size(254, 22)
        Me.copy_kr.Text = "해당 아이템 한글 명칭 복사"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Enabled = False
        Me.Label1.Font = New System.Drawing.Font("Malgun Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(11, 97)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(385, 13)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "ID, 영문 명칭, 한글 명칭으로 검색하실 수 있습니다. (띄어쓰기에 주의하세요)"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Enabled = False
        Me.Label2.Font = New System.Drawing.Font("Malgun Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(11, 124)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(74, 13)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "▼ 검색 결과 "
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Enabled = False
        Me.Label3.Font = New System.Drawing.Font("Malgun Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(11, 372)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(407, 13)
        Me.Label3.TabIndex = 19
        Me.Label3.Text = "항목에 대고 오른쪽 마우스를 누르면 즐겨찾기 관리 / ID·명칭 복사가 가능합니다."
        '
        'Bookmark_Btn
        '
        Me.Bookmark_Btn.BackColor = System.Drawing.Color.Transparent
        Me.Bookmark_Btn.BackgroundImage = Global.UAItem.My.Resources.Resources.bookmark
        Me.Bookmark_Btn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Bookmark_Btn.FlatAppearance.BorderSize = 0
        Me.Bookmark_Btn.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Silver
        Me.Bookmark_Btn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Bookmark_Btn.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Bookmark_Btn.Location = New System.Drawing.Point(418, 96)
        Me.Bookmark_Btn.Name = "Bookmark_Btn"
        Me.Bookmark_Btn.Size = New System.Drawing.Size(25, 25)
        Me.Bookmark_Btn.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.Bookmark_Btn, "즐겨찾기된 아이템 보기")
        Me.Bookmark_Btn.UseVisualStyleBackColor = False
        '
        'Search_Btn
        '
        Me.Search_Btn.BackColor = System.Drawing.Color.Transparent
        Me.Search_Btn.BackgroundImage = Global.UAItem.My.Resources.Resources.Search
        Me.Search_Btn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Search_Btn.FlatAppearance.BorderSize = 0
        Me.Search_Btn.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Silver
        Me.Search_Btn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Search_Btn.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Search_Btn.Location = New System.Drawing.Point(419, 66)
        Me.Search_Btn.Name = "Search_Btn"
        Me.Search_Btn.Size = New System.Drawing.Size(25, 25)
        Me.Search_Btn.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.Search_Btn, "해당 아이템 검색")
        Me.Search_Btn.UseVisualStyleBackColor = False
        '
        'offbtn
        '
        Me.offbtn.BackColor = System.Drawing.Color.Transparent
        Me.offbtn.BackgroundImage = Global.UAItem.My.Resources.Resources.ic_settings_power_white_48dp1
        Me.offbtn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.offbtn.FlatAppearance.BorderSize = 0
        Me.offbtn.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Silver
        Me.offbtn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.offbtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.offbtn.Location = New System.Drawing.Point(419, 12)
        Me.offbtn.Name = "offbtn"
        Me.offbtn.Size = New System.Drawing.Size(25, 25)
        Me.offbtn.TabIndex = 13
        Me.ToolTip1.SetToolTip(Me.offbtn, "닫기")
        Me.offbtn.UseVisualStyleBackColor = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = Global.UAItem.My.Resources.Resources.banner
        Me.PictureBox1.Enabled = False
        Me.PictureBox1.Location = New System.Drawing.Point(14, 16)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(367, 39)
        Me.PictureBox1.TabIndex = 20
        Me.PictureBox1.TabStop = False
        '
        'ToolTip1
        '
        Me.ToolTip1.AutomaticDelay = 0
        Me.ToolTip1.AutoPopDelay = 5000
        Me.ToolTip1.InitialDelay = 1
        Me.ToolTip1.ReshowDelay = 1
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(461, 396)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Bookmark_Btn)
        Me.Controls.Add(Me.Search_Btn)
        Me.Controls.Add(Me.offbtn)
        Me.Controls.Add(Me.Search_Lst)
        Me.Controls.Add(Me.idfind_txt)
        Me.Controls.Add(Me.PictureBox1)
        Me.Font = New System.Drawing.Font("Malgun Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "Form1"
        Me.Text = "UAItem - 언턴드 아이템 검색기"
        Me.ContextMenuStrip1.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents idfind_txt As TextBox
    Friend WithEvents Search_Lst As ListView
    Friend WithEvents id As ColumnHeader
    Friend WithEvents en_name As ColumnHeader
    Friend WithEvents kr_name As ColumnHeader
    Friend WithEvents offbtn As Button
    Friend WithEvents Search_Btn As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents add_bookmark As ToolStripMenuItem
    Friend WithEvents copy_id As ToolStripMenuItem
    Friend WithEvents copy_en As ToolStripMenuItem
    Friend WithEvents copy_kr As ToolStripMenuItem
    Friend WithEvents Label3 As Label
    Friend WithEvents Bookmark_Btn As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents remove_bookmark As ToolStripMenuItem
End Class
