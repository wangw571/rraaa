Imports System.ComponentModel
Imports Microsoft.VisualBasic

Public Class MainForm
    Dim smStrL As String = "百度贴吧:覅是（原作者xpulai） QQ:54306352" '中文提示在此
    Dim 正在批量修改 As Boolean = False
    Dim 当前存档 As New 存档Class
    Private Sub MainForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.当前模块ID箱.Items.AddRange(模块名称s.ToArray)
        Me.模块选定箱.Items.AddRange(模块名称s.ToArray)
        Me.模块选定箱.Items.Add("未选定") '中文提示在此
        ToolStripComboBox1.SelectedItem = My.MySettings.Default.显示位数.ToString
        OpenFileDialog1.InitialDirectory = My.MySettings.Default.插入存档位置

        载入存档列表(False, 存档搜索箱.Text)
    End Sub
    Private Sub 路径设置ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 路径设置ToolStripMenuItem.Click
        载入存档列表(True, 存档搜索箱.Text)
    End Sub
    Private Sub ToolStripComboBox1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        My.MySettings.Default.显示位数 = CInt(ToolStripComboBox1.SelectedItem)
        My.MySettings.Default.Save()
        输入框检查()
    End Sub
    Friend Sub 载入存档列表(重选C As Boolean, 文件名 As String)
        Dim sitemL As Object = Me.ListBox1.SelectedItem
        Me.ListBox1.Items.Clear()
        If My.MySettings.Default.存档位置 = "" Then
            My.MySettings.Default.存档位置 = IO.Path.GetFullPath(My.Application.Info.DirectoryPath) & "\"
        End If
        If My.MySettings.Default.插入存档位置 = "" Then
            My.MySettings.Default.插入存档位置 = My.MySettings.Default.存档位置
        End If
        My.MySettings.Default.Save()
        FolderBrowserDialog1.SelectedPath = My.MySettings.Default.存档位置
        If 重选C = True AndAlso FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            My.MySettings.Default.存档位置 = FolderBrowserDialog1.SelectedPath & "\"
            My.MySettings.Default.Save()
        End If
        Dim fsL() As String = IO.Directory.GetFiles(My.MySettings.Default.存档位置)
        For i001 As Integer = 0 To fsL.Length - 1
            Dim sss As String = IO.Path.GetExtension(fsL(i001))
            Dim listCount As Int16 = 0
            If IO.Path.GetExtension(fsL(i001)) = ".bsg" AndAlso （IO.Path.GetFileName(fsL(i001)).Contains(文件名) Or 文件名 = ""） Then
                Me.ListBox1.Items.Add(IO.Path.GetFileNameWithoutExtension(fsL(i001))) '列表列出处
                listCount += 1
            End If
            If listCount = 0 Then
                当前存档 = Nothing
                ListBox2.Items.Clear()
                TextBox5.Clear()
            End If
        Next
        If Me.ListBox1.Items.Count > 0 Then
            If sitemL IsNot Nothing AndAlso Me.ListBox1.Items.IndexOf(sitemL) >= 0 Then
                Me.ListBox1.SelectedItem = sitemL
            Else
                Me.ListBox1.SelectedIndex = 0
            End If
        Else
        End If
        按钮检查()
    End Sub
    Friend Function 载入存档(文件名称C As String) As 存档Class
        Dim 存档L As New 存档Class
        Dim fstrLS() As String = My.Computer.FileSystem.ReadAllText(文件名称C).Split(vbNewLine)
        存档L.文件名称 = IO.Path.GetFileNameWithoutExtension(文件名称C)
        存档L.文件路径 = IO.Path.GetDirectoryName(IO.Path.GetFullPath(文件名称C)) & "\"
        存档L.世界坐标 = New 坐标A(fstrLS(13).Trim)
        存档L.世界旋转 = fstrLS(15).Trim

        Dim 编号Ls() As String = fstrLS(1).Split("|")
        Dim 坐标Ls() As String = fstrLS(3).Split("|")
        Dim 四元旋转坐标Ls() As String = fstrLS(5).Split("|")
        Dim 反转Ls() As String = fstrLS(7).Split("|")
        Dim 向量标开始Ls() As String = fstrLS(9).Split("|")
        Dim 向量标结束Ls() As String = fstrLS(11).Split("|")
        Dim 功能键1Ls() As String = fstrLS(17).Split("|")
        Dim 功能键2Ls() As String = fstrLS(19).Split("|")
        Dim 参数值Ls() As String = fstrLS(21).Split("|")
        Dim 开关模式Ls() As String = fstrLS(23).Split("|")
        For i001 As Integer = 0 To 编号Ls.Length - 1
            Dim 模块L As New 模块Class
            模块L.编号 = 编号Ls(i001).Trim
            模块L.坐标 = New 坐标A(坐标Ls(i001).Trim)
            模块L.四元旋转坐标 = New 坐标B(四元旋转坐标Ls(i001).Trim)
            模块L.反转 = 反转Ls(i001).Trim
            模块L.向量标开始 = New 坐标A(向量标开始Ls(i001).Trim)
            模块L.向量标结束 = New 坐标A(向量标结束Ls(i001).Trim)
            模块L.参数值 = 参数值Ls(i001).Trim
            模块L.功能键1 = 功能键1Ls(i001).Trim
            模块L.功能键2 = 功能键2Ls(i001).Trim
            模块L.开关模式 = 开关模式Ls(i001).Trim
            模块L._p模块Coll = 存档L.模块s
            存档L.模块s.Add(模块L)
        Next
        Return 存档L
    End Function
    Dim _编辑状态 As Integer = -1
    Private Sub ListBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        当前存档 = Nothing
        _编辑状态 = 0
        Me.ListBox2.Items.Clear()
        If Me.ListBox1.Items.Count > 0 AndAlso Me.ListBox1.SelectedIndex >= 0 Then
            Dim fPathL As String = My.MySettings.Default.存档位置 & Me.ListBox1.SelectedItem & ".bsg"
            当前存档 = 载入存档(fPathL)
        End If
        If 当前存档 Is Nothing Then
            TextBox1.Text = ""
            世界X坐标box.Text = ""
            世界Y坐标box.Text = ""
            世界Z坐标box.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            Me.Panel5.BackgroundImage = Nothing
            Me.Panel5.Refresh()
        Else
            TextBox1.Text = 当前存档.文件名称
            世界X坐标box.Text = Math.Round(Double.Parse(当前存档.世界坐标.Roll), My.MySettings.Default.显示位数)
            世界Y坐标box.Text = Math.Round(Double.Parse(当前存档.世界坐标.Yaw), My.MySettings.Default.显示位数)
            世界Z坐标box.Text = Math.Round(Double.Parse(当前存档.世界坐标.Pitch), My.MySettings.Default.显示位数)
            If My.MySettings.Default.显示位数 >= 15 Then
                TextBox3.Text = 当前存档.世界旋转
            Else
                TextBox3.Text = Math.Round(当前存档.世界旋转, My.MySettings.Default.显示位数)
            End If
            TextBox4.Text = 当前存档.备注
            Dim fnameL As String = My.MySettings.Default.存档位置 & "Thumbnails\" & 当前存档.文件名称 & ".png"
            If IO.File.Exists(fnameL) = True Then
                Dim bL As Bitmap = Image.FromFile(fnameL)
                bL.MakeTransparent(Color.FromArgb(11, 11, 18, 255))
                Me.Panel5.BackgroundImage = bL
                Label15.Visible = False
            Else
                Me.Panel5.BackgroundImage = Nothing
                Me.Label15.Visible = True
                Me.Label15.Parent = Me.Panel5
                Me.Label15.Location = New Point(Panel5.Location.X + Panel5.Size.Width / 2, Panel5.Location.Y + Panel5.Size.Height / 2)
            End If
            For i003 As Integer = 0 To 当前存档.模块s.Count - 1
                Me.ListBox2.Items.Add(当前存档.模块s(i003))
            Next
            ListBox2.SelectedIndex = 0
            Me.模块选定箱.SelectedIndex = Me.模块选定箱.Items.Count - 1
        End If
        按钮检查()
        TextBox5.Text = 生成存档文本()
        _编辑状态 = -1
    End Sub
    Private Sub ListBox2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ListBox2.SelectedIndexChanged
        输入框检查()
        If Me.ListBox2.SelectedIndices.Count > 1 Then
            模块选定箱.SelectedIndex = 模块选定箱.Items.Count - 1
            正在批量修改 = True
            X旋转输入框.Enabled = False
            Y旋转输入框.Enabled = False
            Z旋转输入框.Enabled = False
            Label21.Enabled = False
            Label6.Text = "批量叠加坐标:"
            'remember to rewrite
            X坐标输入框.Text = 0
            Y坐标输入框.Text = 0
            Z坐标输入框.Text = 0
            X旋转输入框.Text = “不可用”
            Y旋转输入框.Text = “不可用”
            Z旋转输入框.Text = “不可用”
            参数值Box.Text = 1
            Key1Box.Text = "a"
            Key2Box.Text = "b"
            Flip选定箱.SelectedIndex = 0
            Toggle选定箱.SelectedIndex = 0
            TextBox12.Enabled = False
            TextBox13.Enabled = False
            确定批量修改.Visible = True
            确定批量修改.Enabled = True
        Else
            正在批量修改 = False
            Label6.Text = "相对于第一个初始方块的坐标:"
            X旋转输入框.Enabled = True
            Y旋转输入框.Enabled = True
            Z旋转输入框.Enabled = True
            Label21.Enabled = True
            TextBox12.Enabled = True
            TextBox13.Enabled = True
            确定批量修改.Visible = False
            确定批量修改.Enabled = False
        End If
    End Sub
    Private Sub 确定批量修改_Click(sender As Object, e As EventArgs) Handles 确定批量修改.Click
        Dim tempSelectIndex = ListBox2.SelectedIndices
        For Each 选定的index As Int32 In ListBox2.SelectedIndices
            当前存档.模块s(选定的index)._编号 = 当前模块ID箱.SelectedIndex
            当前存档.模块s(选定的index).坐标 = New 坐标A（
                当前存档.模块s(选定的index).坐标.Roll + Double.Parse(X坐标输入框.Text),
                当前存档.模块s(选定的index).坐标.Yaw + Double.Parse(Y坐标输入框.Text),
                当前存档.模块s(选定的index).坐标.Pitch + Double.Parse(Z坐标输入框.Text)
            ）
            当前存档.模块s(选定的index)._功能键1 = Key1Box.Text
            当前存档.模块s(选定的index)._功能键2 = Key2Box.Text
            当前存档.模块s(选定的index).参数值 = 参数值Box.Text
            当前存档.模块s(选定的index)._反转 = Flip选定箱.Text
            当前存档.模块s(选定的index)._开关模式 = Toggle选定箱.Text
        Next
        Dim 被选中编号 As Long = 0
        被选中编号 = Me.模块选定箱.SelectedIndex
        Me.ListBox2.Items.Clear()
        For i003 As Integer = 0 To 当前存档.模块s.Count - 1
            If 被选中编号 < Me.模块选定箱.Items.Count - 1 AndAlso 当前存档.模块s(i003).编号 = 被选中编号 Then
                Me.ListBox2.Items.Add(当前存档.模块s(i003))
            ElseIf 被选中编号 = Me.模块选定箱.Items.Count - 1
                Me.ListBox2.Items.Add(当前存档.模块s(i003))
            End If

        Next
        ListBox2.SelectedIndex = 0
        TextBox5.Text = 生成存档文本()
    End Sub

    Friend Sub 按钮检查()
        If Me.当前存档 Is Nothing Then
            Button1.Enabled = False
            Button2.Enabled = False

            Button4.Enabled = False
            Button5.Enabled = False
            Button6.Enabled = False
        Else
            Button1.Enabled = True
            Button2.Enabled = True

            Button4.Enabled = True
            Button5.Enabled = True
            Button6.Enabled = True
        End If
    End Sub
    Public Function 输入框检查() As Boolean
        _编辑状态 = 0
        Dim 模块L As 模块Class = Me.ListBox2.SelectedItem
        If 模块L Is Nothing Then
            模块L = New 模块Class
            GroupBox4.Enabled = False
        Else
            GroupBox4.Enabled = True
        End If
        Me.当前模块ID箱.SelectedIndex = 模块L.编号
        Me.X坐标输入框.Text = Math.Round(Double.Parse(模块L.坐标.Roll), My.MySettings.Default.显示位数)
        Me.Y坐标输入框.Text = Math.Round(Double.Parse(模块L.坐标.Yaw), My.MySettings.Default.显示位数)
        Me.Z坐标输入框.Text = Math.Round(Double.Parse(模块L.坐标.Pitch), My.MySettings.Default.显示位数)
        Me.X旋转输入框.Text = Math.Round(Double.Parse(模块L.三维旋转坐标.Roll), My.MySettings.Default.显示位数)
        Me.Y旋转输入框.Text = Math.Round(Double.Parse(模块L.三维旋转坐标.Yaw), My.MySettings.Default.显示位数)
        Me.Z旋转输入框.Text = Math.Round(Double.Parse(模块L.三维旋转坐标.Pitch), My.MySettings.Default.显示位数)
        Me.QuaternionBox.Text = 模块L.四元旋转坐标.ToRoundString
        Me.Flip选定箱.Text = 模块L.反转
        Me.Toggle选定箱.Text = 模块L.开关模式
        Me.Key1Box.Text = 模块L.功能键1
        Me.Key2Box.Text = 模块L.功能键2
        If My.MySettings.Default.显示位数 <= 15 Then
            Me.参数值Box.Text = Math.Round(CDbl(模块L.参数值))
        Else
            Me.参数值Box.Text = 模块L.参数值
        End If
        Me.TextBox12.Text = 模块L.向量标开始.ToRoundString
        Me.TextBox13.Text = 模块L.向量标结束.ToRoundString


        _编辑状态 = -1
        Return True
    End Function


    Private Sub TextBox1_Validated(sender As Object, e As System.EventArgs) Handles TextBox1.Validated
        If _编辑状态 = -1 Then
            If 当前存档 IsNot Nothing Then
                当前存档.文件名称 = Me.TextBox1.Text.Trim
                Me.TextBox1.Text = 当前存档.文件名称
                TextBox5.Text = 生成存档文本()
            End If
        End If
    End Sub
    Private Sub 世界X坐标box_Validated(sender As Object, e As System.EventArgs) Handles 世界X坐标box.Validated
        If _编辑状态 = -1 Then
            If 当前存档 IsNot Nothing Then
                Me.世界X坐标box.Text = Me.世界X坐标box.Text.Trim
                If isdou(Me.世界X坐标box.Text, 1) = True Then
                    Dim nzbL As New 坐标A(Double.Parse(Me.世界X坐标box.Text), 当前存档.世界坐标.Yaw, 当前存档.世界坐标.Pitch)
                    If nzbL.ToRoundString <> 当前存档.世界坐标.ToRoundString Then
                        当前存档.世界坐标 = nzbL
                    End If
                End If
                Me.世界X坐标box.Text = Math.Round(Double.Parse(世界X坐标box.Text), My.MySettings.Default.显示位数)
                TextBox5.Text = 生成存档文本()
            End If
        End If
    End Sub
    Private Sub 世界Y坐标box_Validated(sender As Object, e As System.EventArgs) Handles 世界Y坐标box.Validated
        If _编辑状态 = -1 Then
            If 当前存档 IsNot Nothing Then
                Me.世界Y坐标box.Text = Me.世界Y坐标box.Text.Trim
                If isdou(Me.世界Y坐标box.Text, 1) = True Then
                    Dim nzbL As New 坐标A(当前存档.世界坐标.Roll, Double.Parse(Me.世界Y坐标box.Text), 当前存档.世界坐标.Pitch)
                    If nzbL.ToRoundString <> 当前存档.世界坐标.ToRoundString Then
                        当前存档.世界坐标 = nzbL
                    End If
                End If
                Me.世界Y坐标box.Text = Math.Round(Double.Parse(世界Y坐标box.Text), My.MySettings.Default.显示位数)
                TextBox5.Text = 生成存档文本()
            End If
        End If
    End Sub
    Private Sub 世界Z坐标box_Validated(sender As Object, e As System.EventArgs) Handles 世界Z坐标box.Validated
        If _编辑状态 = -1 Then
            If 当前存档 IsNot Nothing Then
                Me.世界Z坐标box.Text = Me.世界Z坐标box.Text.Trim
                If isdou(Me.世界Z坐标box.Text, 1) = True Then
                    Dim nzbL As New 坐标A(当前存档.世界坐标.Roll, 当前存档.世界坐标.Yaw, Double.Parse(Me.世界Z坐标box.Text))
                    If nzbL.ToRoundString <> 当前存档.世界坐标.ToRoundString Then
                        当前存档.世界坐标 = nzbL
                    End If
                End If
                Me.世界Z坐标box.Text = Math.Round(Double.Parse(世界Z坐标box.Text), My.MySettings.Default.显示位数)
                TextBox5.Text = 生成存档文本()
            End If
        End If
    End Sub
    Private Sub TextBox3_Validated(sender As Object, e As System.EventArgs) Handles TextBox3.Validated
        If _编辑状态 = -1 Then
            If 当前存档 IsNot Nothing Then
                Me.TextBox3.Text = Me.TextBox3.Text.Trim
                If isdou(Me.TextBox3.Text, 1) = True Then
                    Dim nzb1L As Double = 当前存档.世界旋转
                    Dim nzb2L As Double = CDbl(Me.TextBox3.Text)
                    If My.MySettings.Default.显示位数 <= 15 Then
                        nzb1L = Math.Round(当前存档.世界旋转, My.MySettings.Default.显示位数)
                        nzb2L = Math.Round(CDbl(Me.TextBox3.Text), My.MySettings.Default.显示位数)
                    End If
                    If nzb1L <> nzb2L Then
                        当前存档.世界旋转 = 当前存档.世界旋转
                    End If
                End If
                If My.MySettings.Default.显示位数 > 15 Then
                    Me.TextBox3.Text = 当前存档.世界旋转
                Else
                    Me.TextBox3.Text = Math.Round(当前存档.世界旋转, My.MySettings.Default.显示位数)
                End If
                TextBox5.Text = 生成存档文本()
            End If
        End If
    End Sub
    Private Sub TextBox4_Validated(sender As Object, e As System.EventArgs) Handles TextBox4.Validated
        If _编辑状态 = -1 Then
            If 当前存档 IsNot Nothing Then
                当前存档.备注 = Me.TextBox4.Text.Trim
                Me.TextBox4.Text = 当前存档.备注
                TextBox5.Text = 生成存档文本()
            End If
        End If
    End Sub

    Private Sub 当前模块ID箱_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles 当前模块ID箱.SelectedIndexChanged
        If _编辑状态 = -1 AndAlso 正在批量修改 <> True Then
            Dim 模块L As 模块Class = Me.ListBox2.SelectedItem
            If 模块L IsNot Nothing Then
                模块L.编号 = Me.当前模块ID箱.SelectedIndex
                Dim inidL As Integer = Me.ListBox2.SelectedIndex
                Me.ListBox2.Items.RemoveAt(inidL)
                Me.ListBox2.Items.Insert(inidL, 模块L)
                Me.ListBox2.SelectedIndex = inidL
                TextBox5.Text = 生成存档文本()
            End If
        End If
    End Sub

    Private Sub 模块选定箱_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles 模块选定箱.SelectedIndexChanged

        Dim 被选中编号 As Long = 0
        被选中编号 = Me.模块选定箱.SelectedIndex
        Me.ListBox2.Items.Clear()
        For i003 As Integer = 0 To 当前存档.模块s.Count - 1
            If 被选中编号 < Me.模块选定箱.Items.Count - 1 AndAlso 当前存档.模块s(i003).编号 = 被选中编号 Then
                Me.ListBox2.Items.Add(当前存档.模块s(i003))
            ElseIf 被选中编号 = Me.模块选定箱.Items.Count - 1
                Me.ListBox2.Items.Add(当前存档.模块s(i003))
            End If

        Next
        TextBox5.Text = 生成存档文本()


    End Sub

    Private Sub X坐标输入框_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles X坐标输入框.KeyDown
        If e.KeyCode = Keys.Enter Then
            X坐标输入框_Validated(X坐标输入框, System.EventArgs.Empty)
        End If
    End Sub
    Private Sub X坐标输入框_Validated(sender As Object, e As System.EventArgs) Handles X坐标输入框.Validated 'X模块位置
        If _编辑状态 = -1 AndAlso 正在批量修改 <> True Then
            Dim 模块L As 模块Class = Me.ListBox2.SelectedItem
            If 模块L IsNot Nothing Then
                Me.X坐标输入框.Text = Me.X坐标输入框.Text.Trim
                If isdou(Me.X坐标输入框.Text, 1) = True Then
                    Dim nzbL As New 坐标A(Double.Parse(Me.X坐标输入框.Text), 模块L.坐标.Yaw, 模块L.坐标.Pitch)
                    If nzbL.ToRoundString <> 模块L.坐标.ToRoundString Then
                        模块L.坐标 = nzbL
                    End If
                End If
                Me.X坐标输入框.Text = Math.Round(Double.Parse(X坐标输入框.Text), My.MySettings.Default.显示位数)
                TextBox5.Text = 生成存档文本()
                '模块L.坐标.
            End If
        End If
    End Sub

    Private Sub Y坐标输入框_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Y坐标输入框.KeyDown
        If e.KeyCode = Keys.Enter Then
            Y坐标输入框_Validated(Y坐标输入框, System.EventArgs.Empty)
        End If
    End Sub
    Private Sub Y坐标输入框_Validated(sender As Object, e As System.EventArgs) Handles Y坐标输入框.Validated 'Y模块位置
        If _编辑状态 = -1 AndAlso 正在批量修改 <> True Then
            Dim 模块L As 模块Class = Me.ListBox2.SelectedItem
            If 模块L IsNot Nothing Then
                Me.Y坐标输入框.Text = Me.Y坐标输入框.Text.Trim
                If isdou(Me.Y坐标输入框.Text, 1) = True Then
                    Dim nzbL As New 坐标A(模块L.坐标.Roll, Double.Parse(Me.Y坐标输入框.Text), 模块L.坐标.Pitch)
                    If nzbL.ToRoundString <> 模块L.坐标.ToRoundString Then
                        模块L.坐标 = nzbL
                    End If
                End If
                Me.Y坐标输入框.Text = Math.Round(Double.Parse(Y坐标输入框.Text), My.MySettings.Default.显示位数)
                TextBox5.Text = 生成存档文本()
            End If
        End If
    End Sub

    Private Sub Z坐标输入框_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Z坐标输入框.KeyDown
        If e.KeyCode = Keys.Enter Then
            Z坐标输入框_Validated(Z坐标输入框, System.EventArgs.Empty)
        End If
    End Sub
    Private Sub Z坐标输入框_Validated(sender As Object, e As System.EventArgs) Handles Z坐标输入框.Validated 'z模块位置
        If _编辑状态 = -1 AndAlso 正在批量修改 <> True Then
            Dim 模块L As 模块Class = Me.ListBox2.SelectedItem
            If 模块L IsNot Nothing Then
                Me.Z坐标输入框.Text = Me.Z坐标输入框.Text.Trim
                If isdou(Me.Z坐标输入框.Text, 1) = True Then
                    Dim nzbL As New 坐标A(模块L.坐标.Roll, 模块L.坐标.Yaw, Double.Parse(Me.Z坐标输入框.Text))
                    If nzbL.ToRoundString <> 模块L.坐标.ToRoundString Then
                        模块L.坐标 = nzbL
                    End If
                End If
                Me.Z坐标输入框.Text = Math.Round(Double.Parse(Z坐标输入框.Text), My.MySettings.Default.显示位数)
                TextBox5.Text = 生成存档文本()
            End If
        End If
    End Sub

    Private Sub X旋转输入框_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles X旋转输入框.KeyDown '模块旋转
        If e.KeyCode = Keys.Enter Then
            X旋转输入框_Validated(X旋转输入框, System.EventArgs.Empty)
        End If
    End Sub
    Private Sub X旋转输入框_Validated(sender As Object, e As System.EventArgs) Handles X旋转输入框.Validated
        If _编辑状态 = -1 AndAlso 正在批量修改 <> True Then
            Dim 模块L As 模块Class = Me.ListBox2.SelectedItem
            If 模块L IsNot Nothing Then
                Me.X旋转输入框.Text = Me.X旋转输入框.Text.Trim
                If isdou(Me.X旋转输入框.Text, 1) = True Then
                    Dim nzbL As New 坐标A(Double.Parse(Me.X旋转输入框.Text), 模块L.三维旋转坐标.Yaw, 模块L.三维旋转坐标.Pitch)
                    If nzbL.ToRoundString <> 模块L.三维旋转坐标.ToRoundString Then
                        模块L.三维旋转坐标 = nzbL
                    End If
                End If
                Me.X旋转输入框.Text = Math.Round(Double.Parse(X旋转输入框.Text), My.MySettings.Default.显示位数)
                QuaternionBox.Text = 模块L.四元旋转坐标.ToRoundString
                TextBox5.Text = 生成存档文本()
            End If
        End If
    End Sub

    Private Sub Y旋转输入框_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Y旋转输入框.KeyDown '模块旋转
        If e.KeyCode = Keys.Enter Then
            Y旋转输入框_Validated(Y旋转输入框, System.EventArgs.Empty)
        End If
    End Sub
    Private Sub Y旋转输入框_Validated(sender As Object, e As System.EventArgs) Handles Y旋转输入框.Validated
        If _编辑状态 = -1 AndAlso 正在批量修改 <> True Then
            Dim 模块L As 模块Class = Me.ListBox2.SelectedItem
            If 模块L IsNot Nothing Then
                Me.Y旋转输入框.Text = Me.Y旋转输入框.Text.Trim
                If isdou(Me.Y旋转输入框.Text, 1) = True Then
                    Dim nzbL As New 坐标A(模块L.三维旋转坐标.Roll, Double.Parse(Me.Y旋转输入框.Text), 模块L.三维旋转坐标.Pitch)
                    If nzbL.ToRoundString <> 模块L.三维旋转坐标.ToRoundString Then
                        模块L.三维旋转坐标 = nzbL
                    End If
                End If
                Me.Y旋转输入框.Text = Math.Round(Double.Parse(Y旋转输入框.Text), My.MySettings.Default.显示位数)
                QuaternionBox.Text = 模块L.四元旋转坐标.ToRoundString
                TextBox5.Text = 生成存档文本()
            End If
        End If
    End Sub

    Private Sub Z旋转输入框_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Z旋转输入框.KeyDown 'z模块旋转
        If e.KeyCode = Keys.Enter Then
            Z旋转输入框_Validated(Z旋转输入框, System.EventArgs.Empty)
        End If
    End Sub
    Private Sub Z旋转输入框_Validated(sender As Object, e As System.EventArgs) Handles Z旋转输入框.Validated
        If _编辑状态 = -1 AndAlso 正在批量修改 <> True Then
            Dim 模块L As 模块Class = Me.ListBox2.SelectedItem
            If 模块L IsNot Nothing Then
                Me.Z旋转输入框.Text = Me.Z旋转输入框.Text.Trim
                If isdou(Me.Z旋转输入框.Text, 1) = True Then
                    Dim nzbL As New 坐标A(模块L.三维旋转坐标.Roll, 模块L.三维旋转坐标.Yaw, Double.Parse(Me.Z旋转输入框.Text))
                    If nzbL.ToRoundString <> 模块L.三维旋转坐标.ToRoundString Then
                        模块L.三维旋转坐标 = nzbL
                    End If
                End If
                Me.Z旋转输入框.Text = Math.Round(Double.Parse(Z旋转输入框.Text), My.MySettings.Default.显示位数)
                QuaternionBox.Text = 模块L.四元旋转坐标.ToRoundString
                TextBox5.Text = 生成存档文本()
            End If
        End If
    End Sub


    Private Sub Flip选定箱_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles Flip选定箱.SelectedIndexChanged
        If _编辑状态 = -1 AndAlso 正在批量修改 <> True Then
            Dim 模块L As 模块Class = Me.ListBox2.SelectedItem
            If 模块L IsNot Nothing Then
                模块L.反转 = Me.Flip选定箱.SelectedItem
                Me.Flip选定箱.SelectedItem = 模块L.反转
                TextBox5.Text = 生成存档文本()
            End If
        End If
    End Sub
    Private Sub Toggle选定箱_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles Toggle选定箱.SelectedIndexChanged
        If _编辑状态 = -1 AndAlso 正在批量修改 <> True Then
            Dim 模块L As 模块Class = Me.ListBox2.SelectedItem
            If 模块L IsNot Nothing Then
                模块L.开关模式 = Me.Toggle选定箱.SelectedItem
                Me.Toggle选定箱.SelectedItem = 模块L.开关模式
                TextBox5.Text = 生成存档文本()
            End If
        End If
    End Sub
    Private Function getKeyStr(keyEC As System.Windows.Forms.KeyEventArgs, kzC As String) As String
        Dim kL As Keys = keyEC.KeyCode
        If kL >= 48 AndAlso kL <= 57 Then
            Return ChrW(keyEC.KeyCode)
        End If
        If kL >= 96 AndAlso kL <= 105 Then
            Return "[" & ChrW(keyEC.KeyCode - 48) & "]"
        End If
        If kL = 13 Then
            Return "Enter"
        End If
        If kL = 106 Then
            Return "[*]"
        End If
        If kL = 107 Then
            Return "[+]"
        End If
        If kL = 109 Then
            Return "[-]"
        End If
        If kL = 110 Then
            Return "[.]"
        End If
        If kL = 111 Then
            Return "[/]"
        End If
        If kL = 38 Then
            Return "up"
        End If
        If kL = 40 Then
            Return "down"
        End If
        If kL = 37 Then
            Return "left"
        End If
        If kL = 39 Then
            Return "right"
        End If
        If kL = 192 Then
            Return "`"
        End If
        If kL >= 65 AndAlso kL <= 90 Then
            Return ChrW(keyEC.KeyCode).ToString.ToLower
        End If
        Return kzC
    End Function
    Private Sub Key1Box_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Key1Box.KeyDown
        Key1Box.Text = getKeyStr(e, "a")
        If _编辑状态 = -1 AndAlso 正在批量修改 <> True Then
            Dim 模块L As 模块Class = Me.ListBox2.SelectedItem
            If 模块L IsNot Nothing Then
                模块L.功能键1 = getKeyStr(e, 模块L.功能键1)
                Key1Box.Text = 模块L.功能键1
                TextBox5.Text = 生成存档文本()
            End If
        End If
    End Sub
    Private Sub Key2Box_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Key2Box.KeyDown
        Key2Box.Text = getKeyStr(e, "b")
        If _编辑状态 = -1 AndAlso 正在批量修改 <> True Then
            Dim 模块L As 模块Class = Me.ListBox2.SelectedItem
            If 模块L IsNot Nothing Then
                模块L.功能键2 = getKeyStr(e, 模块L.功能键2)
                Key2Box.Text = 模块L.功能键2
                TextBox5.Text = 生成存档文本()
            End If
        End If
    End Sub
    Private Sub 参数值Box_改动(sender As Object, e As System.EventArgs) Handles 参数值Box.Validated
        If _编辑状态 = -1 AndAlso 正在批量修改 <> True Then
            Dim 模块L As 模块Class = Me.ListBox2.SelectedItem
            If 模块L IsNot Nothing Then
                Me.参数值Box.Text = Me.参数值Box.Text.Trim
                If isdou(Me.参数值Box.Text, 1) = True Then
                    Dim nzb1L As Double = 模块L.参数值
                    Dim nzb2L As Double = CDbl(Me.参数值Box.Text)
                    If My.MySettings.Default.显示位数 <= 15 Then
                        nzb1L = Math.Round(CDbl(模块L.参数值), My.MySettings.Default.显示位数)
                        nzb2L = Math.Round(CDbl(Me.参数值Box.Text), My.MySettings.Default.显示位数)
                    End If
                    If nzb1L <> nzb2L Then
                        模块L.参数值 = CDbl(Me.参数值Box.Text)
                    End If
                End If
                If My.MySettings.Default.显示位数 > 15 Then
                    Me.参数值Box.Text = 模块L.参数值
                Else
                    Me.参数值Box.Text = Math.Round(CDbl(模块L.参数值), My.MySettings.Default.显示位数)
                End If
                TextBox5.Text = 生成存档文本()
            End If
        End If
    End Sub

    Private Sub TextBox12_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox12.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox12_Validated(TextBox12, System.EventArgs.Empty)
        End If
    End Sub
    Private Sub TextBox13_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox13.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox13_Validated(TextBox13, System.EventArgs.Empty)
        End If
    End Sub


    Private Sub TextBox12_Validated(sender As Object, e As System.EventArgs) Handles TextBox12.Validated
        If _编辑状态 = -1 AndAlso 正在批量修改 <> True Then
            Dim 模块L As 模块Class = Me.ListBox2.SelectedItem
            If 模块L IsNot Nothing Then
                Me.TextBox12.Text = Me.TextBox12.Text.Trim
                If isdou(Me.TextBox12.Text, 3) = True Then
                    Dim nzbL As New 坐标A(Me.TextBox12.Text)
                    If nzbL.ToRoundString <> 模块L.向量标开始.ToRoundString Then
                        模块L.向量标开始 = nzbL
                    End If
                End If
                Me.TextBox12.Text = 模块L.向量标开始.ToRoundString
                TextBox5.Text = 生成存档文本()
            End If
        End If
    End Sub
    Private Sub TextBox13_Validated(sender As Object, e As System.EventArgs) Handles TextBox13.Validated
        If _编辑状态 = -1 AndAlso 正在批量修改 <> True Then
            Dim 模块L As 模块Class = Me.ListBox2.SelectedItem
            If 模块L IsNot Nothing Then
                Me.TextBox13.Text = Me.TextBox13.Text.Trim
                If isdou(Me.TextBox13.Text, 3) = True Then
                    Dim nzbL As New 坐标A(Me.TextBox13.Text)
                    If nzbL.ToRoundString <> 模块L.向量标结束.ToRoundString Then
                        模块L.向量标结束 = nzbL
                    End If
                End If
                Me.TextBox13.Text = 模块L.向量标结束.ToRoundString
                TextBox5.Text = 生成存档文本()
            End If
        End If
    End Sub

    Friend Function 生成存档文本() As String
        Dim 存档文本L As String = ""
        Try
            If 当前存档 IsNot Nothing Then
                Dim 编号Ls As String = "编号" & vbNewLine
                Dim 坐标Ls As String = "坐标" & vbNewLine
                Dim 四元旋转坐标Ls As String = "四元旋转坐标" & vbNewLine
                Dim is反转Ls As String = "反转" & vbNewLine
                Dim 向量标开始Ls As String = "向量标开始" & vbNewLine
                Dim 向量标结束Ls As String = "向量标结束" & vbNewLine
                Dim 世界坐标Ls As String = "世界坐标" & vbNewLine & 当前存档.世界坐标.ToString
                Dim 世界旋转Ls As String = "世界旋转" & vbNewLine & 当前存档.世界旋转.ToString
                Dim 功能键1Ls As String = "功能键1" & vbNewLine
                Dim 功能键2Ls As String = "功能键2" & vbNewLine
                Dim 参数值Ls As String = "参数值" & vbNewLine
                Dim 开关模式Ls As String = "开关模式" & vbNewLine
                Dim 备注Ls As String = "备注" & vbNewLine & 当前存档.备注

                For i001 As Integer = 0 To 当前存档.模块s.Count - 1
                    If i001 > 0 Then
                        编号Ls &= "|"
                        坐标Ls &= "|"
                        四元旋转坐标Ls &= "|"
                        is反转Ls &= "|"
                        向量标开始Ls &= "|"
                        向量标结束Ls &= "|"
                        功能键1Ls &= "|"
                        功能键2Ls &= "|"
                        参数值Ls &= "|"
                        开关模式Ls &= "|"
                    End If
                    编号Ls &= 当前存档.模块s(i001).编号
                    坐标Ls &= 当前存档.模块s(i001).坐标.ToString
                    四元旋转坐标Ls &= 当前存档.模块s(i001).四元旋转坐标.ToString
                    is反转Ls &= 当前存档.模块s(i001).反转.Trim
                    向量标开始Ls &= 当前存档.模块s(i001).向量标开始.ToString
                    向量标结束Ls &= 当前存档.模块s(i001).向量标结束.ToString
                    功能键1Ls &= 当前存档.模块s(i001).功能键1
                    功能键2Ls &= 当前存档.模块s(i001).功能键2
                    参数值Ls &= 当前存档.模块s(i001).参数值
                    开关模式Ls &= 当前存档.模块s(i001).开关模式
                Next
                存档文本L &= 编号Ls & vbNewLine
                存档文本L &= 坐标Ls & vbNewLine
                存档文本L &= 四元旋转坐标Ls & vbNewLine
                存档文本L &= is反转Ls & vbNewLine
                存档文本L &= 向量标开始Ls & vbNewLine
                存档文本L &= 向量标结束Ls & vbNewLine
                存档文本L &= 世界坐标Ls & vbNewLine
                存档文本L &= 世界旋转Ls & vbNewLine
                存档文本L &= 功能键1Ls & vbNewLine
                存档文本L &= 功能键2Ls & vbNewLine
                存档文本L &= 参数值Ls & vbNewLine
                存档文本L &= 开关模式Ls & vbNewLine
                存档文本L &= 备注Ls & vbNewLine
                存档文本L &= vbNewLine
                存档文本L &= vbNewLine
                存档文本L &= smStrL & vbNewLine
            End If
        Catch ex As Exception
        End Try
        Return 存档文本L
    End Function

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        If SaveFileDialog1.FileName.Trim = "" Then
            SaveFileDialog1.FileName = 当前存档.文件名称 & ".bsg"
            SaveFileDialog1.InitialDirectory = My.MySettings.Default.存档位置
        End If
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, 生成存档文本, False, System.Text.Encoding.Default)
            当前存档.文件路径 = IO.Path.GetDirectoryName(IO.Path.GetFullPath(SaveFileDialog1.FileName)) & "\"
            当前存档.文件名称 = IO.Path.GetFileNameWithoutExtension(SaveFileDialog1.FileName)
            SaveFileDialog1.InitialDirectory = 当前存档.文件路径
            载入存档列表(False, 存档搜索箱.Text)
            Me.模块选定箱.SelectedIndex = Me.模块选定箱.Items.Count - 1
        End If
    End Sub
    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        If 当前存档 IsNot Nothing Then
            Dim fns As String = 当前存档.文件路径 & 当前存档.文件名称 & ".bsg"
            If IO.File.Exists(fns) = True Then
                If MessageBox.Show("你是否要从磁盘中删除此存档？此操作不可恢复！", "提示！", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then '中文提示在此
                    IO.File.Delete(fns)
                    载入存档列表(False, 存档搜索箱.Text)
                    Me.模块选定箱.SelectedIndex = Me.模块选定箱.Items.Count - 1
                End If
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If 当前存档 IsNot Nothing AndAlso ListBox2.Items.Count > 0 Then
            _编辑状态 = 0
            Dim nmkL As 模块Class = Nothing
            If ListBox2.SelectedItems.Count > 0 Then
                For i001 As Integer = 0 To ListBox2.SelectedItems.Count - 1
                    nmkL = ListBox2.SelectedItems(i001)
                    nmkL = nmkL.Clone
                    当前存档.模块s.Add(nmkL)
                    Me.ListBox2.Items.Add(nmkL)
                Next
            Else
                nmkL = New 模块Class
                nmkL._p模块Coll = 当前存档.模块s
                当前存档.模块s.Add(nmkL)
                Me.ListBox2.Items.Add(nmkL)
            End If
            _编辑状态 = -1
            '/*/*/*
            Dim 被选中编号 As Long = 0
            被选中编号 = Me.模块选定箱.SelectedIndex
            Me.ListBox2.Items.Clear()
            For i003 As Integer = 0 To 当前存档.模块s.Count - 1
                If 被选中编号 < Me.模块选定箱.Items.Count - 1 AndAlso 当前存档.模块s(i003).编号 = 被选中编号 Then
                    Me.ListBox2.Items.Add(当前存档.模块s(i003))
                ElseIf 被选中编号 = Me.模块选定箱.Items.Count - 1
                    Me.ListBox2.Items.Add(当前存档.模块s(i003))
                End If

            Next
            '/*/*/*
            ListBox2.SelectedIndices.Clear()
            ListBox2.SelectedItem = nmkL
            TextBox5.Text = 生成存档文本()
        Else
            MsgBox("您并未选中任何模块！", MsgBoxStyle.Critical, "警告！") '中文提示在此
        End If
    End Sub
    Private Sub 存档搜索箱被改动(sender As System.Object, e As System.EventArgs) Handles 存档搜索箱.TextChanged
        载入存档列表(False, 存档搜索箱.Text)
        If ListBox1.Items.Count = 0 Then
            当前存档 = Nothing
            ListBox2.Items.Clear()
            TextBox5.Clear()
        End If
    End Sub
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click ''Sigh
        If 当前存档 IsNot Nothing AndAlso ListBox2.Items.Count > 0 Then
            If ListBox2.SelectedItems.Count > 0 Then
                _编辑状态 = 0
                For i001 As Integer = 0 To ListBox2.SelectedItems.Count - 1
                    Dim mkL As 模块Class = ListBox2.SelectedItems(i001)
                    当前存档.模块s.Remove(mkL)
                Next
                Dim szidL As Integer = ListBox2.SelectedIndex - 1
                Me.ListBox2.Items.Clear()
                For i002 As Integer = 0 To 当前存档.模块s.Count - 1
                    Me.ListBox2.Items.Add(当前存档.模块s(i002))
                Next
                If szidL >= Me.ListBox2.Items.Count Then
                    szidL = Me.ListBox2.Items.Count - 1
                End If
                If szidL >= 0 Then
                    ListBox2.SelectedIndex = szidL
                End If
                _编辑状态 = -1
                '/*/*/*
                Dim 被选中编号 As Long = 0
                被选中编号 = Me.模块选定箱.SelectedIndex
                Me.ListBox2.Items.Clear()
                For i003 As Integer = 0 To 当前存档.模块s.Count - 1
                    If 被选中编号 < Me.模块选定箱.Items.Count - 1 AndAlso 当前存档.模块s(i003).编号 = 被选中编号 Then
                        Me.ListBox2.Items.Add(当前存档.模块s(i003))
                    ElseIf 被选中编号 = Me.模块选定箱.Items.Count - 1
                        Me.ListBox2.Items.Add(当前存档.模块s(i003))
                    End If

                Next
                '/*/*/*
                TextBox5.Text = 生成存档文本()
            End If
        Else
            MsgBox("您并未选中任何模块！", MsgBoxStyle.Critical, "警告！") '中文提示在此

        End If
    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        If 当前存档 IsNot Nothing Then
            If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim 存档L As 存档Class = 载入存档(OpenFileDialog1.FileName)
                Dim fstrLS() As String = My.Computer.FileSystem.ReadAllText(当前存档.文件路径 + 当前存档.文件名称 + ".bsg").Split(vbNewLine)
                Dim 原存档坐标 As 坐标A = New 坐标A(fstrLS(13).Trim)
                My.MySettings.Default.插入存档位置 = IO.Path.GetDirectoryName(OpenFileDialog1.FileName)
                My.MySettings.Default.Save()
                OpenFileDialog1.InitialDirectory = My.MySettings.Default.插入存档位置
                Dim incfL As New InCForm
                incfL.初始化(存档L, 原存档坐标)
                If incfL.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
                    Me.模块选定箱.SelectedIndex = Me.模块选定箱.Items.Count - 1
                    当前存档.模块s.AddRange(incfL._存档L.模块s.ToArray)
                    Me.ListBox2.Items.AddRange(incfL._存档L.模块s.ToArray)
                    TextBox5.Text = 生成存档文本()
                End If
            End If
        End If

    End Sub



    Private Sub Button8_Click(sender As System.Object, e As System.EventArgs) Handles Button8.Click
        Dim PEFormL As New PEForm
        PEFormL.ShowDialog()
    End Sub

    Private Sub 存档修复ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 存档修复ToolStripMenuItem.Click
        Dim 核心存在数 = 0
        'Dim 核心的索引 As ArrayList
        For Each 被检测模块 As 模块Class In 当前存档.模块s
            If 被检测模块.编号 = 0 Then
                核心存在数 += 1
            End If
        Next
        If 核心存在数 Mod 2 = 0 AndAlso 核心存在数 <> 0 Then
            Dim 这是第几个核心 = 0
            For index = 0 To 当前存档.模块s.Count - 1
                If 这是第几个核心 < （核心存在数 / 2） + 1 Then
                    If 当前存档.模块s(index).编号 = 0 Then
                        这是第几个核心 += 1
                    End If
                Else
                    'MsgBox(index)
                    当前存档.模块s.RemoveRange(index, 当前存档.模块s.Count - index)
                    当前存档.模块s.RemoveAt(当前存档.模块s.Count - 1)
                    TextBox5.Text = 生成存档文本()
                    Me.ListBox2.Items.Clear()
                    For i003 As Integer = 0 To 当前存档.模块s.Count - 1
                        Me.ListBox2.Items.Add(当前存档.模块s(i003))
                    Next
                    ListBox2.SelectedIndex = 0
                    Me.模块选定箱.SelectedIndex = Me.模块选定箱.Items.Count - 1

                    Exit For
                End If
            Next
        Else
            MsgBox("无法寻找到重叠部分！请确认有偶数个起始模块存在！", MsgBoxStyle.Exclamation, "存档修复警告") '中文提示在此
        End If
    End Sub


End Class


Public Class 存档Class
    Public Sub New()

    End Sub

    Dim _文件路径 As String
    Public Property 文件路径 As String
        Get
            Return _文件路径
        End Get
        Set(value As String)
            If Object.Equals(MyClass._文件路径, value) = False Then
                MyClass._文件路径 = value
            End If
        End Set
    End Property

    Dim _文件名称 As String
    Public Property 文件名称 As String
        Get
            Return _文件名称
        End Get
        Set(value As String)
            If Object.Equals(MyClass._文件名称, value) = False Then
                MyClass._文件名称 = value
            End If
        End Set
    End Property
    Dim _世界坐标 As 坐标A
    Public Property 世界坐标 As 坐标A
        Get
            Return _世界坐标
        End Get
        Set(value As 坐标A)
            If Object.Equals(MyClass._世界坐标, value) = False Then
                MyClass._世界坐标 = value
            End If
        End Set
    End Property
    Dim _世界旋转 As Double
    Public Property 世界旋转 As Double
        Get
            Return _世界旋转
        End Get
        Set(value As Double)
            If Object.Equals(MyClass._世界旋转, value) = False Then
                MyClass._世界旋转 = value
            End If
        End Set
    End Property
    Dim _备注 As String
    Public Property 备注 As String
        Get
            Return _备注
        End Get
        Set(value As String)
            If Object.Equals(MyClass._备注, value) = False Then
                MyClass._备注 = value
            End If
        End Set
    End Property

    Dim _模块s As New 模块Coll
    Public Property 模块s As 模块Coll
        Get
            Return _模块s
        End Get
        Set(value As 模块Coll)
            If Object.Equals(MyClass._模块s, value) = False Then
                MyClass._模块s = value
            End If
        End Set
    End Property
End Class

Public Class 模块Class
    Inherits Object
    Protected Friend _p模块Coll As 模块Coll
    Public Sub New()
        MyBase.New()
    End Sub

    Protected Friend _编号 As String = "0"
    Public Property 编号 As String
        Get
            Return _编号
        End Get
        Set(value As String)
            If CInt(value) >= 模块名称s.Count Then
                value = "0"
            End If
            If CInt(value) < 0 Then
                value = "0"
            End If
            If Object.Equals(MyClass._编号, value) = False Then
                MyClass._编号 = value
            End If
        End Set
    End Property
    Public ReadOnly Property 名称 As String
        Get
            If CInt(编号) >= 模块名称s.Count OrElse CInt(编号) < 0 Then
                Return ""
            End If
            Return 模块名称s(CInt(编号))
        End Get
    End Property
    Protected Friend _坐标 As 坐标A
    Public Property 坐标 As 坐标A
        Get
            Return _坐标
        End Get
        Set(value As 坐标A)
            If Object.Equals(MyClass._坐标, value) = False Then
                MyClass._坐标 = value
            End If
        End Set
    End Property
    Protected Friend _四元旋转坐标 As 坐标B
    Public Property 四元旋转坐标 As 坐标B
        Get
            Return _四元旋转坐标
        End Get
        Set(value As 坐标B)
            If Object.Equals(MyClass._四元旋转坐标, value) = False Then
                MyClass._四元旋转坐标 = value
            End If
        End Set
    End Property
    Public Property 三维旋转坐标 As 坐标A
        Get
            Return New 坐标A(_四元旋转坐标)
        End Get
        Set(value As 坐标A)
            MyClass._四元旋转坐标 = New 坐标B(value)
        End Set
    End Property
    Protected Friend _反转 As String = "0"
    Public Property 反转 As String
        Get
            Return _反转
        End Get
        Set(value As String)
            If Object.Equals(MyClass._反转, value) = False Then
                MyClass._反转 = value
            End If
        End Set
    End Property
    Protected Friend _向量标开始 As New 坐标A("90000", "90000", "90000")
    Public Property 向量标开始 As 坐标A
        Get
            Return _向量标开始
        End Get
        Set(value As 坐标A)
            If Object.Equals(MyClass._向量标开始, value) = False Then
                MyClass._向量标开始 = value
            End If
        End Set
    End Property
    Protected Friend _向量标结束 As New 坐标A("90000", "90000", "90000")
    Public Property 向量标结束 As 坐标A
        Get
            Return _向量标结束
        End Get
        Set(value As 坐标A)
            If Object.Equals(MyClass._向量标结束, value) = False Then
                MyClass._向量标结束 = value
            End If
        End Set
    End Property
    Protected Friend _功能键1 As String = "a"
    Public Property 功能键1 As String
        Get
            Return _功能键1
        End Get
        Set(value As String)
            If Object.Equals(MyClass._功能键1, value) = False Then
                MyClass._功能键1 = value
            End If
        End Set
    End Property
    Protected Friend _功能键2 As String = "b"
    Public Property 功能键2 As String
        Get
            Return _功能键2
        End Get
        Set(value As String)
            If Object.Equals(MyClass._功能键2, value) = False Then
                MyClass._功能键2 = value
            End If
        End Set
    End Property
    Protected Friend _参数值 As String = "1"
    Public Property 参数值 As String
        Get
            Return _参数值
        End Get
        Set(value As String)
            If Object.Equals(MyClass._参数值, value) = False Then
                MyClass._参数值 = value
            End If
        End Set
    End Property
    Protected Friend _开关模式 As String = "False"
    Public Property 开关模式 As String
        Get
            Return _开关模式
        End Get
        Set(value As String)
            If Object.Equals(MyClass._开关模式, value) = False Then
                MyClass._开关模式 = value
            End If
        End Set
    End Property
    Public Overrides Function ToString() As String
        Dim msL As String = Me._p模块Coll.IndexOf(Me).ToString.PadRight(6) & Me.名称.PadRight(15)
        Return (msL)
    End Function

    Protected Friend _Image As Image = Nothing
    Public Property Image As Image
        Get
            Return _Image
        End Get
        Set(value As Image)
            If Object.Equals(MyClass._Image, value) = False Then
                MyClass._Image = value
            End If
        End Set
    End Property
    Public ReadOnly Property 体积 As 坐标A
        Get
            If CInt(编号) >= 模块体积s.Count OrElse CInt(编号) < 0 Then
                Return New 坐标A
            End If
            Return 模块体积s(CInt(编号))
        End Get
    End Property
    Public ReadOnly Property 原点 As 坐标A
        Get
            If CInt(编号) >= 模块原点s.Count OrElse CInt(编号) < 0 Then
                Return New 坐标A
            End If
            Return 模块原点s(CInt(编号))
        End Get
    End Property

 
    Public Function Clone() As 模块Class
        Dim cpL As 模块Class = MyBase.MemberwiseClone()
        cpL._p模块Coll = Me._p模块Coll
        cpL._编号 = Me._编号
        cpL._坐标 = Me._坐标
        cpL._四元旋转坐标 = Me._四元旋转坐标
        cpL._反转 = Me._反转
        cpL._向量标开始 = Me._向量标开始
        cpL._向量标结束 = Me._向量标结束
        cpL._功能键1 = Me._功能键1
        cpL._功能键2 = Me._功能键2
        cpL._参数值 = Me._参数值
        cpL._开关模式 = Me._开关模式
        cpL._Image = Me._Image
        Return cpL
    End Function



End Class

Public Class 模块Coll
    Inherits Collections.Generic.List(Of 模块Class)

End Class

Public Structure 坐标A
    Shared Sub New()
    End Sub
    Public Sub New(坐标BC As 坐标B)
        Dim gdjgYL As Double = 2 * (坐标BC.W * 坐标BC.X - 坐标BC.Y * 坐标BC.Z)
        If gdjgYL > 1 Then
            gdjgYL = 1
        End If
        If gdjgYL < -1 Then
            gdjgYL = -1
        End If
        Roll = Math.Asin(gdjgYL) * 180 / Math.PI
        Yaw = Math.Atan2(2 * (坐标BC.W * 坐标BC.Y + 坐标BC.Z * 坐标BC.X), 1 - 2 * (坐标BC.X * 坐标BC.X + 坐标BC.Y * 坐标BC.Y)) * 180 / Math.PI
        Pitch = Math.Atan2(2 * (坐标BC.W * 坐标BC.Z + 坐标BC.X * 坐标BC.Y), 1 - 2 * (坐标BC.Z * 坐标BC.Z + 坐标BC.X * 坐标BC.X)) * 180 / Math.PI
    End Sub
    Public Sub New(xC As Double, yC As Double, zC As Double)
        Me.Roll = xC
        Me.Yaw = yC
        Me.Pitch = zC
    End Sub
    Public Sub New(xyzC As String)
        Try
            Dim xyzSL() As String = xyzC.Trim.Split(",")
            If xyzSL.Count = 3 Then
                Me.Roll = xyzSL(0)
                Me.Yaw = xyzSL(1)
                Me.Pitch = xyzSL(2)
            End If
        Catch ex As Exception
            Me.Roll = 0
            Me.Yaw = 0
            Me.Pitch = 0
        End Try
    End Sub

    Dim _Pitch As Double
    Public Property Pitch As Double
        Get
            Return _Pitch
        End Get
        Set(value As Double)
            If Object.Equals(MyClass._Pitch, value) = False Then
                MyClass._Pitch = value
            End If
        End Set
    End Property
    Dim _Yaw As Double
    Public Property Yaw As Double
        Get
            Return _Yaw
        End Get
        Set(value As Double)
            If Object.Equals(MyClass._Yaw, value) = False Then
                MyClass._Yaw = value
            End If
        End Set
    End Property
    Dim _Roll As Double
    Public Property Roll As Double
        Get
            Return _Roll
        End Get
        Set(value As Double)
            If Object.Equals(MyClass._Roll, value) = False Then
                MyClass._Roll = value
            End If
        End Set
    End Property

    Public Overrides Function ToString() As String
        Dim msL As String = Me.Roll & "," & Me.Yaw & "," & Me.Pitch
        Return msL
    End Function
    Public Function ToRoundString() As String
        Dim msL As String = Me.Pitch & "," & Me.Yaw & "," & Me.Roll
        If My.MySettings.Default.显示位数 >= 15 Then
            msL = Me.Roll & "," & Me.Yaw & "," & Me.Pitch
        Else
            msL = Math.Round(Me.Roll, My.MySettings.Default.显示位数) & "," & Math.Round(Me.Yaw, My.MySettings.Default.显示位数) & "," & Math.Round(Me.Pitch, My.MySettings.Default.显示位数)
        End If
        Return msL
    End Function
End Structure

Public Structure 坐标B
    Shared Sub New()
    End Sub
    Public Sub New(坐标AC As 坐标A)
        坐标AC.Pitch = 坐标AC.Pitch * Math.PI / 180 / 2
        坐标AC.Yaw = 坐标AC.Yaw * Math.PI / 180 / 2
        坐标AC.Roll = 坐标AC.Roll * Math.PI / 180 / 2
        W = Math.Cos(坐标AC.Pitch) * Math.Cos(坐标AC.Yaw) * Math.Cos(坐标AC.Roll) + Math.Sin(坐标AC.Pitch) * Math.Sin(坐标AC.Yaw) * Math.Sin(坐标AC.Roll)
        X = Math.Sin(坐标AC.Pitch) * Math.Cos(坐标AC.Yaw) * Math.Cos(坐标AC.Roll) + Math.Cos(坐标AC.Pitch) * Math.Sin(坐标AC.Yaw) * Math.Sin(坐标AC.Roll)
        Y = Math.Cos(坐标AC.Pitch) * Math.Sin(坐标AC.Yaw) * Math.Cos(坐标AC.Roll) - Math.Sin(坐标AC.Pitch) * Math.Cos(坐标AC.Yaw) * Math.Sin(坐标AC.Roll)
        Z = Math.Cos(坐标AC.Pitch) * Math.Cos(坐标AC.Yaw) * Math.Sin(坐标AC.Roll) - Math.Sin(坐标AC.Pitch) * Math.Sin(坐标AC.Yaw) * Math.Cos(坐标AC.Roll)
    End Sub
    Public Sub New(WC As Double, XC As Double, YC As Double, ZC As Double)
        Me.W = WC
        Me.X = XC
        Me.Y = YC
        Me.Z = ZC
    End Sub
    Public Sub New(q4C As String)
        Dim q4SL() As String = q4C.Trim.Split(",")
        If q4SL.Count = 4 Then
            Me.X = q4SL(0)
            Me.Y = q4SL(1)
            Me.Z = q4SL(2)
            Me.W = q4SL(3)
        ElseIf q4SL.Count = 1 Then
            Me.X = q4SL(3)
            Me.Y = q4SL(3)
            Me.Z = q4SL(3)
            Me.W = q4SL(3)
        End If
    End Sub

    Dim _W As Double
    Public Property W As Double
        Get
            Return _W
        End Get
        Set(value As Double)
            If Object.Equals(MyClass._W, value) = False Then
                MyClass._W = value
            End If
        End Set
    End Property
    Dim _X As Double
    Public Property X As Double
        Get
            Return _X
        End Get
        Set(value As Double)
            If Object.Equals(MyClass._X, value) = False Then
                MyClass._X = value
            End If
        End Set
    End Property
    Dim _Y As Double
    Public Property Y As Double
        Get
            Return _Y
        End Get
        Set(value As Double)
            If Object.Equals(MyClass._Y, value) = False Then
                MyClass._Y = value
            End If
        End Set
    End Property
    Dim _Z As Double
    Public Property Z As Double
        Get
            Return _Z
        End Get
        Set(value As Double)
            If Object.Equals(MyClass._Z, value) = False Then
                MyClass._Z = value
            End If
        End Set
    End Property

    Public Overrides Function ToString() As String
        Dim msL As String = Me.X & "," & Me.Y & "," & Me.Z & "," & Me.W
        Return msL
    End Function
    Public Function ToRoundString() As String
        Dim msL As String = Me.X & "," & Me.Y & "," & Me.Z & "," & Me.W
        If My.MySettings.Default.显示位数 >= 15 Then
            msL = Me.X & "," & Me.Y & "," & Me.Z & "," & Me.W
        Else
            msL = Math.Round(Me.X, My.MySettings.Default.显示位数) & "," & Math.Round(Me.Y, My.MySettings.Default.显示位数) & "," & Math.Round(Me.Z, My.MySettings.Default.显示位数) & "," & Math.Round(Me.W, My.MySettings.Default.显示位数)
        End If
        Return msL
    End Function
End Structure

