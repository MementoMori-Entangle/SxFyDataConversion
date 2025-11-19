Imports Microsoft.Win32
Imports System.IO
Imports System.Text
Imports System.Windows.Forms

Namespace SxFyDataConversion
    Partial Class MainWindow
        Private Shared ReadOnly SupportedTypes As String() = {"U1", "U2", "U4", "U8", "I1", "I2", "I4", "I8", "F4", "F8", "A", "B"}
        Private lastFolderPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

        ' 出力エンコーディング取得
        Private Function GetSelectedEncoding() As Encoding
            If CBEncoding Is Nothing OrElse CBEncoding.SelectedIndex = -1 Then
                Return Encoding.GetEncoding("Shift_JIS")
            End If
            Select Case CBEncoding.SelectedIndex
                Case 0
                    Return Encoding.GetEncoding("Shift_JIS")
                Case 1
                    Return New UTF8Encoding(True) ' BOMあり
                Case 2
                    Return New UTF8Encoding(False) ' BOMなし
                Case Else
                    Return Encoding.GetEncoding("Shift_JIS")
            End Select
        End Function

        Private Sub BtnConvert_Click(sender As Object, e As RoutedEventArgs) Handles BtnConvert.Click
            Dim csvFiles As New List(Of String)()
            If RBFile.IsChecked Then
                ' 複数ファイル選択
                Dim openFileDialog As New Microsoft.Win32.OpenFileDialog With {
                    .Filter = "CSVファイル (*.csv)|*.csv",
                    .Multiselect = True
                }
                If openFileDialog.ShowDialog() = True Then
                    csvFiles.AddRange(openFileDialog.FileNames)
                Else
                    Return
                End If
            ElseIf RBFolder.IsChecked Then
                ' フォルダ選択
                Dim folderDialog As New System.Windows.Forms.FolderBrowserDialog With {
                    .SelectedPath = lastFolderPath
                }
                Dim result = folderDialog.ShowDialog()
                If result = System.Windows.Forms.DialogResult.OK Then
                    Dim folder = folderDialog.SelectedPath
                    lastFolderPath = folder ' パスを記憶
                    csvFiles.AddRange(Directory.GetFiles(folder, "*.csv"))
                    If csvFiles.Count = 0 Then
                        System.Windows.MessageBox.Show("CSVファイルが見つかりません。", "情報", MessageBoxButton.OK, MessageBoxImage.Information)
                        Return
                    End If
                Else
                    Return
                End If
            Else
                System.Windows.MessageBox.Show("ファイル/フォルダ選択方式を選んでください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error)
                Return
            End If

            Dim encoding = GetSelectedEncoding()
            Dim successCount As Integer = 0
            Dim failCount As Integer = 0
            Dim failFiles As New List(Of String)()
            For Each csvPath In csvFiles
                Dim smlPath = Path.ChangeExtension(csvPath, ".sml")
                Try
                    Dim lines = File.ReadAllLines(csvPath, Encoding.UTF8)
                    Dim headerLine = lines(0).Trim()
                    Dim smlHeader = ConvertHeader(headerLine)
                    Dim sml As String
                    If lines.Length = 1 Then
                        ' データ行がない場合はヘッダーのみ
                        sml = smlHeader & vbCrLf & "."
                    Else
                        Dim index As Integer = 1 ' 1行目(ヘッダー)をスキップ
                        Dim smlBody = ParseAny(lines, index, 0, True)
                        sml = smlHeader & vbCrLf & smlBody & "."
                    End If
                    File.WriteAllText(smlPath, sml, encoding)
                    successCount += 1
                Catch ex As Exception
                    failCount += 1
                    failFiles.Add(Path.GetFileName(csvPath) & ": " & ex.Message)
                End Try
            Next
            Dim msg = $"変換完了: {successCount}件 成功, {failCount}件 失敗"
            If failCount > 0 Then
                msg &= vbCrLf & String.Join(vbCrLf, failFiles)
            End If
            System.Windows.MessageBox.Show(msg, "完了", MessageBoxButton.OK, If(failCount = 0, MessageBoxImage.Information, MessageBoxImage.Warning))
        End Sub

        ' SML→CSV変換ボタン
        Private Sub BtnSmlToCsv_Click(sender As Object, e As RoutedEventArgs) Handles BtnSmlToCsv.Click
            Dim smlFiles As New List(Of String)()
            If RBFile.IsChecked Then
                Dim openFileDialog As New Microsoft.Win32.OpenFileDialog With {
                    .Filter = "SMLファイル (*.sml)|*.sml",
                    .Multiselect = True
                }
                If openFileDialog.ShowDialog() = True Then
                    smlFiles.AddRange(openFileDialog.FileNames)
                Else
                    Return
                End If
            ElseIf RBFolder.IsChecked Then
                Dim folderDialog As New System.Windows.Forms.FolderBrowserDialog With {
                    .SelectedPath = lastFolderPath
                }
                Dim result = folderDialog.ShowDialog()
                If result = System.Windows.Forms.DialogResult.OK Then
                    Dim folder = folderDialog.SelectedPath
                    lastFolderPath = folder
                    smlFiles.AddRange(Directory.GetFiles(folder, "*.sml"))
                    If smlFiles.Count = 0 Then
                        System.Windows.MessageBox.Show("SMLファイルが見つかりません。", "情報", MessageBoxButton.OK, MessageBoxImage.Information)
                        Return
                    End If
                Else
                    Return
                End If
            Else
                System.Windows.MessageBox.Show("ファイル/フォルダ選択方式を選んでください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error)
                Return
            End If

            Dim encoding = GetSelectedEncoding()
            Dim successCount As Integer = 0
            Dim failCount As Integer = 0
            Dim failFiles As New List(Of String)()
            For Each smlPath In smlFiles
                Dim csvPath = Path.ChangeExtension(smlPath, ".csv")
                Try
                    Dim lines = File.ReadAllLines(smlPath, Encoding.UTF8)
                    Dim csv = SmlToCsv(lines)
                    File.WriteAllText(csvPath, csv, encoding)
                    successCount += 1
                Catch ex As Exception
                    failCount += 1
                    failFiles.Add(Path.GetFileName(smlPath) & ": " & ex.Message)
                End Try
            Next
            Dim msg = $"変換完了: {successCount}件 成功, {failCount}件 失敗"
            If failCount > 0 Then
                msg &= vbCrLf & String.Join(vbCrLf, failFiles)
            End If
            System.Windows.MessageBox.Show(msg, "完了", MessageBoxButton.OK, If(failCount = 0, MessageBoxImage.Information, MessageBoxImage.Warning))
        End Sub

        ' SML→CSV変換本体
        Private Function SmlToCsv(lines As String()) As String
            Dim csv As New List(Of String)()
            If lines.Length = 0 Then Return String.Empty
            ' 1行目はSxFy(W)ヘッダー
            Dim header = lines(0).Trim()
            Dim csvHeader As String
            If header.StartsWith("S") AndAlso header.Contains("F") Then
                Dim sIdx = header.IndexOf("S") + 1
                Dim fIdx = header.IndexOf("F")
                Dim sNum = header.Substring(sIdx, fIdx - sIdx)
                Dim rest = header.Substring(fIdx + 1)
                Dim wBit = ""
                Dim fNum = rest
                If rest.Contains(" ") Then
                    fNum = rest.Substring(0, rest.IndexOf(" "))
                    wBit = rest.Substring(rest.IndexOf(" ") + 1)
                End If

                If ChkWD.IsChecked AndAlso wBit = "W" Then
                    wBit = ""
                End If

                csvHeader = sNum & "," & fNum & If(wBit <> "", "," & wBit, "")
            Else
                csvHeader = header
            End If
            csv.Add(csvHeader)
            ' 2行目以降をパース
            Dim idx As Integer = 1
            ParseSmlAny(lines, idx, 0, csv)
            Return String.Join(vbCrLf, csv)
        End Function

        ' SML本体の再帰パース
        Private Sub ParseSmlAny(lines As String(), ByRef idx As Integer, indentLevel As Integer, csv As List(Of String))
            While idx < lines.Length
                Dim line = lines(idx).Trim()
                If line = "." OrElse line = "" Then Exit Sub
                If line.StartsWith("<L [") Then
                    ParseSmlList(lines, idx, indentLevel, csv)
                ElseIf line.StartsWith("<") Then
                    ParseSmlValue(lines, idx, indentLevel, csv)
                ElseIf line = ">" OrElse line = ">." Then
                    idx += 1
                    Exit Sub
                Else
                    idx += 1 ' 無視
                End If
            End While
        End Sub

        Private Sub ParseSmlValue(lines As String(), ByRef idx As Integer, indentLevel As Integer, csv As List(Of String))
            Dim line = lines(idx).Trim()
            Dim typeEnd = line.IndexOf(" ")
            Dim typeName = line.Substring(1, typeEnd - 1)
            Dim valueStart = typeEnd + 1
            Dim valueEnd = line.LastIndexOf(">")
            Dim value = line.Substring(valueStart, valueEnd - valueStart)
            If typeName = "A" Then
                value = value.Trim()
                If value.StartsWith("""") AndAlso value.EndsWith("""") And value.Length >= 2 Then
                    value = value.Substring(1, value.Length - 2)
                End If
            End If
            csv.Add(typeName & "," & value)
            idx += 1
        End Sub

        Private Sub ParseSmlList(lines As String(), ByRef idx As Integer, indentLevel As Integer, csv As List(Of String))
            Dim line = lines(idx).Trim()
            Dim countStr = line.Substring(4, line.IndexOf("]") - 4)
            Dim count = Integer.Parse(countStr)
            csv.Add("L," & count)
            idx += 1
            For i = 1 To count
                ParseSmlAny(lines, idx, indentLevel + 1, csv)
            Next
            If idx < lines.Length AndAlso (lines(idx).Trim() = ">" OrElse lines(idx).Trim() = ">.") Then
                idx += 1
            End If
        End Sub

        ' 例: 6,11 → S6F11, 6,11,W → S6F11 W
        Private Function ConvertHeader(headerLine As String) As String
            Dim parts = headerLine.Split(","c)
            If parts.Length >= 2 AndAlso IsNumeric(parts(0)) AndAlso IsNumeric(parts(1)) Then
                Dim result = $"S{parts(0)}F{parts(1)}"
                If parts.Length >= 3 AndAlso Not String.IsNullOrWhiteSpace(parts(2)) Then
                    result &= " " & parts(2).Trim()
                ElseIf ChkWA.IsChecked Then
                    Dim num As Integer
                    If Integer.TryParse(parts(1), num) AndAlso num Mod 2 = 1 Then
                        result &= " W"
                    End If
                End If
                Return result
            Else
                Return headerLine ' 変換できなければそのまま
            End If
        End Function

        ' リストまたは値型をパース（トップレベル対応）
        Private Function ParseAny(lines As String(), ByRef index As Integer, indentLevel As Integer, isTopLevel As Boolean) As String
            If index >= lines.Length Then Return String.Empty
            Dim parts = lines(index).Split(","c)
            If parts(0) = "L" Then
                Return ParseList(lines, index, indentLevel, isTopLevel)
            ElseIf SupportedTypes.Contains(parts(0)) Then
                Dim result = ParseValue(parts, indentLevel)
                index += 1
                Return result
            Else
                Throw New Exception($"未知の型: {parts(0)}")
            End If
        End Function

        ' 再帰的にリストをパース
        Private Function ParseList(lines As String(), ByRef index As Integer, indentLevel As Integer, isTopLevel As Boolean) As String
            If index >= lines.Length Then Return String.Empty
            Dim header = lines(index).Split(","c)
            If header(0) <> "L" Then Throw New Exception($"リスト開始が不正: {lines(index)}")
            Dim count = Integer.Parse(header(1))
            index += 1
            Dim indent = New String(" "c, indentLevel * 4)
            If count = 0 Then
                Return indent & "<L [0]>" & vbCrLf
            End If
            Dim sb As New StringBuilder()
            sb.Append(indent & "<L [" & count & "]" & vbCrLf)
            For i = 1 To count
                If index >= lines.Length Then Exit For
                sb.Append(ParseAny(lines, index, indentLevel + 1, False))
            Next
            sb.Append(indent & ">" & vbCrLf)
            Return sb.ToString()
        End Function

        Private Function ParseValue(parts As String(), indentLevel As Integer) As String
            Dim indent = New String(" "c, indentLevel * 4)
            Dim typeName = parts(0)
            If typeName = "A" Then
                Return indent & $"<A ""{parts(1)}"">{vbCrLf}"
            Else
                Return indent & $"<{typeName} {parts(1)}>{vbCrLf}"
            End If
        End Function
    End Class
End Namespace
