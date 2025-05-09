Public Class Form1

    Private ReadOnly Port(3) As IO.Ports.SerialPort
    Private COMUpDown As NumericUpDown()
    Private ZeroOffset As Double()()        ' ゼロオフセット値（AD値）
    Private Prf As Double()()               ' アンプ校正値（με/AD値）
    Private Npu As Double()()               ' センサ校正値（N/με）
    Private Len As Double()()               ' センサ距離（m）

    Private Structure FP
        Public F1 As Double
        Public F2 As Double
        Public F3 As Double
        Public F4 As Double
        Public COPx As Double
        Public COPy As Double
    End Structure

    Private Sub StartButton_Click(sender As Object, e As EventArgs) Handles StartButton.Click
        ' COMポートの設定
        ' アンプは4台で、1台目からFP1～4、FP5～8、FP9～12、FP13～15（4台目だけは3台接続）
        Dim p(3) As IO.Ports.SerialPort
        For i = 0 To 3
            p(i) = New IO.Ports.SerialPort
            With p(i)
                .PortName = "COM" + COMUpDown(i).Value.ToString
                .BaudRate = 1000000                             ' 1Mbps
                .DataBits = 8                                   ' 8bit
                .Parity = IO.Ports.Parity.None                  ' ノーパリティ
                .StopBits = IO.Ports.StopBits.One               ' 1ストップビット
                .Handshake = IO.Ports.Handshake.RequestToSend   ' CTS/RTS
                Try
                    .Open()                                     ' ポートオープン
                Catch ex As Exception

                End Try
            End With
        Next

        ' シリアル番号の取得
        Dim serial(3) As String
        For i = 0 To 3
            If p(i).IsOpen Then
                serial(i) = GetSerial(p(i))
            Else
                serial(i) = ""
            End If
        Next

        ' シリアル番号"ARS-0001"が1台目、"ARS-0002"が2台目、"ARS-0003"が3台目、"ARS-0004"が4台目
        For i = 0 To 3
            Select Case serial(i)
                Case "ARS-0001"
                    Port(0) = p(i)
                Case "ARS-0002"
                    Port(1) = p(i)
                Case "ARS-0003"
                    Port(2) = p(i)
                Case "ARS-0004"
                    Port(3) = p(i)
            End Select
        Next

        ' 必要に応じてゼロ調整を行う
        ' ゼロ調整結果は電源OFFしても保存されているので必須ではない
        For i = 0 To 3
            If Port(i) IsNot Nothing Then
                Zero(Port(i))
            End If
        Next

        ' ゼロオフセット値を取得する
        ' ゼロ調整を行ったら必ず取得する必要がある
        ReDim ZeroOffset(3)
        For i = 0 To 3
            If Port(i) IsNot Nothing Then
                ZeroOffset(i) = GetZeroOffset(Port(i))
            End If
        Next

        ' アンプ校正値（με/AD値）を取得する
        ' （AD値をひずみに変換するための値）
        ' ※工場出荷後は変化しないので、繰り返し取得する必要はない
        ReDim Prf(3)
        For i = 0 To 3
            If Port(i) IsNot Nothing Then
                Prf(i) = GetPrf(Port(i))
            End If
        Next

        ' センサ校正値（N/με）を取得する
        ' （ひずみを力に変換するための値）
        ' ※工場出荷後は変化しないので、繰り返し取得する必要はない
        ReDim Npu(3)
        For i = 0 To 3
            If Port(i) IsNot Nothing Then
                Npu(i) = GetNpu(Port(i))
            End If
        Next

        ' センサ距離（m）を取得する
        ' （力からCOPを算出するために必要な値）
        ' ※工場出荷後は変化しないので、繰り返し取得する必要はない
        ReDim Len(3)
        For i = 0 To 3
            If Port(i) IsNot Nothing Then
                Len(i) = GetLen(Port(i))
            End If
        Next

        '--------------------------------------------------------
        ' ここまでで必要な情報がすべて揃ったので、測定を開始する
        '--------------------------------------------------------

        ' 4台のアンプすべてを外部トリガ待ち受け状態に移行させる
        For i = 0 To 3
            If Port(i) IsNot Nothing Then Ext(Port(i))
        Next

        ' 4台のアンプすべてを外部トリガ待ち受け状態に移行させる
        For i = 0 To 3
            If Port(i) IsNot Nothing Then Ext(Port(i))
        Next

        ' 4台のアンプのうち1台からトリガ信号を出力する
        For i = 0 To 3
            If Port(i) IsNot Nothing Then
                Trig(Port(i))
                Exit For
            End If
        Next

        '----------------------------------------------------------------------------------
        ' これで4台とも計測状態となり、測定データを送ってくる
        ' サンプリング周波数100Hzで、10回分（100ms分）データが溜まったらまとめて送ってくる
        ' アンプ1台（フォースプレート4台分）の1回あたりのデータは、
        '  16(ch) x 2(byte) x 10(回) + 8(ヘッダ4byte+フッタ4byte)
        '----------------------------------------------------------------------------------
        Dim n As Integer = 0
        Do While n < 100   ' 100回（10秒）繰り返す
            '-----------------------------------------
            '           バイナリデータ受信
            '-----------------------------------------
            Dim bin As Byte()()
            ReDim bin(3)
            For i = 0 To 3
                If Port(i) Is Nothing Then Continue For  ' 未接続の場合は飛ばす

                ' データ受信を待つ
                Do Until 16 * 2 * 10 + 8 <= Port(i).BytesToRead
                    Threading.Thread.Sleep(10)
                Loop
                ReDim bin(i)(16 * 2 * 10 + 8 - 1)
                Port(i).Read(bin(i), 0, bin(i).Length)
            Next

            '-----------------------------------------
            '       バイナリデータをAD値に変換
            '-----------------------------------------
            Dim ad As Integer()(,)
            ReDim ad(3)
            For i = 0 To 3
                If bin(i) Is Nothing Then Continue For   ' 未接続の場合は飛ばす

                ReDim ad(i)(15, 9)  ' 16(ch) x 10(回)

                ' 最初の4byteは0xAA、最後の4byteは0x55
                If bin(i)(0) <> &HAA OrElse bin(i)(1) <> &HAA OrElse bin(i)(2) <> &HAA OrElse bin(i)(3) <> &HAA OrElse
                   bin(i)(bin(i).Length - 1) <> &H55 OrElse bin(i)(bin(i).Length - 2) <> &H55 OrElse bin(i)(bin(i).Length - 3) <> &H55 OrElse bin(i)(bin(i).Length - 4) <> &H55 Then
                    Continue For
                End If

                Using br As New IO.BinaryReader(New IO.MemoryStream(bin(i)))
                    br.ReadInt32()  ' 最初の4byteを飛ばす
                    For m = 0 To 9
                        For ch = 0 To 15
                            ad(i)(ch, m) = br.ReadInt16 ' 符号付き16bit整数に変換する
                        Next
                    Next
                End Using
            Next

            '-----------------------------------------
            '             AD値を力に変換
            '-----------------------------------------
            Dim data As FP()()()
            ReDim data(3)
            For i = 0 To 3
                If ad(i) Is Nothing Then Continue For    ' 未接続の場合は飛ばす

                ReDim data(i)(3)
                For j = 0 To 3
                    ReDim data(i)(j)(9)
                Next
                For m = 0 To 9
                    ' F1～F4を算出
                    ' F1 : 左前
                    ' F2 : 右前
                    ' F3 : 右後
                    ' F4 : 左後
                    Dim fp1 As FP
                    fp1.F1 = (ad(i)(0, m) - ZeroOffset(i)(0)) * Prf(i)(0) * Npu(i)(0)    ' (AD値 - ゼロオフセット値) x アンプ校正値 x センサ校正値 = 力[N]
                    fp1.F2 = (ad(i)(1, m) - ZeroOffset(i)(1)) * Prf(i)(1) * Npu(i)(1)
                    fp1.F3 = (ad(i)(2, m) - ZeroOffset(i)(2)) * Prf(i)(2) * Npu(i)(2)
                    fp1.F4 = (ad(i)(3, m) - ZeroOffset(i)(3)) * Prf(i)(3) * Npu(i)(3)

                    Dim fp2 As FP
                    fp2.F1 = (ad(i)(4, m) - ZeroOffset(i)(4)) * Prf(i)(4) * Npu(i)(4)
                    fp2.F2 = (ad(i)(5, m) - ZeroOffset(i)(5)) * Prf(i)(5) * Npu(i)(5)
                    fp2.F3 = (ad(i)(6, m) - ZeroOffset(i)(6)) * Prf(i)(6) * Npu(i)(6)
                    fp2.F4 = (ad(i)(7, m) - ZeroOffset(i)(7)) * Prf(i)(7) * Npu(i)(7)

                    Dim fp3 As FP
                    fp3.F1 = (ad(i)(8, m) - ZeroOffset(i)(8)) * Prf(i)(8) * Npu(i)(8)
                    fp3.F2 = (ad(i)(9, m) - ZeroOffset(i)(9)) * Prf(i)(9) * Npu(i)(9)
                    fp3.F3 = (ad(i)(10, m) - ZeroOffset(i)(10)) * Prf(i)(10) * Npu(i)(10)
                    fp3.F4 = (ad(i)(11, m) - ZeroOffset(i)(11)) * Prf(i)(11) * Npu(i)(11)

                    Dim fp4 As FP
                    fp4.F1 = (ad(i)(12, m) - ZeroOffset(i)(12)) * Prf(i)(12) * Npu(i)(12)
                    fp4.F2 = (ad(i)(13, m) - ZeroOffset(i)(13)) * Prf(i)(13) * Npu(i)(13)
                    fp4.F3 = (ad(i)(14, m) - ZeroOffset(i)(14)) * Prf(i)(14) * Npu(i)(14)
                    fp4.F4 = (ad(i)(15, m) - ZeroOffset(i)(15)) * Prf(i)(15) * Npu(i)(15)

                    ' COPx、COPyを算出
                    ' 左が+X、前が+Y
                    Dim Fz1 As Double = fp1.F1 + fp1.F2 + fp1.F3 + fp1.F4
                    Dim Mx1 As Double = (fp1.F1 + fp1.F2 - fp1.F3 - fp1.F4) * Len(i)(0)
                    Dim My1 As Double = (fp1.F2 + fp1.F3 - fp1.F1 - fp1.F4) * Len(i)(0)
                    If 50 < Fz1 Then
                        fp1.COPx = -My1 / Fz1
                        fp1.COPy = Mx1 / Fz1
                    Else
                        fp1.COPx = 0
                        fp1.COPy = 0
                    End If

                    Dim Fz2 As Double = fp2.F1 + fp2.F2 + fp2.F3 + fp2.F4
                    Dim Mx2 As Double = (fp2.F1 + fp2.F2 - fp2.F3 - fp2.F4) * Len(i)(1)
                    Dim My2 As Double = (fp2.F2 + fp2.F3 - fp2.F1 - fp2.F4) * Len(i)(1)
                    If 50 < Fz2 Then
                        fp2.COPx = -My2 / Fz2
                        fp2.COPy = Mx2 / Fz2
                    Else
                        fp2.COPx = 0
                        fp2.COPy = 0
                    End If

                    Dim Fz3 As Double = fp3.F1 + fp3.F2 + fp3.F3 + fp3.F4
                    Dim Mx3 As Double = (fp3.F1 + fp3.F2 - fp3.F3 - fp3.F4) * Len(i)(2)
                    Dim My3 As Double = (fp3.F2 + fp3.F3 - fp3.F1 - fp3.F4) * Len(i)(2)
                    If 50 < Fz3 Then
                        fp3.COPx = -My3 / Fz3
                        fp3.COPy = Mx3 / Fz3
                    Else
                        fp3.COPx = 0
                        fp3.COPy = 0
                    End If

                    Dim Fz4 As Double = fp4.F1 + fp4.F2 + fp4.F3 + fp4.F4
                    Dim Mx4 As Double = (fp4.F1 + fp4.F2 - fp4.F3 - fp4.F4) * Len(i)(3)
                    Dim My4 As Double = (fp4.F2 + fp4.F3 - fp4.F1 - fp4.F4) * Len(i)(3)
                    If 50 < Fz4 Then
                        fp4.COPx = -My4 / Fz4
                        fp4.COPy = Mx4 / Fz4
                    Else
                        fp4.COPx = 0
                        fp4.COPy = 0
                    End If

                    data(i)(0)(m) = fp1
                    data(i)(1)(m) = fp2
                    data(i)(2)(m) = fp3
                    data(i)(3)(m) = fp4
                Next

                '-----------------------------------------
                '                  表示
                '-----------------------------------------
                If data(0) IsNot Nothing Then    ' 未接続の場合は飛ばす
                    ' FP1
                    DataView(0, 0).Value = data(0)(0)(0).F1
                    DataView(1, 0).Value = data(0)(0)(0).F2
                    DataView(2, 0).Value = data(0)(0)(0).F3
                    DataView(3, 0).Value = data(0)(0)(0).F4
                    DataView(4, 0).Value = data(0)(0)(0).COPx * 1000    ' mm表示
                    DataView(5, 0).Value = data(0)(0)(0).COPy * 1000    ' mm表示

                    ' FP2
                    DataView(0, 1).Value = data(0)(1)(0).F1
                    DataView(1, 1).Value = data(0)(1)(0).F2
                    DataView(2, 1).Value = data(0)(1)(0).F3
                    DataView(3, 1).Value = data(0)(1)(0).F4
                    DataView(4, 1).Value = data(0)(1)(0).COPx * 1000
                    DataView(5, 1).Value = data(0)(1)(0).COPy * 1000

                    ' FP3
                    DataView(0, 2).Value = data(0)(2)(0).F1
                    DataView(1, 2).Value = data(0)(2)(0).F2
                    DataView(2, 2).Value = data(0)(2)(0).F3
                    DataView(3, 2).Value = data(0)(2)(0).F4
                    DataView(4, 2).Value = data(0)(2)(0).COPx * 1000
                    DataView(5, 2).Value = data(0)(2)(0).COPy * 1000

                    ' FP4
                    DataView(0, 3).Value = data(0)(3)(0).F1
                    DataView(1, 3).Value = data(0)(3)(0).F2
                    DataView(2, 3).Value = data(0)(3)(0).F3
                    DataView(3, 3).Value = data(0)(3)(0).F4
                    DataView(4, 3).Value = data(0)(3)(0).COPx * 1000
                    DataView(5, 3).Value = data(0)(3)(0).COPy * 1000
                End If

                If data(1) IsNot Nothing Then    ' 未接続の場合は飛ばす
                    ' FP5
                    DataView(0, 4).Value = data(1)(0)(0).F1
                    DataView(1, 4).Value = data(1)(0)(0).F2
                    DataView(2, 4).Value = data(1)(0)(0).F3
                    DataView(3, 4).Value = data(1)(0)(0).F4
                    DataView(4, 4).Value = data(1)(0)(0).COPx * 1000
                    DataView(5, 4).Value = data(1)(0)(0).COPy * 1000

                    ' FP6
                    DataView(0, 5).Value = data(1)(1)(0).F1
                    DataView(1, 5).Value = data(1)(1)(0).F2
                    DataView(2, 5).Value = data(1)(1)(0).F3
                    DataView(3, 5).Value = data(1)(1)(0).F4
                    DataView(4, 5).Value = data(1)(1)(0).COPx * 1000
                    DataView(5, 5).Value = data(1)(1)(0).COPy * 1000

                    ' FP7
                    DataView(0, 6).Value = data(1)(2)(0).F1
                    DataView(1, 6).Value = data(1)(2)(0).F2
                    DataView(2, 6).Value = data(1)(2)(0).F3
                    DataView(3, 6).Value = data(1)(2)(0).F4
                    DataView(4, 6).Value = data(1)(2)(0).COPx * 1000
                    DataView(5, 6).Value = data(1)(2)(0).COPy * 1000

                    ' FP8
                    DataView(0, 7).Value = data(1)(3)(0).F1
                    DataView(1, 7).Value = data(1)(3)(0).F2
                    DataView(2, 7).Value = data(1)(3)(0).F3
                    DataView(3, 7).Value = data(1)(3)(0).F4
                    DataView(4, 7).Value = data(1)(3)(0).COPx * 1000
                    DataView(5, 7).Value = data(1)(3)(0).COPy * 1000
                End If

                If data(2) IsNot Nothing Then    ' 未接続の場合は飛ばす
                    ' FP9
                    DataView(0, 8).Value = data(2)(0)(0).F1
                    DataView(1, 8).Value = data(2)(0)(0).F2
                    DataView(2, 8).Value = data(2)(0)(0).F3
                    DataView(3, 8).Value = data(2)(0)(0).F4
                    DataView(4, 8).Value = data(2)(0)(0).COPx * 1000
                    DataView(5, 8).Value = data(2)(0)(0).COPy * 1000

                    ' FP10
                    DataView(0, 9).Value = data(2)(1)(0).F1
                    DataView(1, 9).Value = data(2)(1)(0).F2
                    DataView(2, 9).Value = data(2)(1)(0).F3
                    DataView(3, 9).Value = data(2)(1)(0).F4
                    DataView(4, 9).Value = data(2)(1)(0).COPx * 1000
                    DataView(5, 9).Value = data(2)(1)(0).COPy * 1000

                    ' FP11
                    DataView(0, 10).Value = data(2)(2)(0).F1
                    DataView(1, 10).Value = data(2)(2)(0).F2
                    DataView(2, 10).Value = data(2)(2)(0).F3
                    DataView(3, 10).Value = data(2)(2)(0).F4
                    DataView(4, 10).Value = data(2)(2)(0).COPx * 1000
                    DataView(5, 10).Value = data(2)(2)(0).COPy * 1000

                    ' FP12
                    DataView(0, 11).Value = data(2)(3)(0).F1
                    DataView(1, 11).Value = data(2)(3)(0).F2
                    DataView(2, 11).Value = data(2)(3)(0).F3
                    DataView(3, 11).Value = data(2)(3)(0).F4
                    DataView(4, 11).Value = data(2)(3)(0).COPx * 1000
                    DataView(5, 11).Value = data(2)(3)(0).COPy * 1000
                End If

                If data(3) IsNot Nothing Then    ' 未接続の場合は飛ばす
                    ' FP13
                    DataView(0, 12).Value = data(3)(0)(0).F1
                    DataView(1, 12).Value = data(3)(0)(0).F2
                    DataView(2, 12).Value = data(3)(0)(0).F3
                    DataView(3, 12).Value = data(3)(0)(0).F4
                    DataView(4, 12).Value = data(3)(0)(0).COPx * 1000
                    DataView(5, 12).Value = data(3)(0)(0).COPy * 1000

                    ' FP14
                    DataView(0, 13).Value = data(3)(1)(0).F1
                    DataView(1, 13).Value = data(3)(1)(0).F2
                    DataView(2, 13).Value = data(3)(1)(0).F3
                    DataView(3, 13).Value = data(3)(1)(0).F4
                    DataView(4, 13).Value = data(3)(1)(0).COPx * 1000
                    DataView(5, 13).Value = data(3)(1)(0).COPy * 1000

                    ' FP15
                    DataView(0, 14).Value = data(3)(2)(0).F1
                    DataView(1, 14).Value = data(3)(2)(0).F2
                    DataView(2, 14).Value = data(3)(2)(0).F3
                    DataView(3, 14).Value = data(3)(2)(0).F4
                    DataView(4, 14).Value = data(3)(2)(0).COPx * 1000
                    DataView(5, 14).Value = data(3)(2)(0).COPy * 1000

                    ' FP16は存在しない
                End If

                My.Application.DoEvents()
            Next
            n += 1
        Loop

        ' 4台のアンプすべてを停止させる
        For i = 0 To 3
            If Port(i) IsNot Nothing Then StopMeasure(Port(i))
        Next

        ' 次の測定のためにポートを閉じる
        ' ※このサンプル特有の動作であり、実際のソフトウェア作成時は必要があればゼロ調整、なければ外部トリガ待ち受けへの移行へ
        For i = 0 To 3
            If Port(i) IsNot Nothing Then
                If Port(i).IsOpen Then Port(i).Close()
                Port(i) = Nothing
            End If
        Next
    End Sub

    Private Function GetSerial(p As IO.Ports.SerialPort) As String
        Dim reply As String = ""
        Dim success As Boolean = GetReply(p, "GET_SERIAL" + vbLf, reply)    ' シリアル番号を取得するコマンドは、"GET_SERIAL" + \n(vbLf)で、応答は、"SERIAL_xxxxx" + \n
        If success Then
            If reply.StartsWith("SERIAL_") Then reply = reply.Substring("SERIAL_".Length)       ' 応答が、"SERIAL_"で始まる場合はその部分を削除
            If reply.EndsWith(vbLf) Then reply = reply.Substring(0, reply.Length - vbLf.Length) ' 応答が、\nで終わる場合はその部分を削除
            Return reply
        Else
            Return ""
        End If
    End Function

    Private Function Zero(p As IO.Ports.SerialPort) As Boolean
        Dim reply As String = ""
        Dim success As Boolean = GetReply(p, "ZERO" + vbLf, reply)    ' ゼロ調整を行うコマンドは、"ZERO" + \n(vbLf)で、応答は、"ZERO_OK" + \n
        If success AndAlso reply = "ZERO_OK" + vbLf Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function GetZeroOffset(p As IO.Ports.SerialPort) As Double()
        Dim reply As String = ""
        Dim success As Boolean = GetReply(p, "GET_ZERO" + vbLf, reply)    ' ゼロオフセット値を取得するコマンドは、"GET_ZERO" + \n(vbLf)で、応答は、"ZERO_x, ... ,x" + \n
        If success Then
            If reply.StartsWith("ZERO_") Then reply = reply.Substring("ZERO_".Length)           ' 応答が、"ZERO_"で始まる場合はその部分を削除
            If reply.EndsWith(vbLf) Then reply = reply.Substring(0, reply.Length - vbLf.Length) ' 応答が、\nで終わる場合はその部分を削除
            Dim values(15) As Double
            Dim s As String() = reply.Split(","c)   ' カンマ区切りで16個の浮動小数点で送られてくる
            For i = 0 To 15
                Dim value As Double
                If i < s.Length AndAlso Double.TryParse(s(i), value) Then
                    values(i) = value
                Else
                    values(i) = 0
                End If
            Next
            Return values
        Else
            Return New Double() {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
        End If
    End Function

    Private Function GetPrf(p As IO.Ports.SerialPort) As Double()
        Dim reply As String = ""
        Dim success As Boolean = GetReply(p, "GET_PRF" + vbLf, reply)       ' アンプ校正値を取得するコマンドは、"GET_PRF" + \n(vbLf)で、応答は、"PRF_x, ... ,x" + \n
        If success Then
            If reply.StartsWith("PRF_") Then reply = reply.Substring("PRF_".Length)             ' 応答が、"PRF_"で始まる場合はその部分を削除
            If reply.EndsWith(vbLf) Then reply = reply.Substring(0, reply.Length - vbLf.Length) ' 応答が、\nで終わる場合はその部分を削除
            Dim values(15) As Double
            Dim s As String() = reply.Split(","c)   ' カンマ区切りで16個の浮動小数点で送られてくる
            For i = 0 To 15
                Dim value As Double
                If i < s.Length AndAlso Double.TryParse(s(i), value) Then
                    values(i) = value
                Else
                    values(i) = 0
                End If
            Next
            Return values
        Else
            Return New Double() {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
        End If
    End Function

    Private Function GetNpu(p As IO.Ports.SerialPort) As Double()
        Dim reply As String = ""
        Dim success As Boolean = GetReply(p, "GET_NPU" + vbLf, reply)       ' センサ校正値を取得するコマンドは、"GET_NPU" + \n(vbLf)で、応答は、"NPU_x, ... ,x" + \n
        If success Then
            If reply.StartsWith("NPU_") Then reply = reply.Substring("NPU_".Length)             ' 応答が、"NPU_"で始まる場合はその部分を削除
            If reply.EndsWith(vbLf) Then reply = reply.Substring(0, reply.Length - vbLf.Length) ' 応答が、\nで終わる場合はその部分を削除
            Dim values(15) As Double
            Dim s As String() = reply.Split(","c)   ' カンマ区切りで16個の浮動小数点で送られてくる
            For i = 0 To 15
                Dim value As Double
                If i < s.Length AndAlso Double.TryParse(s(i), value) Then
                    values(i) = value
                Else
                    values(i) = 0
                End If
            Next
            Return values
        Else
            Return New Double() {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
        End If
    End Function

    Private Function GetLen(p As IO.Ports.SerialPort) As Double()
        Dim reply As String = ""
        Dim success As Boolean = GetReply(p, "GET_LEN" + vbLf, reply)       ' センサ距離を取得するコマンドは、"GET_LEN" + \n(vbLf)で、応答は、"LEN_x,x,x,x" + \n
        If success Then
            If reply.StartsWith("LEN_") Then reply = reply.Substring("LEN_".Length)             ' 応答が、"LEN_"で始まる場合はその部分を削除
            If reply.EndsWith(vbLf) Then reply = reply.Substring(0, reply.Length - vbLf.Length) ' 応答が、\nで終わる場合はその部分を削除
            Dim values(3) As Double
            Dim s As String() = reply.Split(","c)   ' カンマ区切りで4個の浮動小数点で送られてくる
            For i = 0 To 3
                Dim value As Double
                If i < s.Length AndAlso Double.TryParse(s(i), value) Then
                    values(i) = value
                Else
                    values(i) = 0
                End If
            Next
            Return values
        Else
            Return New Double() {0, 0, 0, 0}
        End If
    End Function

    Private Function Ext(p As IO.Ports.SerialPort) As Boolean
        Dim reply As String = ""
        Dim success As Boolean = GetReply(p, "EXT" + vbLf, reply)   ' トリガ待ち受け状態に移行させるコマンドは、"EXT" + \n(vbLf)で、応答は、"EXT_OK" + \n
        If success AndAlso reply = "EXT_OK" + vbLf Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function Trig(p As IO.Ports.SerialPort) As Boolean
        Dim reply As String = ""
        Dim success As Boolean = GetReply(p, "TRIG" + vbLf, reply)  ' トリガ出力するコマンドは、"TRIG" + \n(vbLf)で、応答は、"TRIG_OK" + \n
        If success AndAlso reply = "TRIG_OK" + vbLf Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub StopMeasure(p As IO.Ports.SerialPort)
        p.Write("STOP" + vbLf)    ' 測定を停止するコマンドは、"STOP" + \n(vbLf)で、応答は、"STOP_OK" + \nであるが、測定データが送られてくるので応答を確認しない

        ' 1秒待って受信バッファをクリアする
        Threading.Thread.Sleep(1000)
        p.DiscardInBuffer()
    End Sub

    ''' <summary>
    ''' コマンド応答文字列の取得
    ''' </summary>
    ''' <param name="command">コマンド</param>
    ''' <param name="reply">応答文字列</param>
    ''' <param name="t1">無応答と判断するまでの時間[ms]</param>
    ''' <param name="t2">後続データなしと判断するまでの時間[ms]</param>
    ''' <returns></returns>
    Private Function GetReply(p As IO.Ports.SerialPort, command As String, ByRef reply As String, Optional t1 As Integer = 2000, Optional t2 As Integer = 100) As Boolean
        reply = ""
        p.DiscardInBuffer()
        p.Write(command)
        Dim sw As Stopwatch = Stopwatch.StartNew
        Do While p.BytesToRead = 0
            If t1 <= sw.ElapsedMilliseconds Then    ' t1[ms](標準では2000ms)待っても1byteも応答がない場合はエラー
                Return False
            Else
                Threading.Thread.Sleep(0)
            End If
        Loop

        ' 1byteづつチェックして\nを探す(コマンドに対する応答は\n(vbLf)で終わる)
        Dim buf As New List(Of Byte)
        Dim isFound As Boolean = False
        Do While 0 < p.BytesToRead
            Dim b As Byte = CByte(p.ReadByte)
            buf.Add(b)
            If b = System.Text.Encoding.ASCII.GetBytes(vbLf)(0) Then
                isFound = True
                Exit Do
            End If
        Loop
        If isFound Then
            reply = System.Text.Encoding.ASCII.GetString(buf.ToArray)   ' \nが見つかった場合は、応答取得成功
            Return True
        End If

        ' 不足部分を受信
        sw.Restart()
        Do While sw.ElapsedMilliseconds < t2
            Do While 0 < p.BytesToRead
                Dim b As Byte = CByte(p.ReadByte)
                buf.Add(b)
                If b = System.Text.Encoding.ASCII.GetBytes(vbLf)(0) Then
                    reply = System.Text.Encoding.ASCII.GetString(buf.ToArray)   ' \nが見つかった場合は、応答取得成功
                    Return True
                End If
                sw.Restart()    ' 受信データがあれば、タイマーをリセットする
            Loop
        Loop

        Return False    ' \nが見つからなかった場合はエラー
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        COMUpDown = New NumericUpDown() {COMUpDown1, COMUpDown2, COMUpDown3, COMUpDown4}

        With DataView
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeColumns = False
            .AllowUserToResizeRows = False
            .ScrollBars = ScrollBars.None
            .BorderStyle = BorderStyle.None
            .ColumnHeadersVisible = False
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .DefaultCellStyle.WrapMode = DataGridViewTriState.False
            .DefaultCellStyle.BackColor = Color.White
            .DefaultCellStyle.Format = "N0" ' 表示は整数とする
            .GridColor = Color.DimGray
            .RowHeadersVisible = False
            .RowTemplate.Height = 22
            .ColumnCount = 6
            .Columns(0).Width = 79
            .Columns(1).Width = 79
            .Columns(2).Width = 79
            .Columns(3).Width = 79
            .Columns(4).Width = 79
            .Columns(5).Width = 79
            .RowCount = 15
        End With

    End Sub

End Class
