Module GetLineStationDatabase
    '駅と路線のリストを作成するためのモジュール　@author: hasega60
    Private Class Station
        ReadOnly Property station_id As Long
        ReadOnly Property station_code As Long
        ReadOnly Property station_name As String
        ReadOnly Property station_type As Integer
        ReadOnly Property corp_id As Integer
        ReadOnly Property line_id As Integer
        ReadOnly Property order As Integer
        ReadOnly Property lat As Double
        ReadOnly Property lon As Double

        Sub New(id As Long, code As Long, name As String, type As Integer, corpid As Integer, lineid As Integer, station_order As Integer, latitude As Double, longitude As Double)
            station_id = id
            station_code = code
            station_name = name
            station_type = type
            corp_id = corpid
            line_id = lineid
            order = station_order
            lat = latitude
            lon = longitude
        End Sub

    End Class
    
    Private Class Station_s
        ReadOnly Property station_code As Long
        ReadOnly Property station_name As String
        ReadOnly Property station_name_long As String
        ReadOnly Property station_type As Integer
        ReadOnly Property pass_line_count As Integer
        ReadOnly Property lat As Double
        ReadOnly Property lon As Double
        'station_code,station_name,station_name_long,station_type,pass_line_count,lat,lon
        Sub New(code As Long, name As String, name_long As String, type As Integer, line_count As Integer, latitude As Double, longitude As Double)
            station_code = code
            station_name = name
            station_name_long = name_long
            station_type = type
            pass_line_count = line_count
            lat = latitude
            lon = longitude
        End Sub
     End Class


    Private Class Line
        ReadOnly Property line_id As Long
        ReadOnly Property line_name As String
        ReadOnly Property line_name_long As String
        ReadOnly Property line_type As Integer
        ReadOnly Property station_count As Integer
        ReadOnly Property corp_id As Long
        ReadOnly Property corp_name As String

        Sub New(lineId As Long, lineName As String, lineName_L As String, lineType As Integer, sCount As Integer, corpId As Integer, corpName As String)
            line_id = lineId
            line_name = lineName
            line_name_long = lineName_L
            line_type = lineType
            station_count = sCount
            corp_id = corpId
            corp_name = corpName

        End Sub
    End Class

    'line_id,line_name,line_name_long,line_name_short,line_type,station_count,corp_id,corp_name,line_type
    Private Class Line_s
        ReadOnly Property line_id As Long
        ReadOnly Property line_name As String
        ReadOnly Property line_name_long As String
        ReadOnly Property line_name_short As String
        ReadOnly Property line_type As Integer
        ReadOnly Property station_count As Integer
        ReadOnly Property corp_id As Long
        ReadOnly Property corp_name As String

        Sub New(lineId As Long, lineName As String, lineName_L As String, lineName_S As String, lineType As Integer, sCount As Integer, corpId As Integer, corpName As String)
            line_id = lineId
            line_name = lineName
            line_name_long = lineName_L
            line_name_short = lineName_S
            line_type = lineType
            station_count = sCount
            corp_id = corpId
            corp_name = corpName

        End Sub
    End Class


    Public dataObj As EXPDENGNLib.ExpDataExtraction2
    Public data2 As EXPDENGNLib.ExpData2
    Public data_diaDB As EXPDENGNLib.ExpDiaDB10
    Public iSearch As EXPDENGNLib.ExpDiaSearch10
    Public iConvert As EXPDENGNLib.ExpConvert
    Public list_file_type = New List(Of String)()

    'データ出力用
    Public byte_limit As Integer = 5000000
    Public output_count_station As Integer = 0

    Public output_count_station_all As Integer = 0
    Public output_count_edge As Integer = 0
    Public output_count_fare As Integer = 0
    Public output_count_nearest_station As Integer = 0
    Public output_count_station_in_line As Integer = 0

    Public dict_added_corp_local_bus As New Dictionary(of Tuple(of Integer, Integer), List(Of Integer))

    Private Class Transfer
        ReadOnly Property station_code As Long
        ReadOnly Property line_name_long As String
        ReadOnly Property line_type As Integer


        Sub New(code As Long, name As String, type As Integer)
            station_code = code
            line_type = type
            line_name_long = name

        End Sub

    End Class
    Function GetPrivateProfileString(
       ByVal lpAppName As String,
       ByVal lpKeyName As String,
       ByVal lpDefault As String,
       ByVal lpReturnedString As System.Text.StringBuilder,
       ByVal nSize As Integer,
       ByVal lpFileName As String) As Integer
    End Function

    Function GetINIValue(ByRef Section As String,
                            ByRef KEY As String,
                            ByRef ININame As String) As String
        Try
            Dim Value As New System.Text.StringBuilder
            Call GetPrivateProfileString(Section, KEY, "Error", Value, 511, ININame)
            Return Left(Value.ToString(), InStr(1, Value.ToString(), vbNullChar) - 1)
        Catch ex As Exception
            Throw ex
        End Try
    End Function



    Private Function export_csv_part(ByRef csv As String, csv_base_path As String, ByRef file_count As Integer, max_byte As Integer, ts As Scripting.TextStream, fso As Scripting.FileSystemObject)
        If Len(csv) > max_byte Then
            export_csv(csv, csv_base_path, file_count, max_byte, ts, fso)

        End If

    End Function

    Private Function export_csv(ByRef csv As String, csv_base_path As String, ByRef file_count As Integer, max_byte As Integer, ts As Scripting.TextStream, fso As Scripting.FileSystemObject)
        ts = fso.OpenTextFile(csv_base_path + Str(file_count) + ".csv", Scripting.IOMode.ForWriting, True)
        ts.Write(csv)
        ts.Close()
        file_count = file_count + 1
        csv = ""
    End Function

    Sub get_station_list(path As String, export_path As String, data2 As EXPDENGNLib.ExpData2)
        Dim byte_limit As Integer = 1000000
        Dim sb_station As New System.Text.StringBuilder(byte_limit)
        Dim station As EXPDENGNLib.ExpDiaStation

        sb_station.Append("station_code,station_name,station_name_long,station_type,pass_line_count,lat,lon")
        sb_station.Append(vbCrLf)

        Dim station_code As Long
        Dim lat As Double
        Dim lon As Double

        Dim fso As Scripting.FileSystemObject
        fso = New Scripting.FileSystemObject
        Dim list_station_code As List(Of Integer) = New List(Of Integer)
        Dim row_str As String
        Dim row As String()
        Dim tr As Scripting.TextStream
        tr = fso.OpenTextFile(path, 1)
        Dim init As Boolean = True

        Dim csv_out = ""

        Console.CursorVisible = False
        Dim count = 0
        Do While tr.AtEndOfStream = False
            row_str = tr.ReadLine
            If init = False Then
                Try
                    count = count + 1
                    Console.Write("駅データ読み込み中: f" & Str(count))
                    Console.SetCursorPosition(0, Console.CursorTop)
                    row = Split(row_str, ",")
                    station_code = Integer.Parse(row(0))
                    If list_station_code.Contains(station_code) = False Then
                        list_station_code.Add(station_code)
                    End If

                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                End Try

            Else
                init = False '一行目のheader飛ばす
            End If
        Loop

        For k = 0 To list_station_code.Count - 1
            Console.Write("取得状況:" & Str(k) & " / " & Str(list_station_code.Count))
            Console.SetCursorPosition(0, Console.CursorTop)

            station_code = list_station_code(k)

            station = data2.GetStationByCode(station_code)
            ' 緯度経度が度、分、秒、1/100秒の配列でくるので小数点に変更 ちなみに日本測地系
            Dim latlon = convert_latlon(station)
            lat = latlon(0)
            lon = latlon(1)

            ' 駅のデータは重いので、StringBuilder使う形に変更
            sb_station.Append(Str(station_code))
            sb_station.Append(",")
            sb_station.Append(station.Name)
            sb_station.Append(",")
            sb_station.Append(station.LongName)
            sb_station.Append(",")
            sb_station.Append(Str(station.Type))
            sb_station.Append(",")
            sb_station.Append(Str(station.PassLineCount))
            sb_station.Append(",")
            sb_station.Append(Str(lat))
            sb_station.Append(",")
            sb_station.Append(Str(lon))
            sb_station.Append(vbCrLf)

        Next

        Dim ts As Scripting.TextStream

        csv_out = sb_station.ToString()
        csv_out = Replace(csv_out, " ", "")
        export_csv_part(csv_out, export_path, byte_limit, 0, ts, fso)

        sb_station.Clear()

        Console.CursorVisible = True


    End Sub

    Private function convert_latlon(station As EXPDENGNLib.ExpDiaStation)
        Dim latlon As New List(Of Double)
        Dim lat= station.Latitude(1) + (station.Latitude(2) * 60 + station.Latitude(3) + station.Latitude(4) / 100) / 3600
        Dim lon = station.Longitude(1) + (station.Longitude(2) * 60 + station.Longitude(3) + station.Longitude(4) / 100) / 3600

        latlon.add(lat)
        latlon.add(lon)

        Return latlon

    End Function

    Private function distance_latlon(latlon As List(Of Double), latlon_to As List(Of Double))
        Dim semiMajorAxis As Double = 6378137.0
        Dim flattening As Double = 1 / 298.257223563 
        Dim e_2 As Double = flattening * (2 - flattening)
        Dim degree As Double = 3.1415926535 / 180

        Dim y1 As Double = latlon(0)
        Dim x1 As Double = latlon(1)
        Dim y2 As Double = latlon_to(0)
        Dim x2 As Double = latlon_to(1)

        Dim coslat As Double = Math.Cos((y1 + y2) / 2 * degree)
        Dim w2 As Double = 1 / (1 - e_2 * (1 - coslat * coslat))
        Dim dx As Double = (x1 - x2) * coslat
        dim dy As Double = (y1 - y2) * w2 * (1 - e_2)
        Dim dist As Double = (Math.Sqrt((dx * dx + dy * dy) * w2) * semiMajorAxis * degree)/1000

        return dist

    End Function


    Private Function  create_fare_database(corp As EXPDENGNLib.ExpDiaCorp, corp_id As Integer,line_type As Integer,line_id As Integer, line_name As String,  dictStationCorp As Dictionary(Of Integer, EXPDENGNLib.ExpDiaStation),iSearch As EXPDENGNLib.ExpDiaSearch10,iNavi As EXPDENGNLib.ExpDiaNavi6, iRange As EXPDENGNLib.ExpRange, data2 As EXPDENGNLib.ExpData2, list_station_main As List(Of Integer),f_to_station As EXPDENGNLib.ExpDiaStation , base_dir As String, csv_fare As String,sb_fare As System.Text.StringBuilder)
        
        ' 運賃情報を取得
        Dim s_count = dictStationCorp.Count
        Dim list_station_range As New List(Of Integer)
        Dim fare As Integer
        Dim surcharge As Integer

        Dim fso = New Scripting.FileSystemObject

        Dim ts As Scripting.TextStream


        For k = 0 To s_count -1'進捗表示
                
            ' カーソル位置を初期化
            'Console.SetCursorPosition(0, Console.CursorTop)

            Dim station_code = dictStationCorp.Keys(k)
            Dim station =  dictStationCorp.Values(k)

            
            If corp_id = 1
                iRange.EnabledStationType = 1
                ' JRの場合，新幹線駅と周辺100kmだけに限定
                Dim range_station = iRange.Search(station.LongName,2,100)
                list_station_range = New List(Of Integer)
                For s_index = 1 To range_station.Count
                    Dim range_s = range_station.GetRangeStation(s_index).GetStation()
                    If range_s.Type = 1
                        Dim s_code = data2.GetStationCode(range_s.LongName)
                        If dictStationCorp.ContainsKey(s_code) 
                            list_station_range.Add(s_code)
                        End If
                    End if
                Next
            
            End If

            Dim latlon = convert_latlon(station)
            Dim fareCount As Integer
            Dim fare_surcharge As List(Of Integer)


            For l = k+1 To s_count -1
                If line_type = 1 Then
                    Console.Write("進捗状況 運賃データ取得中:" & Str(l) &" "& Str(k)  & " / " & Str(s_count)&"  " &  corp.Name &"  " &  station.LongName &"                   ")
                    Console.SetCursorPosition(0, Console.CursorTop)
                End If

                Dim f_to_station_code = dictStationCorp.Keys(l)
                f_to_station = dictStationCorp.Values(l)

                If corp_id = 1
                    If Not list_station_main.Contains(f_to_station_code) And Not list_station_range.Contains(f_to_station_code)
                        Continue For
                    End If
                ElseIf line_type = 32 And s_count > 100 then
                    ' 駅が多い路線バス会社の駅は，周辺30kmに設定 
                    Dim latlon_to = convert_latlon(f_to_station)
                    Dim dist = distance_latlon(latlon, latlon_to)
                    If dist > 30 Then
                        Continue For

                    End If


                End If

                f_to_station =  dictStationCorp.Values(l)

                Dim from_to As Tuple(Of Integer, Integer) = Tuple.Create(station_code, f_to_station_code)
                

                If Not dict_added_corp_local_bus.ContainsKey(from_to) Then
                    iNavi.RemoveAllKey()
                    iNavi.AddKey(station.LongName)
                    iNavi.AddKey(f_to_station.LongName)
                    Dim errcd = iSearch.CheckNavi6(iNavi)   ' ●探索条件チェック
                    Call iSearch.SearchCourse10(iNavi)
                    Try
                        Dim courseCount As Integer = iSearch.CourseCount
                    
                        If (courseCount > 0) Then
                        
                            fareCount = iSearch.FareSectionCount
                            ' ●探索結果数＞０の場合→探索結果が１件以上ある場合

                            If (fareCount > 1) Then
                                If line_type=1 or line_type=2 then
                                    ' 鉄道は他の経路も探索
                                    For c = 2 To courseCount
                                        iSearch.CurrentCourse = c
                                        fareCount = iSearch.FareSectionCount
                                        If fareCount = 1
                                            Exit For
                                        End if

                                    Next
                                Else
                                    ' 他の交通手段は探索せずスキップ
                                    Continue For

                                End If

                            End If

                        End If

                        fare = iSearch.TotalFareS
                        surcharge = iSearch.TotalSurchargeWf
                    Catch ex As Exception
                        'Console.WriteLine(ex.Message)
                        Continue For
                    End Try

                    If line_type = 32 Then
                        fare_surcharge = New List(Of Integer)
                        fare_surcharge.add(fare)
                        fare_surcharge.add(surcharge)

                        dict_added_corp_local_bus.Add(from_to, fare_surcharge)
                    End if

                    If (fareCount = 1) Then
                        sb_fare.Append(Str(corp_id))
                        sb_fare.Append(",")
                        sb_fare.Append(corp.Name)
                        sb_fare.Append(",")
                        sb_fare.Append(Str(line_type))
                        sb_fare.Append(",")
                        sb_fare.Append(Str(line_id))
                        sb_fare.Append(",")
                        sb_fare.Append(line_name)
                        sb_fare.Append(",")
                        sb_fare.Append(Str(station_code))
                        sb_fare.Append(",")
                        sb_fare.Append(Str(f_to_station_code))
                        sb_fare.Append(",")
                        sb_fare.Append(Str(fare))
                        sb_fare.Append(",")
                        sb_fare.Append(Str(surcharge))

                        sb_fare.Append(vbCrLf)

                    Else
                        Continue For

                    End If

                Else 
                    fareCount=1
                    'fare_surcharge = dict_added_corp_local_bus(from_to)
                    
                    'fare = fare_surcharge(0)
                    'surcharge = fare_surcharge(1)


                End If
                
            Next
            
            If sb_fare.Length > byte_limit Then
                csv_fare = sb_fare.ToString()
                csv_fare = Replace(csv_fare, " ", "")
                export_csv(csv_fare, base_dir & "\fare", output_count_fare, byte_limit, ts, fso)
                sb_fare.Clear()
            End If

        Next
        Return sb_fare
    End Function

    Sub create_database(base_dir As String, dataObj As EXPDENGNLib.ExpDataExtraction2, data2 As EXPDENGNLib.ExpData2, iSearch As EXPDENGNLib.ExpDiaSearch10, add_bus As Boolean,database_version As Long)
        Dim iNavi As EXPDENGNLib.ExpDiaNavi6
        iNavi = iSearch.CreateNavi6()
        
        ' 検索の日付　ダイヤ更新日から半年程度以内にする（航空便のダイヤが入っていないため）
        Dim search_date As Long = database_version + 1
        ' 範囲検索用
        Dim iRange = New EXPDENGNLib.ExpRange

        ' ●ExpDiaNavi6オブジェクトに探索条件を設定
        iNavi.AnswerCount = 20         ' ●探索経路要求回答数
        iNavi.LocalBus = True          ' ●路線バスを利用
        iNavi.Date = search_date

        ' 会社一覧を取得
        Dim corp_list As EXPDENGNLib.ExpDiaCorpSet
        Dim corp As EXPDENGNLib.ExpDiaCorp
        Dim csv_line As String  ' CSV に書き込む全データ
        Dim csv_line_train As String  ' CSV に書き込む全データ
        Dim csv_line_local_bus As String  ' CSV に書き込む全データ
        Dim row_line As String ' 1 行分のデータ

        Dim csv_station_all As String
        Dim csv_station As String  ' CSV に書き込む全データ
        Dim row_station As String ' 1 行分のデータ

        Dim from_station As Station
        Dim to_station As Station

        Dim csv_nearest_station As String  ' CSV に書き込む全データ
        Dim row_nearest_station As String ' 1 行分のデータ

        Dim csv_edge As String  ' CSV に書き込む全データ
        Dim row_edge As String ' 1 行分のデータ

        
        Dim csv_fare As String  ' CSV に書き込む全データ

        Dim csv_transfer_line As String  ' CSV に書き込む全データ
        Dim row_transfer_line As String ' 1 行分のデータ

        Dim transfer_station_code_list As New System.Collections.ArrayList()

        Dim fso As Scripting.FileSystemObject
        fso = New Scripting.FileSystemObject

        Dim ts As Scripting.TextStream


        ' 路線一覧, 駅一覧を取得
        Dim line As EXPDENGNLib.ExpDiaLine
        Dim line_list As EXPDENGNLib.ExpDiaLineSet
        Dim line_id As Integer
        Dim line_count As Integer
        Dim line_name_before As String
        line_name_before = ""

        Dim station As EXPDENGNLib.ExpDiaStation
        Dim next_station As EXPDENGNLib.ExpDiaStation
        Dim f_to_station As EXPDENGNLib.ExpDiaStation
        Dim station_list As EXPDENGNLib.ExpDiaStationSet
        Dim station_list_in_corp As EXPDENGNLib.ExpDiaStationSet
        Dim fare_list As EXPDENGNLib.ExpDiaFareSection3
        Dim nearest_station As EXPDENGNLib.ExpDiaStation
        Dim nearest_station_list As EXPDENGNLib.ExpDiaStationSet

        Dim pass_line_list As EXPDENGNLib.ExpDiaLineSet
        Dim pass_line_bus_train_list As EXPDENGNLib.ExpDiaLineSet
        Dim pass_line_stations As New System.Collections.ArrayList()

        Dim transfer_station As EXPDENGNLib.ExpDiaStationSet

        Dim transfer_line As EXPDENGNLib.ExpDiaLine

        Dim station_id As Integer
        Dim line_type As Integer
        Dim line_type_before = -1
        Dim station_code As Long
        Dim next_station_code As Long
        Dim f_to_station_code As Long
        Dim nearest_station_code As Long
        Dim added_station_code As List(Of Long) = New List(Of Long)
        Dim lat As Double
        Dim lon As Double

        Dim aa As Boolean
        Dim errcd As Integer

        Dim board_time As Integer
        Dim walk_time As Integer
        Dim other_time As Integer
        Dim total_time As Integer
        Dim distance As Double

        Dim fare As Integer
        Dim surcharge As Integer

        line_id = 1
        station_id = 1

        csv_line = "line_id,line_name,line_name_long,line_name_short,line_type,station_count,corp_id,corp_name,line_type" & vbCrLf
        ' 駅,edgeのデータは重いので、StringBuilder使う形に変更
        Dim sb_station As New System.Text.StringBuilder(byte_limit)
        Dim sb_station_all As New System.Text.StringBuilder(byte_limit)
        Dim sb_edge As New System.Text.StringBuilder(byte_limit)
        Dim sb_fare As New System.Text.StringBuilder(byte_limit)
        
        Dim sb_transfer_in_line As New System.Text.StringBuilder(byte_limit)

        Dim sb_station_transfer As New System.Text.StringBuilder(byte_limit)

        sb_station.Append("station_code,station_name,station_name_long,station_type,pass_line_count,lat,lon")
        sb_station.Append(vbCrLf)

        sb_station_all.Append("station_code,station_name,station_name_long,station_type,pass_line_count,corp_id,line_id,line_name")
        sb_station_all.Append(vbCrLf)

        sb_edge.Append("from_station_code,from_station_name,to_station_code,to_station_name,line_id,line_name,line_type,board_time,other_time,total_time,distance")
        sb_edge.Append(vbCrLf)

        sb_fare.Append("corp_id,corp_name,line_type,line_id,line_name,from_station_code,to_station_code,fare,surcharge")
        sb_fare.Append(vbCrLf)

        sb_station_transfer.Append("from_station_code,from_station_name,to_station_code,to_station_name")
        sb_station_transfer.Append(vbCrLf)

        sb_transfer_in_line.Append("station_code,station_name_long,line_name_long,line_type")
        sb_transfer_in_line.Append(vbCrLf)

        ' 進捗を表示
        Console.CursorVisible = False
        line_list = dataObj.SearchLine(0, add_bus)
        line_count = line_list.Count

        corp_list = dataObj.SearchCorp2(0, add_bus)

        Dim exit_ = False

        Dim added_line_id=New List(Of Integer)

        For corp_id = 1 To corp_list.Count
            Dim dictStationCorp As New Dictionary(Of Integer, EXPDENGNLib.ExpDiaStation)
            Dim list_station_main As New List(Of Integer)

            corp = corp_list.GetCorp(corp_id)
            line_list = dataObj.SearchLineByCorp(0, corp.Name)

            ' debug
            'If corp.Name <> "東武鉄道" Then
            '    Continue For
'
'            End if

            If exit_ Then
                Exit For
            End If
            ' 路線バスは登録済の駅間は登録しない
            dict_added_corp_local_bus = New Dictionary(of Tuple(of Integer, Integer), List(Of Integer))


            For j = 1 To line_list.Count ' 路線のループ
                line = line_list.GetLine(j)
                If  added_line_id.Contains(line_id) then
                    Continue For
                Else
                    added_line_id.Add(line_id)

                End If


                '進捗表示
                Console.Write("進捗状況 路線データ取得中:" & Str(line_id) & " / " & Str(line_count) & "  " & line.Name & "                   ")
                ' カーソル位置を初期化
                Console.SetCursorPosition(0, Console.CursorTop)


                line_type = corp.GetServiceLineType()
                If line_type_before = -1 Then
                    line_type_before = line_type
                End if
                
                'TODO debug
                'If line_type<>32 Then
                '    Exit For
                'End If


                row_line = Str(line_id) & "," & line.Name & "," & line.LongName & "," & line.ShortName & "," & Str(line.Type) & "," & Str(line.StopStationCount) & "," & Str(corp_id) & "," & corp.Name & "," & Str(corp.GetServiceLineType())
                row_line = Replace(row_line, " ", "")

                ' 行を結合
                csv_line = csv_line & row_line & vbCrLf
                ' 路線バスだけ、それ以外の路線情報ファイルを
                If line_type = 32 Then
                    csv_line_local_bus = csv_line_local_bus & row_line & vbCrLf
                Else
                    csv_line_train = csv_line_train & row_line & vbCrLf

                End If


                station_list = line.SearchStopStation()

                from_station = Nothing
                to_station = Nothing

                iNavi.LocalBusOnly = False     ' ●路線バスのみでの検索 鉄道駅で検索するときは基本False　バス停同士での検索時のみTrueにする
                If line_type = 32 Then
                    iNavi.LocalBusOnly = True
                End If

                                   
                If line_type <> 1 Then
                 ' 鉄道は会社ごと，その他は路線ごとに運賃計算用駅リストを作成する
                    dictStationCorp = New Dictionary(Of Integer, EXPDENGNLib.ExpDiaStation)       
                    
                End If

                For k = 1 To station_list.Count
                    station = station_list.GetStation(k)
                    station_code = data2.GetStationCode(station.LongName)
                    
                    If Not dictStationCorp.ContainsKey(station_code) Then
                        dictStationCorp.Add(station_code, station)
                        If corp_id = 1 And line.Name.Contains("新幹線")
                            ' 新幹線路線,北海道の主要駅はメイン駅として登録
                            list_station_main.Add(station_code)
                        ElseIf corp_id=1 And (station.Name="札幌"or station.Name="旭川"or station.Name="帯広"or station.Name="釧路"or station.Name="北見")
                            ' 新幹線路線,北海道の主要駅はメイン駅として登録
                            list_station_main.Add(station_code)
                        End If
                    End If

                    sb_station_all.Append(Str(station_code))
                    sb_station_all.Append(",")
                    sb_station_all.Append(station.Name)
                    sb_station_all.Append(",")
                    sb_station_all.Append(station.LongName)
                    sb_station_all.Append(",")
                    sb_station_all.Append(Str(station.Type))
                    sb_station_all.Append(",")
                    sb_station_all.Append(Str(station.PassLineCount))
                    sb_station_all.Append(",")
                    sb_station_all.Append(Str(corp_id))
                    sb_station_all.Append(",")
                    sb_station_all.Append(Str(line_id))
                    sb_station_all.Append(",")
                    sb_station_all.Append(line.LongName)
                    sb_station_all.Append(",")
                    sb_station_all.Append(Str(k))

                    sb_station_all.Append(vbCrLf)


                    If Not added_station_code.Contains(station_code) Then
                        added_station_code.Add(station_code)

                        ' 緯度経度が度、分、秒、1/100秒の配列でくるので小数点に変更 ちなみに日本測地系
                        Dim latlon = convert_latlon(station)
                        lat = latlon(0)
                        lon = latlon(1)

                        ' 駅のデータは重いので、StringBuilder使う形に変更
                        'sb_station.Append(Str(station_id))
                        'sb_station.Append(",")
                        sb_station.Append(Str(station_code))
                        sb_station.Append(",")
                        sb_station.Append(station.Name)
                        sb_station.Append(",")
                        sb_station.Append(station.LongName)
                        sb_station.Append(",")
                        sb_station.Append(Str(station.Type))
                        sb_station.Append(",")
                        sb_station.Append(Str(station.PassLineCount))
                        sb_station.Append(",")
                        sb_station.Append(Str(lat))
                        sb_station.Append(",")
                        sb_station.Append(Str(lon))

                        sb_station.Append(vbCrLf)

                    End If

                    station_id += 1
                    line_name_before = line.ShortName

                    ' edgeデータの作成
                    If k < station_list.Count Then
                        next_station = station_list.GetStation(k + 1)
                        next_station_code = data2.GetStationCode(next_station.LongName)

                        iNavi.RemoveAllKey()

                        iNavi.AddKey(station.LongName, line.LongName)
                        iNavi.AddKey(next_station.LongName, line.LongName)

                        errcd = iSearch.CheckNavi6(iNavi)   ' ●探索条件チェック
                        Call iSearch.SearchCourse10(iNavi)

                        If (iSearch.CourseCount > 0) Then
                            ' ●探索結果数＞０の場合→探索結果が１件以上ある場合
                            iSearch.SortType = 1 '探索順
                            board_time = iSearch.TotalBoardTime
                            other_time = iSearch.TotalOtherTime
                            total_time = iSearch.TotalTime
                            distance = iSearch.TotalDistance

                            sb_edge.Append(Str(station_code))
                            sb_edge.Append(",")
                            sb_edge.Append(station.Name)
                            sb_edge.Append(",")
                            sb_edge.Append(Str(next_station_code))
                            sb_edge.Append(",")
                            sb_edge.Append(next_station.Name)
                            sb_edge.Append(",")
                            sb_edge.Append(Str(line_id))
                            sb_edge.Append(",")
                            sb_edge.Append(line.Name)
                            sb_edge.Append(",")
                            sb_edge.Append(Str(line.Type))
                            sb_edge.Append(",")
                            sb_edge.Append(Str(board_time))
                            sb_edge.Append(",")
                            sb_edge.Append(Str(other_time))
                            sb_edge.Append(",")
                            sb_edge.Append(Str(total_time))
                            sb_edge.Append(",")
                            sb_edge.Append(Str(distance / 10))

                            sb_edge.Append(vbCrLf)
                        End If

                    End If

                    '乗り換え情報作成
                    If Not transfer_station_code_list.Contains(station_code) Then
                        pass_line_list = station.SearchPassLine(True)
                        pass_line_bus_train_list = station.SearchPassLine(False)
                        transfer_station = station.SearchNearestStation()
                        transfer_station_code_list.Add(station_code)
                        If pass_line_list.Count > 0 Then
                            For l = 1 To pass_line_list.Count
                                transfer_line = pass_line_list.GetLine(l)
                                sb_transfer_in_line.Append(Str(station_code))
                                sb_transfer_in_line.Append(",")
                                sb_transfer_in_line.Append(station.LongName)
                                sb_transfer_in_line.Append(",")
                                sb_transfer_in_line.Append(transfer_line.LongName)
                                sb_transfer_in_line.Append(",")
                                sb_transfer_in_line.Append(Str(transfer_line.Type))
                                sb_transfer_in_line.Append(vbCrLf)

                            Next
                        End If

                        If transfer_station.Count > 0 Then
                            For t = 1 To transfer_station.Count
                                nearest_station = transfer_station.GetStation(t)
                                nearest_station_code = data2.GetStationCode(nearest_station.LongName)
                                'row_nearest_station = Str(station_code) & "," & station.Name & "," & Str(nearest_station_code) & "," & nearest_station.Name
                                'row_nearest_station = Replace(row_nearest_station, " ", "")
                                'csv_nearest_station = csv_nearest_station & row_nearest_station & vbCrLf

                                sb_station_transfer.Append(Str(station_code))
                                sb_station_transfer.Append(",")
                                sb_station_transfer.Append(station.Name)
                                sb_station_transfer.Append(",")
                                sb_station_transfer.Append(Str(nearest_station_code))
                                sb_station_transfer.Append(",")
                                sb_station_transfer.Append(nearest_station.Name)
                                sb_station_transfer.Append(vbCrLf)

                            Next
                        End If

                    End If

                Next
                
                '運賃情報を取得 （処理に時間がかかるのでスキップ）
                If line_type <> 1 Then
                    ' 鉄道以外は路線ID，路線名ごとに運賃を計算するので，line_idとline_nameつける
                    If line_type = 32 Then
                        ' 路線バスの運賃取得は件数が多すぎるのでTODO
                       'sb_fare = create_fare_database(corp, corp_id,line_type,-1 , "", dictStationCorp,iSearch,iNavi, iRange, data2, list_station_main,f_to_station, base_dir, csv_fare, sb_fare)

                    Else
                        'sb_fare = create_fare_database(corp, corp_id,line_type,line_id , line.LongName, dictStationCorp,iSearch,iNavi, iRange, data2, list_station_main,f_to_station, base_dir, csv_fare, sb_fare)

                    End If


                    
                End If

                line_id += 1

                'TODO　出力データが一定サイズを超えたらいったん出力
                If sb_station_all.Length > byte_limit Or line_type_before <> line_type Then
                    csv_station_all = sb_station_all.ToString()
                    csv_station_all = Replace(csv_station_all, " ", "")
                    export_csv(csv_station_all, base_dir & "\all_station", output_count_station_all, byte_limit, ts, fso)
                    sb_station_all.Clear()
                End If

                If sb_station.Length > byte_limit Or line_type_before <> line_type Then
                    csv_station = sb_station.ToString()
                    csv_station = Replace(csv_station, " ", "")
                    export_csv(csv_station, base_dir & "\station", output_count_station, byte_limit, ts, fso)
                    sb_station.Clear()
                End If

                If sb_edge.Length > byte_limit Or line_type_before <> line_type Then
                    csv_edge = sb_edge.ToString()
                    csv_edge = Replace(csv_edge, " ", "")
                    export_csv(csv_edge, base_dir & "\edge", output_count_edge, byte_limit, ts, fso)
                    sb_edge.Clear()
                End If

                If sb_station_transfer.Length > byte_limit Or line_type_before <> line_type Then
                    csv_nearest_station = sb_station_transfer.ToString()
                    csv_nearest_station = Replace(csv_nearest_station, " ", "")
                    export_csv(csv_nearest_station, base_dir & "\sta_transfer", output_count_nearest_station, byte_limit, ts, fso)
                    sb_station_transfer.Clear()
                End If

                If sb_transfer_in_line.Length > byte_limit Or line_type_before <> line_type Then
                    csv_transfer_line = sb_transfer_in_line.ToString()
                    csv_transfer_line = Replace(csv_transfer_line, " ", "")
                    export_csv(csv_transfer_line, base_dir & "\sta_in_line", output_count_station_in_line, byte_limit, ts, fso)
                    sb_transfer_in_line.Clear()
                End If
                
                If sb_fare.Length > byte_limit Or line_type_before <> line_type Then
                    csv_fare = sb_fare.ToString()
                    csv_fare = Replace(csv_fare, " ", "")
                    export_csv(csv_fare, base_dir & "\fare", output_count_fare, byte_limit, ts, fso)
                    sb_fare.Clear()
                End If

                line_type_before = line_type 

            Next

            ' 運賃情報を取得(処理に時間がかかるのでスキップ)
            If line_type=1Then
                    
                ' 鉄道は会社ごとに運賃を計算するので，line_id=-1とline_name=""とする

                'sb_fare = create_fare_database(corp, corp_id,line_type, -1, "", dictStationCorp,iSearch,iNavi, iRange, data2, list_station_main,f_to_station, base_dir, csv_fare, sb_fare)

            End if

        Next

        '進捗表示
        Console.Write("進捗状況 データ出力中")
        ' カーソル位置を初期化
        Console.SetCursorPosition(0, Console.CursorTop)

        ' 路線一覧出力
        ts = fso.OpenTextFile(base_dir & "\line.csv", Scripting.IOMode.ForWriting, True)
        list_file_type.Add("line")

        ts.Write(csv_line)
        ts.Close()

        'ts = fso.OpenTextFile(base_dir & "\line_train.csv", Scripting.IOMode.ForWriting, True)
        'ts.Write(csv_line_train)
        'ts.Close()

        'ts = fso.OpenTextFile(base_dir & "\line_local_bus.csv", Scripting.IOMode.ForWriting, True)
        'ts.Write(csv_line_local_bus)
        'ts.Close()


        ' 駅一覧出力
        csv_station = sb_station.ToString()
        csv_station = Replace(csv_station, " ", "")
        export_csv(csv_station, base_dir & "\station", output_count_station, byte_limit, ts, fso)

        csv_station_all = sb_station_all.ToString()
        csv_station_all = Replace(csv_station_all, " ", "")
        export_csv(csv_station_all, base_dir & "\all_station", output_count_station_all, byte_limit, ts, fso)

        csv_edge = sb_edge.ToString()
        csv_edge = Replace(csv_edge, " ", "")
        export_csv(csv_edge, base_dir & "\edge", output_count_edge, byte_limit, ts, fso)

        ' 駅-バス停乗り換え情報出力
        'ts = fso.OpenTextFile(base_dir & "\station_transfer" + Str(output_count_nearest_station) + ".csv", Scripting.IOMode.ForWriting, True)
        'list_file_type.Add("station_transfer")
        'ts.Write(csv_nearest_station)
        'ts.Close()
        csv_nearest_station = sb_station_transfer.ToString()
        csv_nearest_station = Replace(csv_nearest_station, " ", "")
        export_csv(csv_nearest_station, base_dir & "\sta_transfer", output_count_nearest_station, byte_limit, ts, fso)

        '駅乗り入れ路線情報情報出力

        'ts = fso.OpenTextFile(base_dir & "\station_in_line" + Str(output_count_station_in_line) + ".csv", Scripting.IOMode.ForWriting, True)
        'list_file_type.Add("station_in_line")
        'ts.Write(csv_transfer_line)
        'ts.Close()
        csv_transfer_line = sb_transfer_in_line.ToString()
        csv_transfer_line = Replace(csv_transfer_line, " ", "")
        export_csv(csv_transfer_line, base_dir & "\sta_in_line", output_count_station_in_line, byte_limit, ts, fso)

        csv_fare = sb_fare.ToString()
        csv_fare = Replace(csv_fare, " ", "")
        export_csv(csv_fare, base_dir & "\fare", output_count_fare, byte_limit, ts, fso)

        Console.CursorVisible = True
    End Sub


    Sub check_update_corp_line_station(data2 As EXPDENGNLib.ExpData2,iConvert As EXPDENGNLib.ExpConvert, line_old As String, station_old As String, base_dir_old As String)
        ' アップデート確認
        Dim fso As Scripting.FileSystemObject
        fso = New Scripting.FileSystemObject
        Dim ts As Scripting.TextStream

        Dim tr As Scripting.TextStream
        Dim byte_limit = 1000000
        Dim row_str As String
        Dim row As String()
        Dim line_row As Line_s
        Dim list_line = New List(Of Line_s)()
        Dim list_line_str = New List(Of String)
        Dim station_row As Station_s
        Dim list_station = New List(Of Station_s)()
        Dim list_station_str = New List(Of String)

        Dim station_code As Long
        Dim list_station_code = New List(Of Long)()

        Dim export_edge_diagram_list As List(Of String) = New List(Of String)
        Dim update_line_list = New List(Of String)()


        ' lineデータ読み込み
        tr = fso.OpenTextFile(line_old, 1)
        Dim init As Boolean = True
        Do While tr.AtEndOfStream = False
            row_str = tr.ReadLine
            If init = False Then
                Try
                    row = Split(row_str, ",")
                    'line_id,line_name,line_name_long,line_name_short,line_type,station_count,corp_id,corp_name
                    If row.Length > 1 Then
                        line_row = New Line_s(Long.Parse(row(0)), row(1), row(2),row(3), Integer.Parse(row(4)), Integer.Parse(row(5)), Integer.Parse(row(6)), row(7))
                        list_line.Add(line_row)
                        list_line_str.Add(row_str)
                    End If

                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                End Try

            Else
                init = False '一行目のheader飛ばす
            End If
        Loop

        tr.Close()

        ' station読み込み
        tr = fso.OpenTextFile(station_old, 1)
        init = True
        Do While tr.AtEndOfStream = False
            row_str = tr.ReadLine
            If init = False Then
                Try
                    row = Split(row_str, ",")
                    'station_code,station_name,station_name_long,station_type,pass_line_count,lat,lon
                    If row.Length > 1 Then
                        station_row = New Station_s(Long.Parse(row(0)), row(1), row(2), Integer.Parse(row(3)), Integer.Parse(row(4)), Double.Parse(row(5)), Double.Parse(row(6)))
                        list_station.Add(station_row)
                        list_station_str.Add(row_str)
                    End If

                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                End Try

            Else
                init = False '一行目のheader飛ばす
            End If
        Loop

        Dim corp_name
        Dim update_flg_corp
        Dim update_flg_line

        Dim corp_name_up
        Dim line_name
        Dim line_name_mid
        Dim line_name_short
        Dim line_name_up = ""
        Dim csv_line_diff
        Dim sb_out As New System.Text.StringBuilder(byte_limit)
        
        Dim out_line As String
        sb_out.append("line_id,line_name,line_name_long,line_name_short,line_type,station_count,corp_id,corp_name,line_type,update_flg_corp,corp_name_up,update_flg_line,line_name_up")
        sb_out.Append(vbCrLf)

        For i = 0 To list_line.Count - 1
            Console.Write("更新確認状況・路線:" & Str(i) & " / " & Str(list_line.Count))
            Console.SetCursorPosition(0, Console.CursorTop)

            line_row = list_line(i)
            row_str=list_line_str(i)
            corp_name = line_row.corp_name
            update_flg_corp= data2.CheckRenewalCorp(corp_name)
            corp_name_up=""
            If update_flg_corp = 2 Then
                corp_name_up = data2.GetRenewalCorpName(corp_name)

            End If
            line_name_up=""
            line_name = line_row.line_name_long
            line_name_mid = line_row.line_name
            line_name_short = line_row.line_name_short
            update_flg_line= data2.CheckRenewalLine(line_name)
            If update_flg_line = -1 Then
                update_flg_line= data2.CheckRenewalLine(line_name_mid)
                If update_flg_line = -1 Then
                    update_flg_line= data2.CheckRenewalLine(line_name_short)

                End If

            End If
            'それでも-1になる場合は変換を試す
            If update_flg_line = -1 Then
                line_name_mid = iConvert.ConvertLineNameIntraToSDK(line_name_mid)
                update_flg_line= data2.CheckRenewalLine(line_name_mid)
            End If

            If update_flg_line = 2 Then
                line_name_up = data2.GetRenewalLineName(line_name)

            End If

            out_line = row_str
            out_line = $"{out_line},{update_flg_corp},{corp_name_up},{update_flg_line},{line_name_up}"
            sb_out.append(out_line)
            sb_out.Append(vbCrLf)

        Next

        csv_line_diff = sb_out.ToString()
        csv_line_diff = Replace(csv_line_diff, " ", "")
        ts = fso.OpenTextFile(base_dir_old & "\line_corp_diff.csv", Scripting.IOMode.ForWriting, True)
        ts.Write(csv_line_diff)
        ts.Close()

        Dim update_flg_station
        Dim station_name
        Dim station_name_short

        Dim station_name_up

        'station_code,station_name,station_name_long,station_type,pass_line_count,lat,lon
        Dim out_station As String

        sb_out.clear()
        sb_out.append("station_code,station_name,station_name_long,station_type,pass_line_count,lat,lon,update_flg_station,station_name_up")
        sb_out.Append(vbCrLf)
        Dim station As EXPDENGNLib.ExpDiaStation

        For i = 0 To list_station.Count - 1
            Console.Write("更新確認状況・駅:" & Str(i) & " / " & Str(list_station.Count))
            Console.SetCursorPosition(0, Console.CursorTop)

            station_row = list_station(i)
            row_str = list_station_str(i)

            station_name = station_row.station_name_long
            station_name_short = station_row.station_name
            station_code = station_row.station_code
            update_flg_station= data2.CheckRenewalStation(station_name)
            station_name_up=""
            If update_flg_station = -1 Then
                update_flg_station= data2.CheckRenewalStation(station_name_short)
                
            End if
            
            station_name_up = data2.GetRenewalStationName(station_name)

            If update_flg_station = -1 Then
                Try
                    station = data2.GetStationByCode(station_code)
                    station_name_up = station.LongName
                    update_flg_station = -2
                        
                Catch 
                      

                End Try

            End if

            out_station = row_str
            out_station = $"{out_station},{update_flg_station},{station_name_up}"
            sb_out.append(out_station)
            sb_out.Append(vbCrLf)

        Next

        Dim csv_station_diff
        csv_station_diff = sb_out.ToString()
        csv_station_diff = Replace(csv_station_diff, " ", "")
        ts = fso.OpenTextFile(base_dir_old & "\station_diff.csv", Scripting.IOMode.ForWriting, True)
        ts.Write(csv_station_diff)
        ts.Close()


    End Sub


    Sub Main()
        ' バス会社探索するか
        Dim add_bus As Boolean = True

        ' 検索オブジェクト
        dataObj = New EXPDENGNLib.ExpDataExtraction2
        data2 = New EXPDENGNLib.ExpData2
        data_diaDB = New EXPDENGNLib.ExpDiaDB10
        iSearch = New EXPDENGNLib.ExpDiaSearch10
        iConvert = New EXPDENGNLib.ExpConvert

        ' バージョン取得
        Dim database_version = data_diaDB.KNBVersion

        Dim base_dir As String = $"D:\exp\database\{database_version}"
        Dim temp_dir As String = $"{base_dir}\database_temp"

        ' 会社・路線・駅の更新をチェックするための作成済バージョンのデータ
        'Dim database_version_old = 20200301
        'Console.WriteLine($"update check: {database_version_old} to {database_version}")

        'Dim base_dir_old As String = $"D:\exp\database\{database_version_old}"
        'Dim line_old As String = $"{base_dir_old}\line.csv"
        'Dim station_old As String = $"{base_dir_old}\station.csv"

        ' 差分をチェック（駅すぱあとの機能を利用して）
        'check_update_corp_line_station(data2,iConvert, line_old, station_old,base_dir_old)


        Console.WriteLine("output_dir:" & base_dir)

        If Dir(base_dir, vbDirectory) = "" Then
            MkDir(base_dir)
        End If

        If Dir(temp_dir, vbDirectory) = "" Then
            MkDir(temp_dir)
        End If


        '路線一覧、駅一覧、駅-バス乗り換え情報、駅乗り入れ路線一覧を取得
        create_database(temp_dir, dataObj, data2, iSearch, add_bus, database_version)

        '分割したファイルをマージ

        list_file_type.Add("line")
        list_file_type.Add("station")
        list_file_type.Add("all_station")
        list_file_type.Add("edge")
        list_file_type.Add("sta_transfer")
        list_file_type.Add("sta_in_line")


        'Processオブジェクトを作成
        Dim p As New System.Diagnostics.Process()
        Dim command As String

        'ComSpec(cmd.exe)のパスを取得して、FileNameプロパティに指定
        p.StartInfo.FileName = System.Environment.GetEnvironmentVariable("ComSpec")
        '出力を読み取れるようにする
        p.StartInfo.UseShellExecute = False
        p.StartInfo.RedirectStandardOutput = True
        p.StartInfo.RedirectStandardInput = False
        'ウィンドウを表示しないようにする
        p.StartInfo.CreateNoWindow = False

        For Each file_type As String In list_file_type

            command = "/c copy　/B " + temp_dir + "\" + file_type + "*.csv " + base_dir + "\" + file_type + ".csv"
            'コマンドラインを指定（"/c"は実行後閉じるために必要）
            p.StartInfo.Arguments = command

            '起動
            p.Start()

            '出力を読み取る
            Dim results As String = p.StandardOutput.ReadToEnd()

            'プロセス終了まで待機する
            'WaitForExitはReadToEndの後である必要がある
            '(親プロセス、子プロセスでブロック防止のため)
            p.WaitForExit()
            p.Close()
            '出力された結果を表示
            Console.WriteLine(results)


        Next

    End Sub

End Module
