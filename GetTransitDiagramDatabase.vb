Module GetTransitDiagramDatabase
    ’ダイヤ情報を作成するためのモジュール 　@author: hasega60
    ’
    ’
    Public iSearch As EXPDENGNLib.ExpDiaDB10
    Public data2 As EXPDENGNLib.ExpData2
    Public dt1 As DateTime

    Public byte_limit As Integer = 1000000

    Private Class Station
        ReadOnly Property station_id As Long
        ReadOnly Property station_code As Long
        ReadOnly Property station_type As Integer
        ReadOnly Property corp_id As Integer

        Sub New(stationId As Long, stationCode As Long, stationType As Integer, corpId As Integer)
            station_id = stationId
            station_code = stationCode
            station_type = stationType
            corp_id = corpId
            
        End Sub


    End Class

    Private Class Edge
        ReadOnly Property from_station_code As Long
        ReadOnly Property to_station_code As Long

        ReadOnly Property line_id As Long
        ReadOnly Property line_name As String
        ReadOnly Property line_type As Integer

        Sub New(from_code As Long, to_code As Long, lineid As Long, linename As String, linetype As Integer)
            from_station_code = from_code
            to_station_code = to_code
            line_id = lineid
            line_name = linename
            line_type = linetype

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

    Private Class Train_stop
        ReadOnly Property line_id As Long
        ReadOnly Property line_name As String
        ReadOnly Property line_type As Integer
        ReadOnly Property train_id As Long
        ReadOnly Property train_code As Integer
        ReadOnly Property train_name As String
        ReadOnly Property airline_code As String
        ReadOnly Property station_code As Long
        ReadOnly Property station_name As String
        ReadOnly Property stop_order As Integer
        ReadOnly Property arrive_time As Integer
        ReadOnly Property departure_time As Integer

        ReadOnly Property direction As Integer
        Sub New(lineId As Long, lineName As String, lineType As Integer, trainId As Integer, trainCode As Integer, trainName As String, airlineCode As String, stationCode As Long, stationName As String, stopOrder As Integer, arriveTime As Integer, departureTime As Integer, d As Integer)
            line_id = lineId
            line_name = lineName
            line_type = lineType
            train_id = trainId
            train_code = trainCode
            train_name = trainName
            airline_code = airlineCode
            station_code = stationCode
            station_name = stationName
            stop_order = stopOrder

            arrive_time = arriveTime
            departure_time = departureTime
            direction = d
        End Sub

    End Class

    Private Function export_csv_part(ByRef csv As String, csv_base_path As String, max_byte As Integer, ts As Scripting.TextStream, fso As Scripting.FileSystemObject, export_force As Boolean, file_ext As String)
        If Len(csv) > max_byte Or export_force = True Then
            If file_ext = "" Then
                dt1 = DateTime.Now
                export_csv(csv, csv_base_path, dt1.ToString("MMddHHmmssfff"), max_byte, ts, fso)
            Else
                export_csv(csv, csv_base_path, file_ext, max_byte, ts, fso)

            End If
        End If

    End Function

    Private Function export_csv(ByRef csv As String, csv_base_path As String, ByRef file_ext As String, max_byte As Integer, ts As Scripting.TextStream, fso As Scripting.FileSystemObject)
        Dim path As String = csv_base_path + file_ext + ".csv"
        path = Replace(path, " ", "")
        ts = fso.OpenTextFile(path, Scripting.IOMode.ForWriting, True)

        ts.Write(csv)
        ts.Close()
        csv = ""

    End Function

    Private Function _search_od(ByRef export_edge_diagram_list, from_station_name, to_station_name, from_station_code, to_station_code, time, line_id, line_name, iNavi, iSearch, search_mode)
        Dim search_results As EXPDENGNLib.ExpDiaCourseSet10
        Dim search_result As EXPDENGNLib.ExpDiaCourse10
        Dim section As EXPDENGNLib.ExpDiaRouteSection3

        Dim board_time As Integer

        Dim errcd As Integer
        Dim aa As Boolean

        ' 検索の日付　ダイヤ更新日から半年程度以内にする（航空便のダイヤが入っていないため）
        Dim search_date As Long = 20190401

        Dim fso As Scripting.FileSystemObject
        fso = New Scripting.FileSystemObject

        Dim section_train As EXPDENGNLib.ExpDiaTrainInfo
        Dim section_fare As EXPDENGNLib.ExpDiaFareSection3

        Dim list_station = New List(Of Station)()

        Dim list_station_from = New List(Of Station)()

        Dim access_set1 As EXPDENGNLib.ExpDiaRouteAccessSet2
        Dim access_set2 As EXPDENGNLib.ExpDiaRouteAccessSet2
        Dim access As EXPDENGNLib.ExpDiaRouteAccess2
        Dim train_stop_station As EXPDENGNLib.ExpDiaStationSet
        Dim terminal_station As EXPDENGNLib.ExpDiaStation
        Dim section_start_time As Integer
        Dim section_end_time As Integer

        Dim name As String
        Dim id As Long
        Dim airline_cd As String
        Dim init As Boolean

        Dim row_edge_diagram = ""
        Dim row_train_info = ""
        Dim edge_bin As Integer
        Dim route_section_count As Long
        Dim access_count As Long

        Dim section_type_str As String

        aa = iNavi.RemoveAllKey
        ' 駅のfrom_toでループまわす
        'aa = iNavi.ReplaceKeyDDEStyle("高円寺,[ＪＲ総武線],新宿")
        aa = iNavi.AddKey(from_station_name)
        aa = iNavi.AddKey(to_station_name)

        errcd = iSearch.CheckNavi6(iNavi)   ' ●探索条件チェック

        ' チェック＝ＯＫでなくても探索できる場合（チェック＝ワーニング）もあるのでとにかく探索する

        'Call iSearch.SearchCourse10(iNavi)

        search_results = iSearch.SearchCourse10(iNavi)
        init = True
        If (search_results.CourseCount > 0) Then
            ' ●探索結果数＞０の場合→探索結果が１件以上ある場合
            search_result = search_results.GetCourse10(search_mode, 1)

            edge_bin = 1
            route_section_count = search_result.RouteSectionCount
            For rc = 1 To route_section_count
                Try
                    If route_section_count < rc Then
                        ' 再検索でroute_section_countが更新されることがあり、エラーになるケースがあるので対処
                        Continue For
                    End If
                    access_set1 = search_result.GetTrainList3(rc, search_date, 1)
                    access_count = access_set1.Count
                    If access_count > 0 Then
                        For item = 1 To access_count
                            access = access_set1.Item(item)

                            name = access.LongName
                            If name.Contains(line_name) Then

                                'train_stop_station = access.SearchStopStation(True)

                                'train_deperture_station = train_stop_station.GetStation(1)
                                'terminal_station = train_stop_station.GetStation(train_stop_station.Count)

                                '列車を指定して再検索
                                search_result.AssignCourseByTrain2(rc, item, 1)

                                route_section_count = search_result.RouteSectionCount

                                section = search_result.GetRouteSection3(1)
                                If section.StopStationCount >= 1 Then
                                    ' エッジ間なので、必ずStopStationCountは0になるはず
                                    Continue For

                                End If
                                section_train = section.TrainInfo
                                board_time = section.BoardTime

                                section_start_time = section_train.TimetableTime(1) '起点出発時間
                                section_end_time = section_train.TimetableTime(2) '

                                id = access.Id
                                airline_cd = access.Symbol

                                ' from_station_code,to_station_code,from_station_name,to_station_name,line_id,line_name,line_bin,edge_bin,airline_cd,section_start_time,section_end_time,board_time
                                row_edge_diagram = Str(from_station_code) & "," & Str(to_station_code) & "," & from_station_name & "," & to_station_name & "," & Str(line_id) & "," & line_name & "," & Str(id) & "," & Str(edge_bin) & "," & airline_cd & "," & Str(section_start_time) & "," & Str(section_end_time)
                                row_edge_diagram = Replace(row_edge_diagram, " ", "")

                                'row_train_info = ""
                                'row_train_info = Replace(row_train_info, " ", "")

                                export_edge_diagram_list.Add(row_edge_diagram)
                                edge_bin = edge_bin + 1

                            End If

                        Next
                    End If

                Catch ex As Exception

                    Console.WriteLine(ex.Message & " :" & name)
                End Try
            Next
        End If
    End Function


    Private Function search_train_on_line(line_list As List(Of Line), station_code_line_dict As  Dictionary(Of Integer, List(of Long)), iNavi As EXPDENGNLib.ExpDiaNavi6, iSearch As EXPDENGNLib.ExpDiaDB10, line_type_dict As Dictionary(Of Integer, String), export_path As String, search_mode As Integer, search_date As Long)
        Dim list_search_od = New List(Of Station())()
        Dim time As Integer = 0
        Dim min As Integer = 0
        Dim hour As Integer = 0
        Dim ts As Scripting.TextStream

        Dim line_id As Long
        Dim line_name_before As String = ""

        Dim byte_limit As Integer = 1000000

        Dim fso As Scripting.FileSystemObject
        fso = New Scripting.FileSystemObject

        Dim csv_out = ""
        Dim direction = 1

        Dim train_id As Long

        Dim line_train_list As List(Of Train_stop) = New List(Of Train_stop)
        Dim station_code_in_line As List(of Long)
        Dim hit As Boolean
        'Dim row_train_list = ""
        Dim row As String
        Dim st As Train_stop
        Dim export_dia_list = New List(Of String)

        Dim export_row_count As Integer = 100000
        Dim sb_out As New System.Text.StringBuilder(byte_limit)

        '追加済列車　
        Dim added_train_dict = New Dictionary(Of String, Long)

        ' 新しい路線リスト
        Dim new_line_dict = New Dictionary(Of String, Long())

        Dim sb As New System.Text.StringBuilder(200 * export_row_count)

        For each line_type As Integer In line_type_dict.Keys()

            Dim export_path_type = $"{export_path}\{line_type}"
            If Dir(export_path_type, vbDirectory) = "" Then
                MkDir(export_path_type)
            End If

            For Each line As Line In line_list
                Dim line_type_int = line.line_type
                If line_type_int <> line_type Then
                        Continue For
                End if
                    
                line_id = line.line_id
                
                ' debug
                'If line_id < 40000 Then
                '    Continue For
                'End If

                If station_code_line_dict.ContainsKey(line_id) = False Then
                    Continue For
                End If


                station_code_in_line = station_code_line_dict(line_id)


                Dim station_count = station_code_in_line.Count
                Dim div_n = 3
                Dim from_station_name
                Dim to_station_name
                Dim i = 0
                direction = 0
                If station_count >= 6 Then
                    '5以上なら div_n+1 間隔で取得
                    Dim ite_count = Math.Floor(station_count/(div_n+1))-1

                    For i = 0 To ite_count
                        from_station_name = data2.GetStationNameByCode(station_code_in_line(i * (div_n)))
                        to_station_name = data2.GetStationNameByCode(station_code_in_line((i+1)* (div_n)))
                        _search_train2(line_train_list, new_line_dict, added_train_dict, train_id, hit, from_station_name, to_station_name, line, iNavi, iSearch, data2, search_mode, direction, search_date)

                    Next

                    from_station_name =to_station_name
                    to_station_name = data2.GetStationNameByCode(station_code_in_line(station_count-1))
                    _search_train2(line_train_list, new_line_dict, added_train_dict, train_id, hit, from_station_name, to_station_name, line, iNavi, iSearch, data2, search_mode, direction, search_date)

                    ' 逆方向
                    direction = 1
                    For i = 0 To ite_count
                        from_station_name = data2.GetStationNameByCode(station_code_in_line((i+1) * (div_n)))
                        to_station_name = data2.GetStationNameByCode(station_code_in_line((i)* (div_n)))
                        _search_train2(line_train_list, new_line_dict, added_train_dict, train_id, hit, from_station_name, to_station_name, line, iNavi, iSearch, data2, search_mode, direction, search_date)

                    Next
                    
                    to_station_name =from_station_name
                    from_station_name = data2.GetStationNameByCode(station_code_in_line(station_count-1))
                    _search_train2(line_train_list, new_line_dict, added_train_dict, train_id, hit, from_station_name, to_station_name, line, iNavi, iSearch, data2, search_mode, direction, search_date)


                ElseIf station_count >= 3 Then
                    '3,4,5なら最初, ２番目，最後の駅
                    from_station_name = data2.GetStationNameByCode(station_code_in_line(0))
                    to_station_name = data2.GetStationNameByCode(station_code_in_line(1))
                    _search_train2(line_train_list, new_line_dict, added_train_dict, train_id, hit, from_station_name, to_station_name, line,
                                   iNavi, iSearch, data2, search_mode, direction, search_date)

                    from_station_name = data2.GetStationNameByCode(station_code_in_line(1))
                    to_station_name = data2.GetStationNameByCode(station_code_in_line(station_count-1))
                    _search_train2(line_train_list, new_line_dict, added_train_dict, train_id, hit, from_station_name, to_station_name, line,
                                   iNavi, iSearch, data2, search_mode, direction, search_date)

                    ' 逆方向
                    direction = 1
                    from_station_name = data2.GetStationNameByCode(station_code_in_line(station_count-1))
                    to_station_name = data2.GetStationNameByCode(station_code_in_line(station_count-2))
                    _search_train2(line_train_list, new_line_dict, added_train_dict, train_id, hit, from_station_name, to_station_name, line,
                                   iNavi, iSearch, data2, search_mode, direction, search_date)

                    from_station_name = data2.GetStationNameByCode(station_code_in_line(station_count-2))
                    to_station_name = data2.GetStationNameByCode(station_code_in_line(0))
                    _search_train2(line_train_list, new_line_dict, added_train_dict, train_id, hit, from_station_name, to_station_name, line,
                                   iNavi, iSearch, data2, search_mode, direction, search_date)

                Else
                    '2なら最初と最後の駅
                    from_station_name = data2.GetStationNameByCode(station_code_in_line(0))
                    to_station_name = data2.GetStationNameByCode(station_code_in_line(station_count-1))
                    _search_train2(line_train_list, new_line_dict, added_train_dict, train_id, hit, from_station_name, to_station_name, line,
                                   iNavi, iSearch, data2, search_mode, direction, search_date)

                     ' 逆方向
                    direction = 1
                    from_station_name =  data2.GetStationNameByCode(station_code_in_line(station_count-1))
                    to_station_name =data2.GetStationNameByCode(station_code_in_line(0))
                    _search_train2(line_train_list, new_line_dict, added_train_dict, train_id, hit, from_station_name, to_station_name, line,
                                   iNavi, iSearch, data2, search_mode, direction, search_date)

                End If
               
                'サイズが大きくなったらいったん出力
                If line_train_list.Count > export_row_count Then
                    csv_out=""
                    For l = 0 To line_train_list.Count - 1
                        Console.Write("一旦出力:" & Str(l) & " / " & Str(line_train_list.Count) & "                                   ")
                        Console.SetCursorPosition(0, Console.CursorTop)

                        Dim train_stop As Train_stop  = line_train_list(l)

                        ' "line_id_org", "line_name", "line_type","direction", "train_id_org", "train_code", "train_name", "airline_code","stop_order", "station_code", "station_name", "arrive_time", "departure_time"
                        'row = Str(train_stop.line_id) & "," & train_stop.line_name & "," & Str(train_stop.line_type) & "," & Str(train_stop.direction) & "," & Str(train_stop.train_id) & "," & train_stop.train_code & "," & train_stop.train_name & "," & train_stop.airline_code & "," & Str(train_stop.stop_order) & "," & Str(train_stop.station_code)  & "," & train_stop.station_name & "," & Str(train_stop.arrive_time)&"," &  Str(train_stop.departure_time)
                        'row = Replace(row, " ", "")
                        'csv_out = csv_out & row & vbCrLf

                        sb_out.Append(Str(train_stop.line_id))
                        sb_out.Append(",")
                        sb_out.Append( train_stop.line_name)
                        sb_out.Append(",")
                        sb_out.Append( Str(train_stop.line_type))
                        sb_out.Append(",")
                        sb_out.Append(Str(train_stop.direction))
                        sb_out.Append(",")
                        sb_out.Append(Str(train_stop.train_id))
                        sb_out.Append(",")
                        sb_out.Append(train_stop.train_code)
                        sb_out.Append(",")
                        sb_out.Append(train_stop.train_name)
                        sb_out.Append(",")
                        sb_out.Append(train_stop.airline_code)
                        sb_out.Append(",")
                        sb_out.Append(Str(train_stop.stop_order))
                        sb_out.Append(",")
                        sb_out.Append(Str(train_stop.station_code))
                        sb_out.Append(",")
                        sb_out.Append(train_stop.station_name)
                        sb_out.Append(",")
                        sb_out.Append(Str(train_stop.arrive_time))
                        sb_out.Append(",")
                        sb_out.Append(Str(train_stop.departure_time))

                        sb_out.Append(vbCrLf)


                        

                    Next
                    csv_out = sb_out.ToString()
                    csv_out = Replace(csv_out, " ", "")

                    export_csv_part(csv_out, $"{export_path_type}\dia" , byte_limit, ts, fso, True, "")

                    '出力リスト初期化
                    sb_out.Clear()
                    line_train_list = New List(Of Train_stop)
                End If

       
            

            Next

            ' 最終出力
            csv_out=""
            Console.Write("最終出力:" & "dia                                   ")
            Console.SetCursorPosition(0, Console.CursorTop)

            For l = 0 To line_train_list.Count - 1
                Dim train_stop As Train_stop  = line_train_list(l)

                ' "line_id_org", "line_name", "line_type","direction", "train_id_org", "train_code", "train_name", "airline_code","stop_order", "station_code", "station_name", "arrive_time", "departure_time"
                'row = Str(train_stop.line_id) & "," & train_stop.line_name & "," & Str(train_stop.line_type) & "," & Str(train_stop.direction) & "," & Str(train_stop.train_id) & "," & train_stop.train_code & "," & train_stop.train_name & "," & train_stop.airline_code & "," & Str(train_stop.stop_order) & "," & Str(train_stop.station_code)  & "," & train_stop.station_name & "," & Str(train_stop.arrive_time)&"," &  Str(train_stop.departure_time)
                'row = Replace(row, " ", "")
                'csv_out = csv_out & row & vbCrLf

                sb_out.Append(Str(train_stop.line_id))
                sb_out.Append(",")
                sb_out.Append( train_stop.line_name)
                sb_out.Append(",")
                sb_out.Append( Str(train_stop.line_type))
                sb_out.Append(",")
                sb_out.Append(Str(train_stop.direction))
                sb_out.Append(",")
                sb_out.Append(Str(train_stop.train_id))
                sb_out.Append(",")
                sb_out.Append(train_stop.train_code)
                sb_out.Append(",")
                sb_out.Append(train_stop.train_name)
                sb_out.Append(",")
                sb_out.Append(train_stop.airline_code)
                sb_out.Append(",")
                sb_out.Append(Str(train_stop.stop_order))
                sb_out.Append(",")
                sb_out.Append(Str(train_stop.station_code))
                sb_out.Append(",")
                sb_out.Append(train_stop.station_name)
                sb_out.Append(",")
                sb_out.Append(Str(train_stop.arrive_time))
                sb_out.Append(",")
                sb_out.Append(Str(train_stop.departure_time))

                sb_out.Append(vbCrLf)                        

            Next
            csv_out = sb_out.ToString()
            csv_out = Replace(csv_out, " ", "")

            export_csv_part(csv_out, $"{export_path_type}\dia" , byte_limit, ts, fso, True, "")

            '出力リスト初期化
            sb_out.Clear()
            line_train_list = New List(Of Train_stop)
            
            Console.Write("最終出力:" & "new_line_list                                   ")
            Console.SetCursorPosition(0, Console.CursorTop)
            Dim writer As New System.IO.StreamWriter($"{export_path_type}\new_line_list{line_type}.csv", False,System.Text.Encoding.GetEncoding("shift_jis"))
            writer.AutoFlush = True

            For each kv In new_line_dict
                sb_out.Clear()
                Dim line_name = kv.Key
                Dim line_id_new = kv.Value(0)
                Dim line_id_org = kv.Value(1)

                sb_out.Append(Str(line_id_new))
                sb_out.Append(",")
                sb_out.Append(line_name)
                sb_out.Append(",")
                sb_out.Append(Str(line_id_org))
                 
               ' sb_out.Append(vbCrLf)   

                csv_out = sb_out.ToString()
                csv_out = Replace(csv_out, " ", "")
                writer.WriteLine(csv_out)

            Next
            writer.close()



            Console.Write("最終出力:" & "train_list                                   ")
            Console.SetCursorPosition(0, Console.CursorTop)
            
            writer = New System.IO.StreamWriter($"{export_path_type}\train_list{line_type}.csv", False,System.Text.Encoding.GetEncoding("shift_jis"))
            writer.AutoFlush = True

            For each kv In added_train_dict
                sb_out.Clear()
                Dim train_name = kv.Key
                train_id = kv.Value

                sb_out.Append(Str(train_id))
                sb_out.Append(",")
                sb_out.Append(train_name)
                 
                csv_out = sb_out.ToString()
                csv_out = Replace(csv_out, " ", "")
                writer.WriteLine(csv_out)

            Next
            
            writer.close()

        Next


        Console.WriteLine("完了するには何かキーを押してください．．．")
        Console.ReadKey()
        Console.CursorVisible = True


    End Function

    Private Function _set_train_list_from_access(ByRef line_train_list, ByRef added_train_dict, ByRef train_id, access, line_name, line_id, direction, line_type)
        Dim train_stop_station As EXPDENGNLib.ExpDiaStationSet
        Dim train_name As String
        Dim train_cd As Integer
        Dim stop_order As Integer
        Dim airline_cd As String
        Dim destination = ""
        Dim t_station_name As String
        Dim t_station_code As Long
        Dim t_arrive_time As Integer
        Dim t_deperture_time As Integer
        Dim t_terminal_time As Integer
        Dim train_stop_info As Train_stop
        Dim row_train_list = ""

        Dim name = access.LongName

        ' のぞみ1号　の号数、航空便コードを取得
        train_cd = access.Id
        ' 行先を取得
        If name.Contains("・") Then
            destination = name.Split("・")(name.Split("・").Length - 1)
        End If
        airline_cd = access.Symbol
        ' 航空便のときにairline_cdが入っていなかったら飛ばす
        'If line_type = 2 And airline_cd = "" Then
        '    Continue For
        'End If

        train_stop_station = access.SearchStopStation(True)
        stop_order = 1
        For s = 1 To train_stop_station.Count
            t_deperture_time = access.GetStationDepartureTime(s)
            t_terminal_time = access.GetStationArrivalTime(train_stop_station.Count)
            If s = 1 Then
                train_name = name + "_" + Str(train_cd) + "_" + Str(t_deperture_time) + "_" + Str(t_terminal_time) + "_" + Str(train_stop_station.Count)
                train_name = Replace(train_name, " ", "")

                If added_train_dict.ContainsKey(train_name) Then
                    Exit For
                End If
                train_id = added_train_dict.Count
                added_train_dict.Add(train_name, added_train_dict.Count)
            End If

            t_arrive_time = access.GetStationArrivalTime(s)
            t_station_name = train_stop_station.GetStationLongName(s)
            t_station_code = data2.GetStationCode(t_station_name)
            train_stop_info = New Train_stop(line_id, line_name, line_type, train_id, train_cd, name, airline_cd, t_station_code, t_station_name, stop_order, t_arrive_time, t_deperture_time, direction)

            line_train_list.Add(train_stop_info)
            stop_order = stop_order + 1
        Next

    End Function

    Private Function _search_train(ByRef line_train_list, ByRef added_train_list, ByRef train_id, ByRef search_hit, from_station_name, to_station_name, line, iNavi, iSearch, data2, search_mode, direction, search_date)

        Dim list_search_od = New List(Of Station())()
        Dim line_name As String
        Dim line_name_long As String
        Dim line_type As Integer

        Dim line_obj As EXPDENGNLib.ExpDiaLine
        Dim station_list As EXPDENGNLib.ExpDiaStationSet

        Dim first_station As EXPDENGNLib.ExpDiaStation
        Dim last_station As EXPDENGNLib.ExpDiaStation

        Dim time As Integer = 0
        Dim min As Integer = 0
        Dim hour As Integer = 0

        Dim line_id As Long
        Dim line_name_before As String = ""
        Dim search_results As EXPDENGNLib.ExpDiaCourseSet10
        Dim search_result As EXPDENGNLib.ExpDiaCourse10
        Dim init As Boolean

        Dim access_set1 As EXPDENGNLib.ExpDiaRouteAccessSet2
        Dim access_set2 As EXPDENGNLib.ExpDiaRouteAccessSet2
        Dim access As EXPDENGNLib.ExpDiaRouteAccess2
        Dim train_stop_station As EXPDENGNLib.ExpDiaStationSet

        Dim errcd As Integer
        Dim aa As Boolean

        Dim byte_limit As Integer = 1000000

        Dim fso As Scripting.FileSystemObject
        fso = New Scripting.FileSystemObject

        Dim row_diagram = ""
        Dim route_section_count As Integer

        Dim csv_out = ""
        Dim t_station_name As String
        Dim t_station_code As Long
        Dim t_arrive_time As Integer
        Dim t_deperture_time As Integer
        Dim train_name As String
        Dim train_cd As Integer
        Dim stop_order As Integer
        Dim airline_cd As String
        Dim train_stop_info As Train_stop
        Dim row_train_list = ""
        Dim numbers_str As Char() = New Char() {"０"c, "１"c, "２"c, "３"c, "４"c, "５"c, "６"c, "７"c, "８"c, "９"c}

        Dim time_start
        Dim name
        Dim destination = ""
        Dim access_line_type
        Dim names
        Dim hit_name

        Dim hit = False

        line_name = line.line_name
        line_name_long = line.line_name_long
        line_id = line.line_id
        line_type = line.line_type

        aa = iNavi.RemoveAllKey

        aa = iNavi.AddKey(from_station_name, line.line_name_long)
        aa = iNavi.AddKey(to_station_name, line.line_name_long)


        errcd = iSearch.CheckNavi6(iNavi)   ' ●探索条件チェック

        'Call iSearch.SearchCourse10(iNavi)

        search_results = iSearch.SearchCourse10(iNavi)
        init = True
        If (search_results.CourseCount > 0) Then
            ' ●探索結果数＞０の場合→探索結果が１件以上ある場合
            For cc = 1 To search_results.CourseCount
                search_result = search_results.GetCourse10(search_mode, cc)
                route_section_count = search_result.RouteSectionCount
                For rc = 1 To route_section_count
                    Try
                        If route_section_count < rc Then
                            ' 再検索でroute_section_countが更新されることがあり、エラーになるケースがあるので対処
                            Continue For
                        End If

                        '始点側路線検索
                        access_set1 = search_result.GetTrainList3(rc, search_date, 1)
                        If access_set1.Count > 0 Then
                            For item = 1 To access_set1.Count
                                access = access_set1.Item(item)
                                access_line_type = access.LineType
                                If line_type <> access_line_type Then
                                    ' 路線タイプが同じものだけ登録
                                    Continue For
                                End If
                                time_start = access.TimetableTime(1)
                                name = access.LongName
                                destination = ""

                                If name.Contains(line_name) Or access_line_type = 2 Then

                                    hit = True
                                    _set_train_list_from_access(line_train_list, added_train_list, train_id, access, line_name, line_id, direction, line_type)

                                End If
                            Next
                            If hit = False Then
                                ' hitしなければline_name_longでひっかける
                                For item = 1 To access_set1.Count
                                    access = access_set1.Item(item)
                                    access_line_type = access.LineType

                                    If line_type <> access_line_type Then
                                        ' 路線タイプが同じものだけ登録
                                        Continue For
                                    End If

                                    time_start = access.TimetableTime(1)
                                    name = access.LongName
                                    names = name.Split("・")
                                    hit_name = False
                                    For Each name In names
                                        For Each s In name
                                            If 0 <= Array.IndexOf(numbers_str, s) Then
                                                hit_name = True
                                                Exit For
                                            End If
                                        Next
                                        If hit_name Then
                                            Exit For
                                        End If
                                    Next

                                    destination = ""
                                    If line_name_long.Contains(name) Or access_line_type = 2 Then
                                        hit = True
                                        _set_train_list_from_access(line_train_list, added_train_list, train_id, access, line_name, line_id, direction, line_type)

                                    End If
                                Next
                            End If


                            If hit = False And search_hit = False Then
                                ' それでもヒットしない、かつ路線の検索ヒットがなかった場合、line_typeが同じものをすべて入れる
                                'For item = 1 To access_set1.Count
                                '    access = access_set1.Item(item)
                                '    access_line_type = access.LineType

                                '    If line_type <> access_line_type Then
                                '        ' 路線タイプが同じものだけ登録
                                '        Continue For
                                '    Else
                                '        hit = True
                                '        _set_train_list_from_access(line_train_list, added_train_list, train_id, access, line_name, line_id, direction)

                                '    End If

                                'Next


                            End If
                        End If
                        '終点側路線検索
                        access_set2 = search_result.GetTrainList3(rc, search_date, 2)
                        If access_set2.Count > 0 Then
                            For item = 1 To access_set2.Count
                                access = access_set2.Item(item)
                                access_line_type = access.LineType

                                If line_type <> access_line_type Then
                                    ' 路線タイプが同じものだけ登録
                                    Continue For
                                End If
                                time_start = access.TimetableTime(1)
                                name = access.LongName
                                destination = ""
                                If name.Contains(line_name) Or access_line_type = 2 Then

                                    hit = True
                                    _set_train_list_from_access(line_train_list, added_train_list, train_id, access, line_name, line_id, direction, line_type)

                                End If
                            Next
                            If hit = False Then
                                ' hitしなければline_name_longでひっかける
                                For item = 1 To access_set2.Count
                                    access = access_set2.Item(item)
                                    access_line_type = access.LineType

                                    If line_type <> access_line_type Then
                                        ' 路線タイプが同じものだけ登録
                                        Continue For
                                    End If
                                    time_start = access.TimetableTime(1)
                                    name = access.LongName
                                    destination = ""
                                    names = name.Split("・")
                                    hit_name = False
                                    For Each name In names
                                        For Each s In name
                                            If 0 <= Array.IndexOf(numbers_str, s) Then
                                                hit_name = True
                                                Exit For
                                            End If
                                        Next
                                        If hit_name Then
                                            Exit For
                                        End If
                                    Next

                                    If line_name_long.Contains(name) Or access_line_type = 2 Then

                                        hit = True
                                        _set_train_list_from_access(line_train_list, added_train_list, train_id, access, line_name, line_id, direction, line_type)


                                    End If
                                Next
                            End If

                            If hit = False And search_hit = False Then

                                ' それでもヒットしなければもうすべて入れる
                                'For item = 1 To access_set1.Count
                                '    access = access_set1.Item(item)
                                '    access_line_type = access.LineType

                                '    If line_type <> access_line_type Then
                                '        ' 路線タイプが同じものだけ登録
                                '        Continue For
                                '    Else
                                '        hit = True
                                '        _set_train_list_from_access(line_train_list, added_train_list, train_id, access, line_name, line_id, direction)

                                '    End If

                                'Next
                            End If
                        End If


                    Catch ex As Exception

                        Console.WriteLine(ex.Message & " :" & line_name)
                    End Try
                    If hit Then
                        search_hit = True
                        Exit For
                    End If
                Next
                If hit Then
                    Exit For
                End If
            Next
        End If


    End Function

    Private Function _search_train2(ByRef line_train_list, ByRef new_line_dict, ByRef added_train_dict, ByRef train_id, ByRef search_hit, from_station_name, to_station_name, line, iNavi, iSearch, data2, search_mode, direction, search_date)

        Dim list_search_od = New List(Of Station())()
        Dim line_name As String
        Dim line_name_long As String
        Dim line_type As Integer

        Dim time As Integer = 0
        Dim min As Integer = 0
        Dim hour As Integer = 0

        Dim line_id As Long
        Dim line_name_before As String = ""
        Dim search_results As EXPDENGNLib.ExpDiaCourseSet10
        Dim search_result As EXPDENGNLib.ExpDiaCourse10
        Dim init As Boolean

        Dim access_set1 As EXPDENGNLib.ExpDiaRouteAccessSet2
        Dim access_set2 As EXPDENGNLib.ExpDiaRouteAccessSet2
        Dim access As EXPDENGNLib.ExpDiaRouteAccess2

        Dim route_section As EXPDENGNLib.ExpDiaRouteSection3
        Dim departure_home As String
        Dim arrival_home As String

        Dim errcd As Integer
        Dim aa As Boolean

        Dim byte_limit As Integer = 1000000


        Dim fso As Scripting.FileSystemObject
        fso = New Scripting.FileSystemObject

        Dim row_diagram = ""
        Dim route_section_count As Integer

        Dim csv_out = ""
        Dim row_train_list = ""
        Dim numbers_str As Char() = New Char() {"０"c, "１"c, "２"c, "３"c, "４"c, "５"c, "６"c, "７"c, "８"c, "９"c}
        Dim line_name_new As String
        Dim line_id_new As Long
        Dim time_start
        Dim name
        Dim destination = ""
        Dim access_line_type
        Dim access_time_type
        Dim names
        Dim hit_name

        Dim hit = False

        line_name = line.line_name
        line_name_long = line.line_name_long
        line_id = line.line_id
        line_type = line.line_type

        aa = iNavi.RemoveAllKey

        aa = iNavi.AddKey(from_station_name, line.line_name_long)
        aa = iNavi.AddKey(to_station_name, line.line_name_long)

        errcd = iSearch.CheckNavi6(iNavi)   ' ●探索条件チェック

        search_results = iSearch.SearchCourse10(iNavi)
        init = True

        If (search_results.CourseCount > 0) Then
            ' ●探索結果数＞０の場合→探索結果が１件以上ある場合

            For cc = 1 To search_results.CourseCount

                search_result = search_results.GetCourse10(search_mode, cc)
                route_section_count = search_result.RouteSectionCount
                For rc = 1 To route_section_count
                    route_section = search_result.GetRouteSection3(rc)
                    Console.Write($"経路取得状況:{line_id}_line_count:{new_line_dict.Count} train_count:{added_train_dict.Count}                           ")
                    Console.CursorLeft = 0
                    departure_home = route_section.DeparturePlatform
                    arrival_home = route_section.ArrivalPlatform

                    Try
                        If route_section_count < rc Then
                            ' 再検索でroute_section_countが更新されることがあり、エラーになるケースがあるので対処
                            Continue For
                        End If

                        '始点側路線検索

                        access_set1 = search_result.GetTrainList3(rc, search_date, 1)
                        If access_set1.Count > 0 Then
                            hit = False
                            For item = 1 To access_set1.Count
                                access = access_set1.Item(item)
                                access_line_type = access.LineType
                                access_time_type = access.TimeType

                                If line_type <> access_line_type Then
                                    ' 路線タイプが同じものだけ登録
                                    Continue For
                                End If
                                time_start = access.TimetableTime(1)
                                name = access.LongName

                                If name.Contains(line_name) Or access_line_type = 2 Then

                                    hit = True
                                    ' 名前でフィルタリング
                                    line_name_new = name.Replace("・" + name.Split("・")(name.Split("・").Length - 1), "")
                                    If new_line_dict.ContainsKey(line_name_new) = False Then
                                        line_id_new = new_line_dict.Count
                                        new_line_dict.Add(line_name_new, {line_id_new, line_id})
                                    Else
                                        line_id_new = new_line_dict(line_name_new)(0)
                                    End If

                                    _set_train_list_from_access(line_train_list, added_train_dict, train_id, access, line_name_new, line_id_new, direction, line_type)

                                End If

                            Next
                           
                        End If

                        If hit = False then
                            '終点側路線検索
                            access_set2 = search_result.GetTrainList3(rc, search_date, 2)
                            For item = 1 To access_set2.Count
                                access = access_set2.Item(item)
                                access_line_type = access.LineType
                                If line_type <> access_line_type Then
                                    ' 路線タイプが同じものだけ登録
                                    Continue For
                                End If
                                time_start = access.TimetableTime(1)
                                name = access.LongName

                                If name.Contains(line_name) Or access_line_type = 2 Then

                                    hit = True
                                    ' 名前でフィルタリング
                                    line_name_new = name.Replace("・" + name.Split("・")(name.Split("・").Length - 1), "")
                                    If new_line_dict.ContainsKey(line_name_new) = False Then
                                        line_id_new = new_line_dict.Count
                                        new_line_dict.Add(line_name_new, {line_id_new, line_id})
                                    Else
                                        line_id_new = new_line_dict(line_name_new)(0)
                                    End If

                                    _set_train_list_from_access(line_train_list, added_train_dict, train_id, access, line_name_new, line_id_new, direction, line_type)

                                End If

                            Next

                        End if

                        If hit = False Then
                            ' それでもヒットしない場合は路線名無視で登録（路線名が違う場合も存在する）
                             For item = 1 To access_set1.Count
                                '引っかかったものをすべて登録し、新しい路線リストをつくる
                                access = access_set1.Item(item)
                                access_line_type = access.LineType
                                If line_type <> access_line_type Then
                                    ' 路線タイプが同じものだけ登録
                                    Continue For
                                End If

                                hit = True
                                time_start = access.TimetableTime(1)
                                name = access.LongName

                                ' 引っかかったものをすべて登録し、新しい路線リストをつくる
                                line_name_new = name.Replace("・" + name.Split("・")(name.Split("・").Length - 1), "")
                                If new_line_dict.ContainsKey(line_name_new) = False Then
                                    line_id_new = new_line_dict.Count
                                    new_line_dict.Add(line_name_new, {line_id_new, line_id})
                                Else
                                    line_id_new = new_line_dict(line_name_new)(0)
                                End If

                                _set_train_list_from_access(line_train_list, added_train_dict, train_id, access, line_name_new, line_id_new, direction, line_type)

                            Next
                        End If


                    Catch ex As Exception

                        Console.WriteLine(ex.Message & " :" & line_name)
                    End Try

                    If hit Then
                        search_hit = True
                        Exit For
                    End If
                Next
                If hit Then
                    Exit For
                End If
            Next
        End If


    End Function



    Private Function search_edge_dia(edge_list, export_edge_diagram_list, iNavi, iSearch, line_type_dict, export_path_edge_diagram, time_range, export_all_start_time, search_mode)

        Dim list_search_od = New List(Of Station())()
        Dim edge As Edge
        Dim line_name As String

        Dim from_station_name As String
        Dim to_station_name As String
        Dim from_station_type As Integer
        Dim to_station_type As Integer
        Dim time As Integer = 0
        Dim min As Integer = 0
        Dim hour As Integer = 0

        Dim line_id As Long
        Dim from_station_code As Long
        Dim to_station_code As Long
        Dim section_type_str As String
        Dim line_name_before As String = ""
        Dim line_name_short As String

        Dim board_time As Integer


        Dim ts As Scripting.TextStream


        Dim errcd As Integer
        Dim aa As Boolean

        Dim byte_limit As Integer = 1000000

        ' 検索の日付　ダイヤ更新日から半年程度以内にする（航空便のダイヤが入っていないため）
        Dim search_date As Long = 20190401

        Dim fso As Scripting.FileSystemObject
        fso = New Scripting.FileSystemObject

        Dim list_station = New List(Of Station)()

        Dim list_station_from = New List(Of Station)()

        Dim row_edge_diagram = ""


        Dim csv_out = ""

        'csv_station = "time, from_code,to_code,from_name,to_name,total_time,board_time,walk_time,other_time,fare_total,fare,fare_express,distance,line_count,airline, highway_bus"
        'csv_path = "time, from_code,to_code,from_name,to_name,station_code_path→"
        'csv_path_station_name = "time, from_code,to_code,from_name,to_name,station_name_path→"
        'csv_path_line = "time, from_code,to_code,from_name,to_name,line_path→"
        'csv_section_time = "time, from_code,to_code,from_name,to_name,path_time→"
        'csv_section_line_type = "time, from_code,to_code,from_name,to_name,line_path→"
        'csv_section_info = "time,from_code,to_code,from_name,to_name,section_count,section_from_code,section_to_code,section_from_time,section_to_time_section_line_name,section_line_type,section_fare,section_fare_express,section_distance,waiting_time,transfer"



        For i = 0 To edge_list.Count - 1
            ' 進捗を表示
            Console.CursorVisible = False

            edge = edge_list(i)
            from_station_code = edge.from_station_code
            to_station_code = edge.to_station_code

            from_station_type = data2.GetStationByCode(from_station_code).Type
            to_station_type = data2.GetStationByCode(to_station_code).Type

            line_name = edge.line_name
            line_id = edge.line_id

            If line_name_before <> "" And line_name <> line_name_before Then
                Console.Clear()
            End If
            Console.Write("経路取得状況:" & Str(i) & " / " & Str(edge_list.Count) & " :" & line_name)
            Console.SetCursorPosition(0, Console.CursorTop)
            line_name_before = line_name


            If from_station_code = to_station_code Then
                Continue For
            End If

            'line_nameがLongだとヒットしないことがあるので、line_name_shortを取得する
            line_name = iSearch.GetLine(line_name, search_date).ShortName

            iNavi.LocalBusOnly = False     ' ●路線バスのみでの検索 鉄道駅で検索するときは基本False　バス停同士での検索時のみTrueにする
            'If from_station_type = 8 And to_station_type = 8 Then
            'iNavi.LocalBusOnly = True
            ' バス路線には会社名が"・"区切りで入っているので除去
            'line_name = line_name.Split("・")(1)
            'End If

            from_station_name = data2.GetStationNameByCode(from_station_code)
            to_station_name = data2.GetStationNameByCode(to_station_code)

            ' 時刻指定時の検索
            If time_range.length = 0 Then
                time_range = {0, 0, 1}
                time = 0
            Else
                min = (time_range(0) - Int(time_range(0))) * 60
                hour = Int(time_range(0))

                time = hour * 100 + min
            End If

            If time > 0 Then
                iNavi.Date = search_date
                iNavi.Time = time
            End If


            _search_od(export_edge_diagram_list, from_station_name, to_station_name, from_station_code, to_station_code, time, line_id, line_name, iNavi, iSearch, search_mode)
            '逆方向
            _search_od(export_edge_diagram_list, to_station_name, from_station_name, to_station_code, from_station_code, time, line_id, line_name, iNavi, iSearch, search_mode)

            'サイズが大きくなったらいったん出力
            If export_edge_diagram_list.Count > 100000 Then
                Console.Clear()

                For l = 0 To export_edge_diagram_list.Count - 1
                    Console.Write("一旦出力:" & Str(l) & " / " & Str(export_edge_diagram_list.Count))
                    Console.SetCursorPosition(0, Console.CursorTop)

                    row_edge_diagram = export_edge_diagram_list(l)
                    csv_out = csv_out & row_edge_diagram & vbCrLf
                    export_csv_part(csv_out, export_path_edge_diagram, byte_limit, ts, fso, False, "")

                Next

                export_csv_part(csv_out, export_path_edge_diagram, byte_limit, ts, fso, True, "")

                '出力リスト初期化
                export_edge_diagram_list = New List(Of String)
            End If

        Next

        ' 出力
        Console.Clear()

        For l = 0 To export_edge_diagram_list.Count - 1
            Console.Write("最終出力:" & Str(l) & " / " & Str(export_edge_diagram_list.Count))
            Console.SetCursorPosition(0, Console.CursorTop)

            row_edge_diagram = export_edge_diagram_list(l)
            csv_out = csv_out & row_edge_diagram & vbCrLf
            export_csv_part(csv_out, export_path_edge_diagram, byte_limit, ts, fso, False, "")

        Next

        export_csv_part(csv_out, export_path_edge_diagram, byte_limit, ts, fso, True, "")

        Console.CursorVisible = True

    End Function

    Sub GetSearchResult_Line(iSearch, data2, read_path, read_path_station,read_path_all_station, export_path, search_mode, line_type_dict, database_version)

        ' ●ExpDiaNavi6オブジェクトを構築
        Dim iNavi As EXPDENGNLib.ExpDiaNavi6

        iNavi = iSearch.CreateNavi6()

        Dim list_search_od = New List(Of Station())()

        Dim byte_limit As Integer = 1000000
        ' 検索の日付　ダイヤ更新日から半年程度以内にする（航空便のダイヤが入っていないため）

        Dim search_date As Long = database_version + 1
        Dim fso As Scripting.FileSystemObject
        fso = New Scripting.FileSystemObject

        Dim tr As Scripting.TextStream
        Dim row_str As String
        Dim row As String()
        Dim line_row As Line
        Dim list_line = New List(Of Line)()
        Dim station_code As Long
        Dim list_station_code = New List(Of Long)()
        Dim station_in_line_dict As New Dictionary(Of Integer, List(Of Long))()
        Dim station_in_line_list As New List(Of Long)
        Dim before_line_id As New Integer
        Dim line_id As New Integer
        Dim export_edge_diagram_list As List(Of String) = New List(Of String)


        ' lineデータ読み込み
        tr = fso.OpenTextFile(read_path, 1)
        Dim init As Boolean = True
        Do While tr.AtEndOfStream = False
            row_str = tr.ReadLine
            If init = False Then
                Try
                    row = Split(row_str, ",")
                    'line_id,line_name,line_name_long,line_name_short,line_type,station_count,corp_id,corp_name,line_type
                    '    Sub New(lineId As Long, lineName As String, lineName_L As String, lineType As Integer, sCount As Integer, corpId As Integer, corpName As String)
                    If row.Length > 1 Then
                        line_row = New Line(Long.Parse(row(0)), row(3), row(2), Integer.Parse(row(4)), Integer.Parse(row(5)), Integer.Parse(row(6)), row(7))
                        list_line.Add(line_row)
                    End If

                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                End Try

            Else
                init = False '一行目のheader飛ばす
            End If
        Loop

        tr.Close()

        ' all_line_stationデータ読み込み
        tr = fso.OpenTextFile(read_path_all_station, 1)
        init = True
        before_line_id = -1

        Do While tr.AtEndOfStream = False
            row_str = tr.ReadLine
            If init = False Then
                Try
                    row = Split(row_str, ",")
                    'l"station_code,station_name,station_name_long,station_type,pass_line_count,corp_id,line_id,line_name,order"

                    If row.Length > 1 Then
                       
                        line_id =  Integer.Parse(row(6))
                        station_code =  Integer.Parse(row(0))
                        If before_line_id <> line_id And before_line_id > 0 Then
                            station_in_line_dict.Add(before_line_id, New List(Of Long)(station_in_line_list))
                            station_in_line_list.Clear()
                        End if

                        station_in_line_list.Add(station_code)
                        before_line_id = line_id
                    End If

                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                End Try

            Else
                init = False '一行目のheader飛ばす
            End If
        Loop

        tr.Close()


        fso = New Scripting.FileSystemObject

        ' ●ExpDiaNavi6オブジェクトに探索条件を設定
        iNavi.AnswerCount = 20         ' ●探索経路要求回答数
        iNavi.LocalBus = True          ' ●路線バスを利用
        'iNavi.FuzzyLine = True         ' ●路線名あいまい指定を許可
        ' TODO line_typeで切替
        'iNavi.DiaAirplane = True

        ' 一旦時刻指定
        iNavi.Date = search_date
        iNavi.Time = 500

        search_train_on_line(list_line, station_in_line_dict, iNavi, iSearch, line_type_dict, export_path, search_mode, search_date)


    End Sub

    Sub GetSearchResult_Edge(iSearch, data2, read_path, export_path_path_station_name, time_range, export_all_start_time, search_mode, line_type_dict)


        ' ●ExpDiaNavi6オブジェクトを構築
        Dim iNavi As EXPDENGNLib.ExpDiaNavi6

        iNavi = iSearch.CreateNavi6()

        Dim list_search_od = New List(Of Station())()

        Dim byte_limit As Integer = 1000000
        ' 検索の日付　ダイヤ更新日から半年程度以内にする（航空便のダイヤが入っていないため）
        Dim search_date As Long = 20180402


        Dim fso As Scripting.FileSystemObject
        fso = New Scripting.FileSystemObject

        Dim tr As Scripting.TextStream
        Dim row_str As String
        Dim row As String()
        Dim edge_row As Edge
        Dim list_edge = New List(Of Edge)()

        Dim export_edge_diagram_list As List(Of String) = New List(Of String)

        Dim od_list As List(Of Station()) = New List(Of Station())
        Dim edge_diagram As String


        ' データ読み込み
        tr = fso.OpenTextFile(read_path, 1)
        Dim init As Boolean = True
        Do While tr.AtEndOfStream = False
            row_str = tr.ReadLine
            If init = False Then
                Try
                    row = Split(row_str, ",")
                    edge_row = New Edge(Long.Parse(row(0)), Long.Parse(row(2)), Long.Parse(row(4)), row(5), Integer.Parse(row(6)))
                    list_edge.Add(edge_row)
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                End Try

            Else
                init = False '一行目のheader飛ばす
            End If
        Loop

        tr.Close()

        fso = New Scripting.FileSystemObject

        ' ●ExpDiaNavi6オブジェクトに探索条件を設定
        iNavi.AnswerCount = 20         ' ●探索経路要求回答数
        iNavi.LocalBus = True          ' ●路線バスを利用
        'iNavi.FuzzyLine = True         ' ●路線名あいまい指定を許可


        search_edge_dia(list_edge, export_edge_diagram_list, iNavi, iSearch, line_type_dict, export_path_path_station_name, time_range, export_all_start_time, search_mode)


    End Sub

    Sub Main()

        'iSearch = New EXPDENGNLib.ExpDiaSearch10
        iSearch = New EXPDENGNLib.ExpDiaDB10
        data2 = New EXPDENGNLib.ExpData2


        ' ●ExpDiaNavi6オブジェクトを構築
        Dim iNavi As EXPDENGNLib.ExpDiaNavi6
        iNavi = iSearch.CreateNavi6()

        ' バージョン取得
        Dim database_version = iSearch.KNBVersion

        ' 入力データタイプ　0:matrix 1:od list
        Dim input_data_type As Integer = 1

        ' edgeの通過時間取得（すべてのものを取得するのは無理だった　特定の都市を対象とするなら良いかも）
        'Dim read_path_from As String = "C:\projects\expwin\EXP_search\output\database\201807\edge.csv"
        'Dim export_path_base As String = "C:\projects\expwin\EXP_search\output\database\201807\edge_diagram\"

        '路線から時刻表取得
        Dim data_base_dir = $"D:\exp\database\{database_version}"
        Dim read_path_from As String = $"{data_base_dir}\line.csv"
        Dim read_path_station As String = $"{data_base_dir}\station.csv"
        Dim read_path_all_station As String = $"{data_base_dir}\all_station.csv"
        Dim export_path_base As String = $"{data_base_dir}\train_diagram"


        'test
        'export_path_base = "C:\projects\expwin\EXP_search\output\database\train_diagram\20210417\"

        ' 路線時刻表の取得なら1、OD（edge）時刻表取得なら0
        Dim is_line_search_mode = 1

        ' 検索結果　1: 探索順, 2: 所要時間順, 4: 運賃
        Dim search_mode As Integer = 2

        ' 起点駅コードの指定　-1ならすべての組み合わせ
        Dim origin_station_code As Integer = -1

        ' 時間検索するか　するなら1
        Dim time_search As Integer = 1

        'すべての時間で検索かけるなら１
        Dim export_all_start_time = 0

        'バス路線ダイヤ取得 0なら両方、1ならバス路線探索, 2ならバス路線以外
        Dim search_local_bus = 1

        'コマンドライン引数から取得
        Dim cmds As String() = System.Environment.GetCommandLineArgs()
        If cmds.Length >= 2 Then
            search_local_bus = cmds(1)
        End if
        

        Console.WriteLine("edge:" & read_path_from)
        Console.WriteLine("export path:" & export_path_base)
        Console.WriteLine("search mode:" & search_mode)
        Console.WriteLine("is_line_search_mode:" & is_line_search_mode)


        If Dir(export_path_base, vbDirectory) = "" Then
            MkDir(export_path_base)
        End If

        Dim line_type_dict As New Dictionary(Of Integer, String)()


        If search_local_bus = 0 Or search_local_bus = 2 Then
            line_type_dict.Add(1, "鉄道")
            line_type_dict.Add(2, "空路")
            line_type_dict.Add(4, "連絡バス")
            line_type_dict.Add(8, "フェリー")
            line_type_dict.Add(64, "高速バス")
            line_type_dict.Add(256, "その他")
        End If
        
        If search_local_bus = 0 Or search_local_bus = 1 Then
            line_type_dict.Add(32, "路線バス")
        End If


        Dim export_path As String
        If is_line_search_mode = 1 Then
            GetSearchResult_Line(iSearch, data2, read_path_from, read_path_station,read_path_all_station, export_path_base, search_mode, line_type_dict, database_version)

        Else
            '激遅なので使わない
            'Dim export_path_edge_diagram As String = export_path_base & "edge_diagram"
            'GetSearchResult_Edge(iSearch, data2, read_path_from, export_path, time_range, export_all_start_time, search_mode, line_type_dict)

        End If

    End Sub
End Module
