Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Net.Mail
Imports System.Text
Imports System.Text.Json
Imports MySqlConnector

Module Program
    Private Const LocalConfigFile As String = "appsettings.local.json"
    Private Const ExampleConfigFile As String = "appsettings.local.json.example"

    Private Class AppSettings
        Public Property Mysql As New MysqlSettings()
        Public Property Report As New ReportSettings()
    End Class

    Private Class MysqlSettings
        Public Property ConnectionString As String = ""
    End Class

    Private Class SmtpSettings
        Public Property Host As String = ""
        Public Property Port As Integer = 25
        Public Property EnableSsl As Boolean
        Public Property Username As String = ""
        Public Property Password As String = ""
        Public Property FromAddress As String = ""
        Public Property FromDisplayName As String = ""
        Public Property ToAddresses As New List(Of String)()
        Public Property CcAddresses As New List(Of String)()
    End Class

    Private Class ReportSettings
        Public Property HotelIds As New List(Of String)()
        Public Property SubjectTimestampFormat As String = "MM/dd HH:mm"
        Public Property DetailPastDays As Integer = 1
        Public Property DetailFutureDays As Integer = 120
        Public Property FutureBucketDays As Integer() = {30, 60, 90}
        Public Property MonthCount As Integer = 4
        Public Property InventoryIndex As Integer = 6
        Public Property FaultIndex As Integer = 4
        Public Property RoomTypeOrder As New List(Of String) From {
            "DD", "DT", "ED", "EF", "HD", "S", "SD", "ST", "TP", "TR"
        }
        Public Property OutputDir As String = "output"
    End Class

    Private Class HotelInfo
        Public Property HotelId As String = ""
        Public Property HotelName As String = ""
    End Class

    Private Class ReportRow
        Public Property DataDate As DateTime
        Public Property RoomCount As Integer
        Public Property RoomRevenue As Decimal
        Public Property RoomOccupancyText As String = ""
        Public Property RoomOccupancyValue As Decimal?
        Public Property RoomAdr As Decimal
        Public Property RoomData As String = ""
        Public Property DataLmtime As DateTime?
        Public Property RoomTypeData As New Dictionary(Of String, Integer())(StringComparer.OrdinalIgnoreCase)
    End Class

    Private Class EmailPayload
        Public Property Subject As String = ""
        Public Property Html As String = ""
        Public Property OutputPath As String = ""
    End Class

    Sub Main(args As String())
        Console.OutputEncoding = Encoding.UTF8

        Try
            Dim options = ParseOptions(args)
            Dim settings = LoadSettings()
            Dim baseDate = If(options.BaseDate.HasValue, options.BaseDate.Value.Date, Date.Today)
            Dim outputDir = ResolveOutputDir(settings.Report.OutputDir)
            Directory.CreateDirectory(outputDir)

            Using conn As New MySqlConnection(settings.Mysql.ConnectionString)
                conn.Open()

                Dim hotels = LoadHotels(conn, settings.Report.HotelIds, options.HotelId)
                If hotels.Count = 0 Then
                    Throw New InvalidOperationException("No hotels found to email.")
                End If

                Dim senderSettings = LoadSenderSettings(conn)

                Console.WriteLine($"Hotels to email: {hotels.Count}")
                Console.WriteLine($"Base date: {baseDate:yyyy/MM/dd}")
                Console.WriteLine($"SMTP sender loaded: {senderSettings IsNot Nothing}")

                For Each hotel In hotels
                    Console.WriteLine($"--- Hotel {hotel.HotelId} {hotel.HotelName} ---")

                    Dim rows = LoadHotelRows(conn, hotel.HotelId, baseDate, settings.Report)
                    If rows.Count = 0 Then
                        Console.WriteLine("No report data found, skipped.")
                        Continue For
                    End If

                    Dim payload = BuildEmailPayload(hotel, rows, baseDate, settings.Report, outputDir)
                    File.WriteAllText(payload.OutputPath, payload.Html, Encoding.UTF8)
                    Console.WriteLine($"HTML preview: {payload.OutputPath}")
                    Console.WriteLine($"Subject: {payload.Subject}")

                    Dim recipientSettings = LoadRecipientSettings(conn, hotel.HotelId)
                    Console.WriteLine($"Recipient loaded: {recipientSettings IsNot Nothing}")

                    If Not options.PreviewOnly Then
                        If senderSettings Is Nothing Then
                            Throw New InvalidOperationException("Missing athena_setting sender row: func_name='email_sender', is_use='Y'")
                        End If
                        If recipientSettings Is Nothing Then
                            Throw New InvalidOperationException($"Missing athena_setting recipient row for hotel {hotel.HotelId}: func_name='email_recipient', is_use='Y'")
                        End If

                        Dim smtpSettings = MergeMailSettings(senderSettings, recipientSettings)
                        SendEmail(smtpSettings, payload.Subject, payload.Html)
                        Console.WriteLine("Email sent.")
                    Else
                        Console.WriteLine("Preview only, email not sent.")
                    End If
                Next
            End Using
        Catch ex As Exception
            Console.Error.WriteLine(ex.ToString())
            Environment.ExitCode = 1
        End Try
    End Sub

    Private Function ParseOptions(args As String()) As (PreviewOnly As Boolean, HotelId As String, BaseDate As DateTime?)
        Dim previewOnly = False
        Dim hotelId = ""
        Dim baseDate As DateTime? = Nothing
        Dim i = 0

        While i < args.Length
            Select Case args(i)
                Case "--preview-only"
                    previewOnly = True
                Case "--hotel"
                    If i + 1 >= args.Length Then
                        Throw New ArgumentException("--hotel requires a hotel id")
                    End If
                    hotelId = args(i + 1).Trim()
                    i += 1
                Case "--date"
                    If i + 1 >= args.Length Then
                        Throw New ArgumentException("--date requires yyyy/MM/dd")
                    End If
                    baseDate = DateTime.ParseExact(args(i + 1), "yyyy/MM/dd", CultureInfo.InvariantCulture)
                    i += 1
                Case Else
                    Throw New ArgumentException($"Unknown argument: {args(i)}")
            End Select

            i += 1
        End While

        Return (previewOnly, hotelId, baseDate)
    End Function

    Private Function LoadSettings() As AppSettings
        Dim candidatePaths = New String() {
            Path.Combine(Environment.CurrentDirectory, LocalConfigFile),
            Path.Combine(AppContext.BaseDirectory, LocalConfigFile),
            Path.Combine(AppContext.BaseDirectory, "..", "..", "..", LocalConfigFile),
            Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", LocalConfigFile)
        }.Select(Function(x) Path.GetFullPath(x)).Distinct().ToList()

        Dim localPath = candidatePaths.FirstOrDefault(Function(x) File.Exists(x))
        If String.IsNullOrWhiteSpace(localPath) Then
            Throw New FileNotFoundException($"Missing {LocalConfigFile}. Copy {ExampleConfigFile} and fill it first.", candidatePaths.First())
        End If

        Dim json = File.ReadAllText(localPath, Encoding.UTF8)
        Dim settings = JsonSerializer.Deserialize(Of AppSettings)(json, New JsonSerializerOptions With {
            .PropertyNameCaseInsensitive = True
        })

        If settings Is Nothing Then
            Throw New InvalidOperationException("Failed to load app settings.")
        End If

        If String.IsNullOrWhiteSpace(settings.Mysql.ConnectionString) Then
            Throw New InvalidOperationException("Mysql.ConnectionString is required.")
        End If

        If settings.Report.FutureBucketDays Is Nothing OrElse settings.Report.FutureBucketDays.Length = 0 Then
            settings.Report.FutureBucketDays = {30, 60, 90}
        End If

        Return settings
    End Function

    Private Function ResolveOutputDir(outputDir As String) As String
        If String.IsNullOrWhiteSpace(outputDir) Then
            outputDir = "output"
        End If

        If Path.IsPathRooted(outputDir) Then
            Return outputDir
        End If

        Return Path.Combine(AppContext.BaseDirectory, outputDir)
    End Function

    Private Function LoadHotels(conn As MySqlConnection, configuredHotelIds As List(Of String), forcedHotelId As String) As List(Of HotelInfo)
        Dim hotels As New List(Of HotelInfo)()
        Dim filterIds = New List(Of String)()

        If configuredHotelIds IsNot Nothing AndAlso configuredHotelIds.Count > 0 Then
            filterIds.AddRange(configuredHotelIds.Where(Function(x) Not String.IsNullOrWhiteSpace(x)).Select(Function(x) x.Trim().ToUpperInvariant()))
        End If

        If Not String.IsNullOrWhiteSpace(forcedHotelId) Then
            filterIds.Clear()
            filterIds.Add(forcedHotelId.Trim().ToUpperInvariant())
        End If

        Dim sql = New StringBuilder()
        sql.AppendLine("SELECT hotel_id, COALESCE(hotel_name, hotel_id) AS hotel_name")
        sql.AppendLine("FROM athena_info")
        sql.AppendLine("WHERE UPPER(COALESCE(is_fetch, 'N')) = 'Y'")

        If filterIds.Count > 0 Then
            Dim paramNames = Enumerable.Range(0, filterIds.Count).Select(Function(idx) $"@p{idx}")
            sql.AppendLine($"AND hotel_id IN ({String.Join(", ", paramNames)})")
        End If

        sql.AppendLine("ORDER BY hotel_id")

        Using cmd As New MySqlCommand(sql.ToString(), conn)
            For i = 0 To filterIds.Count - 1
                cmd.Parameters.AddWithValue($"@p{i}", filterIds(i))
            Next

            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    hotels.Add(New HotelInfo With {
                        .HotelId = reader.GetString("hotel_id"),
                        .HotelName = reader.GetString("hotel_name")
                    })
                End While
            End Using
        End Using

        Return hotels
    End Function

    Private Function LoadSenderSettings(conn As MySqlConnection) As SmtpSettings
        Const sql As String = "
SELECT value
FROM athena_setting
WHERE func_name = 'email_sender'
  AND UPPER(COALESCE(is_use, 'N')) = 'Y'
ORDER BY hotel_id
LIMIT 1
"
        Using cmd As New MySqlCommand(sql, conn)
            Dim raw = Convert.ToString(cmd.ExecuteScalar())
            If String.IsNullOrWhiteSpace(raw) Then
                Return Nothing
            End If
            Return ParseSenderValue(raw)
        End Using
    End Function

    Private Function LoadRecipientSettings(conn As MySqlConnection, hotelId As String) As SmtpSettings
        Const sql As String = "
SELECT value
FROM athena_setting
WHERE func_name = 'email_recipient'
  AND hotel_id = @hotel_id
  AND UPPER(COALESCE(is_use, 'N')) = 'Y'
LIMIT 1
"
        Using cmd As New MySqlCommand(sql, conn)
            cmd.Parameters.AddWithValue("@hotel_id", hotelId)
            Dim raw = Convert.ToString(cmd.ExecuteScalar())
            If String.IsNullOrWhiteSpace(raw) Then
                Return Nothing
            End If
            Return ParseRecipientValue(raw)
        End Using
    End Function

    Private Function ParseSenderValue(raw As String) As SmtpSettings
        raw = raw.Trim()

        If raw.StartsWith("{") Then
            Dim parsed = JsonSerializer.Deserialize(Of SmtpSettings)(raw, New JsonSerializerOptions With {
                .PropertyNameCaseInsensitive = True
            })
            If parsed Is Nothing Then
                Throw New InvalidOperationException("athena_setting email_sender JSON is invalid.")
            End If
            Return parsed
        End If

        Dim settings As New SmtpSettings()
        For Each segment In raw.Split(";"c, StringSplitOptions.RemoveEmptyEntries)
            Dim parts = segment.Split("="c, 2)
            If parts.Length <> 2 Then
                Continue For
            End If
            Dim key = parts(0).Trim().ToLowerInvariant()
            Dim value = parts(1).Trim()
            Select Case key
                Case "host"
                    settings.Host = value
                Case "port"
                    settings.Port = ParseInt(value)
                Case "enablessl", "ssl"
                    settings.EnableSsl = value.Equals("true", StringComparison.OrdinalIgnoreCase) OrElse value.Equals("y", StringComparison.OrdinalIgnoreCase)
                Case "username"
                    settings.Username = value
                Case "password"
                    settings.Password = value
                Case "fromaddress", "from"
                    settings.FromAddress = value
                Case "fromdisplayname", "displayname"
                    settings.FromDisplayName = value
            End Select
        Next
        Return settings
    End Function

    Private Function ParseRecipientValue(raw As String) As SmtpSettings
        raw = raw.Trim()

        If raw.StartsWith("{") Then
            Dim parsed = JsonSerializer.Deserialize(Of SmtpSettings)(raw, New JsonSerializerOptions With {
                .PropertyNameCaseInsensitive = True
            })
            If parsed Is Nothing Then
                Throw New InvalidOperationException("athena_setting email_recipient JSON is invalid.")
            End If
            Return parsed
        End If

        Dim settings As New SmtpSettings()
        If raw.Contains("=") Then
            For Each segment In raw.Split(";"c, StringSplitOptions.RemoveEmptyEntries)
                Dim parts = segment.Split("="c, 2)
                If parts.Length <> 2 Then
                    Continue For
                End If
                Dim key = parts(0).Trim().ToLowerInvariant()
                Dim value = parts(1).Trim()
                Select Case key
                    Case "to"
                        settings.ToAddresses = SplitAddresses(value)
                    Case "cc"
                        settings.CcAddresses = SplitAddresses(value)
                End Select
            Next
        Else
            settings.ToAddresses = SplitAddresses(raw)
        End If

        Return settings
    End Function

    Private Function SplitAddresses(raw As String) As List(Of String)
        Return raw.Split(New Char() {";"c, ","c}, StringSplitOptions.RemoveEmptyEntries).
            Select(Function(x) x.Trim()).
            Where(Function(x) Not String.IsNullOrWhiteSpace(x)).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()
    End Function

    Private Function MergeMailSettings(senderSettings As SmtpSettings, recipientSettings As SmtpSettings) As SmtpSettings
        Dim fromAddress = senderSettings.FromAddress
        If String.IsNullOrWhiteSpace(fromAddress) Then
            fromAddress = senderSettings.Username
        End If

        Return New SmtpSettings With {
            .Host = senderSettings.Host,
            .Port = senderSettings.Port,
            .EnableSsl = senderSettings.EnableSsl,
            .Username = senderSettings.Username,
            .Password = senderSettings.Password,
            .FromAddress = fromAddress,
            .FromDisplayName = senderSettings.FromDisplayName,
            .ToAddresses = recipientSettings.ToAddresses,
            .CcAddresses = recipientSettings.CcAddresses
        }
    End Function

    Private Function LoadHotelRows(conn As MySqlConnection, hotelId As String, baseDate As DateTime, settings As ReportSettings) As List(Of ReportRow)
        Dim startDate = baseDate.AddDays(-Math.Max(settings.DetailPastDays, 1))
        Dim futureDays = Math.Max(settings.DetailFutureDays, settings.FutureBucketDays.Max())
        Dim endDate = baseDate.AddDays(futureDays)

        Dim rows As New List(Of ReportRow)()
        Dim sql = "
SELECT
    data_date,
    room_count,
    room_revenue,
    room_occupancy,
    room_adr,
    room_data,
    data_lmtime
FROM athena_report01
WHERE hotel_id = @hotel_id
  AND STR_TO_DATE(data_date, '%Y/%m/%d') BETWEEN @start_date AND @end_date
ORDER BY STR_TO_DATE(data_date, '%Y/%m/%d')
"

        Using cmd As New MySqlCommand(sql, conn)
            cmd.Parameters.AddWithValue("@hotel_id", hotelId)
            cmd.Parameters.AddWithValue("@start_date", startDate)
            cmd.Parameters.AddWithValue("@end_date", endDate)

            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim row = New ReportRow With {
                        .DataDate = DateTime.ParseExact(reader.GetString("data_date"), "yyyy/MM/dd", CultureInfo.InvariantCulture),
                        .RoomCount = ParseInt(reader("room_count")),
                        .RoomRevenue = ParseDecimal(reader("room_revenue")),
                        .RoomOccupancyValue = ParsePercent(reader("room_occupancy").ToString()),
                        .RoomAdr = ParseDecimal(reader("room_adr")),
                        .RoomData = reader("room_data").ToString(),
                        .DataLmtime = ParseDateTimeNullable(reader("data_lmtime").ToString())
                    }
                    row.RoomOccupancyText = FormatPercentText(row.RoomOccupancyValue)
                    row.RoomTypeData = ParseRoomData(row.RoomData)
                    rows.Add(row)
                End While
            End Using
        End Using

        Return rows
    End Function

    Private Function BuildEmailPayload(hotel As HotelInfo, rows As List(Of ReportRow), baseDate As DateTime, settings As ReportSettings, outputDir As String) As EmailPayload
        Dim rowMap = rows.ToDictionary(Function(x) x.DataDate, Function(x) x)
        Dim todayRow = FindNearestRow(rowMap, baseDate)
        Dim yesterdayRow = FindNearestRow(rowMap, baseDate.AddDays(-1))
        Dim latestTime = rows.Where(Function(x) x.DataLmtime.HasValue).Select(Function(x) x.DataLmtime.Value).DefaultIfEmpty(DateTime.Now).Max()
        Dim sendTime = DateTime.Now

        Dim roomTypes = ResolveRoomTypes(rows, settings.RoomTypeOrder)
        Dim summaryBuckets = BuildBucketSummary(rows, baseDate, settings.FutureBucketDays)
        Dim monthlySummary = BuildMonthlySummary(rows, baseDate, settings.MonthCount)
        Dim detailRows = rows.Where(Function(x) x.DataDate >= baseDate.AddDays(-settings.DetailPastDays) AndAlso x.DataDate <= baseDate.AddDays(settings.DetailFutureDays)).ToList()

        Dim subject = $"{hotel.HotelName}*本日:{todayRow.RoomCount}({todayRow.RoomOccupancyText})*昨日:{yesterdayRow.RoomCount}({yesterdayRow.RoomOccupancyText}) - {sendTime.ToString(settings.SubjectTimestampFormat)}"
        Dim html = BuildHtml(hotel, subject, latestTime, summaryBuckets, monthlySummary, detailRows, roomTypes, settings)
        Dim outputPath = Path.Combine(outputDir, $"{hotel.HotelId}_{baseDate:yyyyMMdd}.html")

        Return New EmailPayload With {
            .Subject = subject,
            .Html = html,
            .OutputPath = outputPath
        }
    End Function

    Private Function FindNearestRow(rowMap As Dictionary(Of DateTime, ReportRow), targetDate As DateTime) As ReportRow
        If rowMap.ContainsKey(targetDate) Then
            Return rowMap(targetDate)
        End If

        Return New ReportRow With {
            .DataDate = targetDate,
            .RoomCount = 0,
            .RoomOccupancyText = "0%",
            .RoomRevenue = 0D,
            .RoomAdr = 0D,
            .RoomTypeData = New Dictionary(Of String, Integer())(StringComparer.OrdinalIgnoreCase)
        }
    End Function

    Private Function ResolveRoomTypes(rows As List(Of ReportRow), configuredOrder As List(Of String)) As List(Of String)
        Dim allTypes = rows.SelectMany(Function(x) x.RoomTypeData.Keys).Distinct(StringComparer.OrdinalIgnoreCase).ToList()
        Dim ordered As New List(Of String)()

        If configuredOrder IsNot Nothing Then
            For Each roomType In configuredOrder
                If allTypes.Any(Function(x) String.Equals(x, roomType, StringComparison.OrdinalIgnoreCase)) Then
                    ordered.Add(roomType)
                End If
            Next
        End If

        For Each extra In allTypes.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase)
            If Not ordered.Any(Function(x) String.Equals(x, extra, StringComparison.OrdinalIgnoreCase)) Then
                ordered.Add(extra)
            End If
        Next

        Return ordered
    End Function

    Private Function BuildBucketSummary(rows As List(Of ReportRow), baseDate As DateTime, bucketDays As Integer()) As List(Of (Label As String, Value As String, Color As String))
        Dim items As New List(Of (String, String, String))()
        Dim previous = 0

        For Each dayCount In bucketDays
            Dim startDate = baseDate.AddDays(previous + 1)
            Dim endDate = baseDate.AddDays(dayCount)
            Dim values = rows.Where(Function(x) x.DataDate >= startDate AndAlso x.DataDate <= endDate AndAlso x.RoomOccupancyValue.HasValue).Select(Function(x) x.RoomOccupancyValue.Value).ToList()
            Dim avg = If(values.Count = 0, 0D, values.Average())
            items.Add(($"{previous + 1}~{dayCount}天", $"{Math.Round(avg)}%", OccupancyColor(avg)))
            previous = dayCount
        Next

        Return items
    End Function

    Private Function BuildMonthlySummary(rows As List(Of ReportRow), baseDate As DateTime, monthCount As Integer) As List(Of (Label As String, Value As String, Color As String))
        Dim items As New List(Of (String, String, String))()
        Dim firstMonth = New DateTime(baseDate.Year, baseDate.Month, 1)

        For i = 0 To Math.Max(monthCount, 1) - 1
            Dim monthStart = firstMonth.AddMonths(i)
            Dim monthEnd = monthStart.AddMonths(1).AddDays(-1)
            Dim values = rows.Where(Function(x) x.DataDate >= monthStart AndAlso x.DataDate <= monthEnd AndAlso x.RoomOccupancyValue.HasValue).Select(Function(x) x.RoomOccupancyValue.Value).ToList()
            Dim avg = If(values.Count = 0, 0D, values.Average())
            items.Add(($"{monthStart:yyyy/MM}", $"{Math.Round(avg)}%", OccupancyColor(avg)))
        Next

        Return items
    End Function

    Private Function BuildHtml(hotel As HotelInfo,
                               subject As String,
                               latestTime As DateTime,
                               bucketSummary As List(Of (Label As String, Value As String, Color As String)),
                               monthlySummary As List(Of (Label As String, Value As String, Color As String)),
                               detailRows As List(Of ReportRow),
                               roomTypes As List(Of String),
                               settings As ReportSettings) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("<!DOCTYPE html>")
        sb.AppendLine("<html><head><meta charset=""utf-8"">")
        sb.AppendLine("<style>")
        sb.AppendLine("body{font-family:'Microsoft JhengHei',Arial,sans-serif;color:#111;margin:20px;}")
        sb.AppendLine("h1{font-size:24px;margin:0 0 10px 0;font-weight:700;}")
        sb.AppendLine(".sub{font-size:14px;color:#666;margin-bottom:16px;}")
        sb.AppendLine("table{border-collapse:collapse;margin:14px 0;}")
        sb.AppendLine("th,td{border:1px solid #888;padding:4px 8px;text-align:center;font-size:14px;white-space:nowrap;}")
        sb.AppendLine("th{background:#fafafa;}")
        sb.AppendLine(".label{background:#fff;font-weight:700;}")
        sb.AppendLine(".small-note{font-size:12px;color:#444;margin-top:4px;}")
        sb.AppendLine(".left{text-align:left;}")
        sb.AppendLine(".red{color:#d90000;font-weight:700;}")
        sb.AppendLine(".green{color:#0a8f08;font-weight:700;}")
        sb.AppendLine("</style></head><body>")
        sb.AppendLine($"<h1>{WebUtility.HtmlEncode(subject)}</h1>")
        sb.AppendLine($"<div class=""sub"">飯店: {WebUtility.HtmlEncode(hotel.HotelName)} / 最後抓取: {latestTime:yyyy/MM/dd HH:mm:ss}</div>")

        AppendSummaryTable(sb, "未來天數", "住房率", bucketSummary)
        AppendSummaryTable(sb, "年月", "住房率", monthlySummary)

        sb.AppendLine("<div class=""small-note"">☆ 用房/修參/鎖控/官網/空房/總數</div>")

        sb.AppendLine("<table>")
        sb.AppendLine("<tr>")
        sb.AppendLine("<th>日期</th><th>星期</th><th>入住</th><th>庫存</th><th>故障</th><th>住房率</th>")
        For Each roomType In roomTypes
            sb.AppendLine($"<th>{WebUtility.HtmlEncode(roomType)}</th>")
        Next
        sb.AppendLine("<th>金額</th><th>ADR</th>")
        sb.AppendLine("</tr>")

        For Each row In detailRows
            Dim inventoryTotal = SumIndex(row.RoomTypeData, settings.InventoryIndex)
            Dim faultTotal = SumIndex(row.RoomTypeData, settings.FaultIndex)
            Dim occValue = If(row.RoomOccupancyValue.HasValue, row.RoomOccupancyValue.Value, 0D)
            Dim occColor = OccupancyColor(occValue)

            sb.AppendLine("<tr>")
            sb.AppendLine($"<td>{row.DataDate:yyyy/MM/dd}</td>")
            sb.AppendLine($"<td>{WeekdayLabel(row.DataDate)}</td>")
            sb.AppendLine($"<td>{row.RoomCount}</td>")
            sb.AppendLine($"<td>{inventoryTotal}</td>")
            sb.AppendLine($"<td>{faultTotal}</td>")
            sb.AppendLine($"<td style=""background:{occColor};font-weight:700;"">{WebUtility.HtmlEncode(row.RoomOccupancyText)}</td>")

            For Each roomType In roomTypes
                Dim cellText = BuildRoomTypeCell(row.RoomTypeData, roomType)
                sb.AppendLine($"<td>{cellText}</td>")
            Next

            sb.AppendLine($"<td>{row.RoomRevenue.ToString("0", CultureInfo.InvariantCulture)}</td>")
            sb.AppendLine($"<td>{row.RoomAdr.ToString("0", CultureInfo.InvariantCulture)}</td>")
            sb.AppendLine("</tr>")
        Next

        sb.AppendLine("</table>")
        sb.AppendLine("</body></html>")
        Return sb.ToString()
    End Function

    Private Sub AppendSummaryTable(sb As StringBuilder, leftHeader As String, valueHeader As String, rows As List(Of (Label As String, Value As String, Color As String)))
        sb.AppendLine("<table>")
        sb.AppendLine("<tr>")
        sb.AppendLine($"<th class=""label"">{WebUtility.HtmlEncode(leftHeader)}</th>")
        For Each row In rows
            sb.AppendLine($"<th>{WebUtility.HtmlEncode(row.Label)}</th>")
        Next
        sb.AppendLine("</tr>")
        sb.AppendLine("<tr>")
        sb.AppendLine($"<th class=""label"">{WebUtility.HtmlEncode(valueHeader)}</th>")
        For Each row In rows
            sb.AppendLine($"<td style=""background:{row.Color};font-weight:700;"">{WebUtility.HtmlEncode(row.Value)}</td>")
        Next
        sb.AppendLine("</tr>")
        sb.AppendLine("</table>")
    End Sub

    Private Function BuildRoomTypeCell(roomTypeData As Dictionary(Of String, Integer()), roomType As String) As String
        Dim values As Integer() = Nothing
        If Not roomTypeData.TryGetValue(roomType, values) Then
            Return "0/0/0/0/0/0"
        End If

        Dim parts = Enumerable.Range(0, 6).Select(Function(idx) ReadIndexedValue(values, idx + 1).ToString(CultureInfo.InvariantCulture))
        Return String.Join("/", parts)
    End Function

    Private Function SumIndex(roomTypeData As Dictionary(Of String, Integer()), index As Integer) As Integer
        Return roomTypeData.Values.Sum(Function(values) ReadIndexedValue(values, index))
    End Function

    Private Function ReadIndexedValue(values As Integer(), oneBasedIndex As Integer) As Integer
        Dim idx = oneBasedIndex - 1
        If values Is Nothing OrElse idx < 0 OrElse idx >= values.Length Then
            Return 0
        End If
        Return values(idx)
    End Function

    Private Function WeekdayLabel(value As DateTime) As String
        Dim labels = New String() {"日", "一", "二", "三", "四", "五", "六"}
        Return labels(CInt(value.DayOfWeek))
    End Function

    Private Function OccupancyColor(value As Decimal) As String
        If value >= 80D Then
            Return "#1e9800"
        End If
        If value >= 60D Then
            Return "#f6f600"
        End If
        If value >= 30D Then
            Return "#ff2c2c"
        End If
        Return "#d0b0f4"
    End Function

    Private Function ParseRoomData(roomData As String) As Dictionary(Of String, Integer())
        Dim result As New Dictionary(Of String, Integer())(StringComparer.OrdinalIgnoreCase)
        If String.IsNullOrWhiteSpace(roomData) Then
            Return result
        End If

        For Each segment In roomData.Split(";"c, StringSplitOptions.RemoveEmptyEntries)
            Dim parts = segment.Split(":"c, 2)
            If parts.Length <> 2 Then
                Continue For
            End If

            Dim roomType = parts(0).Trim()
            Dim numbers = parts(1).Split("/"c).Select(Function(x) ParseInt(x)).ToArray()
            result(roomType) = numbers
        Next

        Return result
    End Function

    Private Sub SendEmail(settings As SmtpSettings, subject As String, html As String)
        If String.IsNullOrWhiteSpace(settings.Host) Then
            Throw New InvalidOperationException("Smtp.Host is required when Smtp.Enabled=true.")
        End If
        If settings.ToAddresses Is Nothing OrElse settings.ToAddresses.Count = 0 Then
            Throw New InvalidOperationException("At least one Smtp.ToAddresses entry is required when Smtp.Enabled=true.")
        End If

        Using message As New MailMessage()
            message.From = New MailAddress(settings.FromAddress, settings.FromDisplayName, Encoding.UTF8)
            For Each address In settings.ToAddresses.Where(Function(x) Not String.IsNullOrWhiteSpace(x))
                message.To.Add(address)
            Next
            For Each address In settings.CcAddresses.Where(Function(x) Not String.IsNullOrWhiteSpace(x))
                message.CC.Add(address)
            Next
            message.SubjectEncoding = Encoding.UTF8
            message.BodyEncoding = Encoding.UTF8
            message.Subject = subject
            message.Body = html
            message.IsBodyHtml = True

            Using client As New SmtpClient(settings.Host, settings.Port)
                client.EnableSsl = settings.EnableSsl
                If Not String.IsNullOrWhiteSpace(settings.Username) Then
                    client.Credentials = New NetworkCredential(settings.Username, settings.Password)
                End If
                client.Send(message)
            End Using
        End Using
    End Sub

    Private Function ParsePercent(text As String) As Decimal?
        If String.IsNullOrWhiteSpace(text) Then
            Return Nothing
        End If

        Dim clean = text.Replace("%", "").Trim()
        Dim value As Decimal
        If Decimal.TryParse(clean, NumberStyles.Any, CultureInfo.InvariantCulture, value) OrElse
           Decimal.TryParse(clean, NumberStyles.Any, CultureInfo.CurrentCulture, value) Then
            Return value
        End If

        Return Nothing
    End Function

    Private Function FormatPercentText(value As Decimal?) As String
        If Not value.HasValue Then
            Return "0%"
        End If

        Dim rounded = Math.Round(value.Value, 2)
        If rounded = Math.Truncate(rounded) Then
            Return $"{CInt(rounded)}%"
        End If

        Return $"{rounded.ToString("0.##", CultureInfo.InvariantCulture)}%"
    End Function

    Private Function ParseInt(value As Object) As Integer
        If value Is Nothing Then
            Return 0
        End If

        Dim text = value.ToString().Replace(",", "").Trim()
        Dim result As Integer
        If Integer.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, result) OrElse
           Integer.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, result) Then
            Return result
        End If

        Return 0
    End Function

    Private Function ParseDecimal(value As Object) As Decimal
        If value Is Nothing Then
            Return 0D
        End If

        Dim text = value.ToString().Replace(",", "").Trim()
        Dim result As Decimal
        If Decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, result) OrElse
           Decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, result) Then
            Return result
        End If

        Return 0D
    End Function

    Private Function ParseDateTimeNullable(text As String) As DateTime?
        If String.IsNullOrWhiteSpace(text) Then
            Return Nothing
        End If

        Dim formats = New String() {"yyyy/MM/dd HH:mm:ss", "yyyy-MM-dd HH:mm:ss", "yyyy/MM/dd", "yyyy-MM-dd"}
        Dim value As DateTime
        If DateTime.TryParseExact(text.Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, value) OrElse
           DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, value) Then
            Return value
        End If

        Return Nothing
    End Function
End Module
