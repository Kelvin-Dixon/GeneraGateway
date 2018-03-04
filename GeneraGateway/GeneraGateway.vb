Imports System.IO
Imports System.Timers
Imports System.Collections.ObjectModel
Imports System.Globalization
Public Class GeneraGateway

    Public gateway_dir As String = "Z:\proin\gateway\"
    Public opi_input_dir As String = "Z:\proin\hires\"
    Public file As String = ""
    Dim TimerCleanLocalDirs As New Timer
    Dim Timer1 As New Timer
    Dim advert As Boolean = True
    Dim pic As Boolean = False
    Public stabletime As Integer = 10
    Public PathToHiresPics As String = "z:\proin\hires\"


    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
    End Sub
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
        StartTimers()
    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
    End Sub
    Sub Gateway()
        Timer1.Stop()
        Try
            Makedir(gateway_dir)
            Makedir(opi_input_dir)

            Dim filetype As String = ""
            Dim count As Integer = 0
            Dim subject As String = ""
            Dim errorstring As String = ""
            Dim filename As String = ""
            Dim status As String = ""
            Dim bool_status As Boolean = False
            Try
                For Each File As String In Filesinafolder(gateway_dir, False, "*.*")
                    Dim testFile As FileInfo
                    testFile = My.Computer.FileSystem.GetFileInfo(File)
                    testFile.Delete()
                Next
            Catch ex As Exception
                Errorsub("gateway clean root " & ex.Message)
            End Try
            For Each File As String In Filesinafolder(gateway_dir, True, "*.*")
                Dim testFile As FileInfo
                testFile = My.Computer.FileSystem.GetFileInfo(File)
                If Not Fileinuse(testFile) Then
                    If InStr(UCase(testFile.DirectoryName), "\ERROR") = 0 Then
                        If InStr(UCase(testFile.DirectoryName), "\ADVERT") > 0 Then
                            filetype = ReadBin(testFile, gateway_dir)
                            If filetype = "EPS" Then
                                Call Move_file(testFile, gateway_dir, opi_input_dir, Get_ad_dir(testFile), filetype, advert)
                                EventLog.WriteEntry("processed Ad: " & testFile.FullName & vbCrLf & filetype, EventLogEntryType.Information)
                            Else
                                Call Errorfile(testFile, gateway_dir, filetype)
                                EventLog.WriteEntry("errored ad : " & testFile.FullName, EventLogEntryType.Warning)
                            End If
                        End If
                        If InStr(UCase(testFile.DirectoryName), "\OBJECT") > 0 Then
                            filetype = ReadBin(testFile, gateway_dir)
                            If filetype = "EPS" Then
                                Call Move_file(testFile, gateway_dir, opi_input_dir, Get_ad_dir(testFile), filetype, pic)
                                EventLog.WriteEntry("processed Object: " & testFile.FullName & vbCrLf & filetype, EventLogEntryType.Information)
                            Else
                                Call Errorfile(testFile, gateway_dir, filetype)
                                EventLog.WriteEntry("errored Object : " & testFile.FullName, EventLogEntryType.Warning)
                            End If
                        End If

                        If InStr(UCase(testFile.DirectoryName), "\PHOTO") > 0 Then
                            filetype = ReadBin(testFile, gateway_dir)
                            Dim subdir = Right(testFile.DirectoryName, Len(testFile.DirectoryName) - Len(gateway_dir))
                            Select Case filetype
                                Case "EPS"
                                    Rename_picture(testFile, PathToHiresPics & subdir & "\", gateway_dir)
                                    EventLog.WriteEntry("processed pic: " & testFile.FullName & vbCrLf & filetype, EventLogEntryType.Information)
                                Case "TIF"
                                    Rename_picture(testFile, PathToHiresPics & subdir & "\", gateway_dir)
                                    EventLog.WriteEntry("processed pic: " & testFile.FullName & vbCrLf & filetype, EventLogEntryType.Information)
                                Case Else
                                    Call Errorfile(testFile, gateway_dir, filetype)
                                    EventLog.WriteEntry("errored pic: " & testFile.FullName & vbCrLf & filetype, EventLogEntryType.Warning)
                            End Select
                        End If
                        If InStr(UCase(testFile.DirectoryName), "\BOOKINGPICTURE") > 0 Then
                            filetype = ReadBin(testFile, gateway_dir)
                            Select Case filetype
                                Case "EPS"
                                    Call Move_file(testFile, gateway_dir, opi_input_dir, Get_ad_dir(testFile), filetype, pic)
                                    EventLog.WriteEntry("processed pic: " & testFile.FullName & vbCrLf & filetype, EventLogEntryType.Information)
                                Case "TIF"
                                    Call Move_file(testFile, gateway_dir, opi_input_dir, Get_ad_dir(testFile), filetype, pic)
                                    EventLog.WriteEntry("processed pic: " & testFile.FullName & vbCrLf & filetype, EventLogEntryType.Information)
                                Case Else
                                    Call Errorfile(testFile, gateway_dir, filetype)
                                    EventLog.WriteEntry("errored pic: " & testFile.FullName & vbCrLf & filetype, EventLogEntryType.Warning)
                            End Select

                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            EventLog.WriteEntry("gateway  " & ex.Message)
        Finally
            Timer1.Start()
        End Try
    End Sub

    Public Function Get_ad_dir(ByVal testfile As FileInfo) As Integer
        Try
            Get_ad_dir = 0
            If IsNumeric(Mid(testfile.Name, 4, 1)) Then Get_ad_dir = Mid(testfile.Name, 4, 1)
        Catch ex As Exception
            Errorsub("get_ad_dir  " & ex.Message)
            Get_ad_dir = 0
        End Try
    End Function

    Public Function Filenameonly(ByVal testfile As FileInfo) As String
        Try
            Filenameonly = Left(testfile.Name, Len(testfile.Name) - Len(testfile.Extension))
        Catch ex As Exception
            Errorsub(" filenameonly  " & ex.Message)
            Filenameonly = ""
        End Try
    End Function

    Public Function Fileinuse(ByVal testfile As FileInfo) As Boolean
        Try
            Dim fs = New FileStream(testfile.FullName, FileMode.Open, FileAccess.Read)
            Dim bytes(256) As Byte
            fs.Read(bytes, 0, 20)
            fs.Close()
            fs = Nothing
            Fileinuse = False
        Catch ex As Exception
            Fileinuse = True
        End Try
    End Function

    Public Function ReadBin(ByVal testfile As FileInfo, ByVal pathfrom As String) As String
        Dim filetype As String = "NotRead"
        Try
            Dim variable1 As String = ""
            Dim isread As Boolean = False
            If testfile.Length > 50 Then
                Dim bytes(256) As Byte
                Dim fs = New FileStream(testfile.FullName, FileMode.Open, FileAccess.Read)
                fs.Read(bytes, 0, 20)
                variable1 = Text.Encoding.UTF7.GetString(bytes)
                fs.Close()
                fs = Nothing
                bytes = Nothing
                If variable1 <> "" Then isread = True
                If InStr(variable1, "EPSF") Then filetype = "EPS"
                If InStr(variable1, "II") Then filetype = "TIF"
                If InStr(variable1, "MM") Then filetype = "TIF"
                If InStr(variable1, "PDF") Then filetype = "PDF"
                If InStr(variable1, "GIF87a") Then filetype = "GIF"
                If InStr(variable1, "GIF89a") Then filetype = "GIF"
                If InStr(variable1, "ÿØÿí") Then filetype = "JPG"
                If InStr(variable1, "ÿØÿà") Then filetype = "JPG"
                If InStr(variable1, "BM") Then filetype = "BMP"
                If InStr(variable1, "ÅÐÓÆ") Then filetype = "EPS"
                If Not isread Then filetype = "NotRead"
                ReadBin = filetype
            Else
                isread = False
                ReadBin = "NotRead"
            End If
        Catch ex As Exception
            filetype = "NotRead"
            ReadBin = "NotRead"
            Errorsub("readbin  " & ex.Message)
        Finally
        End Try
    End Function

    Public Function Errorfile(ByVal testfile As FileInfo, ByVal pathfrom As String, ByVal filetype As String) As Boolean
        Try
            Dim now As DateTime = DateTime.Now
            Dim daynow As String = Day(now)
            Dim monthnow As String = Month(now)
            Dim yearnow As String = Year(now)
            If daynow < 10 Then daynow = "0" & daynow
            If monthnow < 10 Then monthnow = "0" & monthnow
            Dim datenow As String = daynow & monthnow & yearnow
            Dim hournow As String = Hour(now)
            Dim minutenow As String = Minute(now)
            Dim secondnow As String = Second(now)
            If hournow < 10 Then hournow = "0" & hournow
            If minutenow < 10 Then minutenow = "0" & minutenow
            If secondnow < 10 Then secondnow = "0" & secondnow
            Dim timenow As String = hournow & minutenow & secondnow
            Makedir(pathfrom & "\error")
            Dim testFile2 As FileInfo
            testFile2 = My.Computer.FileSystem.GetFileInfo(pathfrom & "\error" & "\" & testfile.Name)
            If testFile2.Exists Then
                testfile.MoveTo(pathfrom & "\error" & "\" & datenow & "-" & timenow & "-" & testfile.Name)
            Else
                testfile.MoveTo(pathfrom & "\error" & "\" & testfile.Name)
            End If
            If filetype = "NotRead" Then filetype = "Unknown"
            Errorfile = True
        Catch ex As Exception
            Errorfile = False
            Errorsub("errored " & ex.Message)
        End Try
    End Function

    Public Sub Move_file(ByVal testfile As FileInfo, ByVal pathfrom As String, ByVal moveto As String, ByVal subdir As Integer, ByVal filetype As String, ByVal advert As Boolean)
        Try
            If advert Then
                If subdir >= 0 Then
                    Try
                        Dim stabletime As Integer = 10
                        Dim temp1 As String = Right(testfile.DirectoryName, Len(testfile.DirectoryName) - Len(pathfrom))
                        EventLog.WriteEntry(" temp1: " & temp1)
                        If Stable_time(testfile.LastWriteTime, stabletime) Then
                            Makedir(moveto & temp1 & "\" & subdir.ToString)
                            If testfile.Exists Then
                                EventLog.WriteEntry(" exists: ")
                                Dim killFile As FileInfo
                                killFile = My.Computer.FileSystem.GetFileInfo(moveto & temp1 & "\" & subdir.ToString & "\" & Filenameonly(testfile) & "." & LCase(filetype))
                                killFile.Delete()
                                testfile.MoveTo(moveto & temp1 & "\" & subdir.ToString & "\" & Filenameonly(testfile) & "." & LCase(filetype))
                            Else
                                testfile.MoveTo(moveto & temp1 & "\" & subdir.ToString & "\" & Filenameonly(testfile) & "." & LCase(filetype))
                            End If
                        End If
                    Catch ex As Exception
                        Errorsub("move_file " & ex.Message)
                    End Try
                End If
            Else
                Try
                    Dim stabletime As Integer = 10
                    Dim temp1 As String = Right(testfile.DirectoryName, Len(testfile.DirectoryName) - Len(pathfrom))
                    If Stable_time(testfile.LastWriteTime, stabletime) Then
                        Makedir(moveto & "\" & temp1)
                        If testfile.Exists Then
                            Dim killFile As FileInfo
                            killFile = My.Computer.FileSystem.GetFileInfo(moveto & "\" & temp1 & "\" & Filenameonly(testfile) & "." & LCase(filetype))
                            killFile.Delete()
                            testfile.MoveTo(moveto & "\" & temp1 & "\" & Filenameonly(testfile) & "." & LCase(filetype))
                        Else
                            testfile.MoveTo(moveto & "\" & temp1 & "\" & Filenameonly(testfile) & "." & LCase(filetype))
                        End If
                    End If
                Catch ex As Exception
                    Errorsub("move_file " & ex.Message)
                End Try
            End If
        Catch ex As Exception
            Errorsub("move_file " & ex.Message)
        End Try

    End Sub
    Public Sub Clean_error_folder()
        Cleanlocalactivedirs(gateway_dir & "\error\")
    End Sub
#Region "tools"
    Public Function Pdate(ByVal s As String) As String
        Dim MyCultureInfo As CultureInfo = New CultureInfo("en-NZ")
        Dim result As DateTime
        Dim done As Boolean
        Try
            If Len(s) = 8 Then
                s = Mid(s, 1, 4) & "-" & Mid(s, 5, 2) & "-" & Mid(s, 7, 2)
            End If
            done = DateTime.TryParse(s, result)
            Dim dayfile As String = Day(result)
            Dim monthfile As String = Month(result)
            Dim yearfile As String = Year(result)
            If Len(yearfile) = 4 Then yearfile = Right(yearfile, 2)
            Pdate = dayfile & "/" & monthfile & "/" & yearfile
        Catch ex As Exception
            Errorsub("pdate: " & ex.ToString)
            Pdate = ""
        End Try
    End Function
    Public Function Compslit(ByVal InputText As String) As String()
        Dim delim As Char() = {"-", "_", "\", "/", "$", "."}
        Compslit = InputText.Split(delim)
    End Function

    Public Function Cleanlocalactivedirs(ByVal cleanpath As String) As Boolean
        Try
            Dim keepfiledays As Integer = 1
            For Each File As String In Filesinafolder(cleanpath, False, "*.*")
                Dim testFile As FileInfo
                testFile = My.Computer.FileSystem.GetFileInfo(File)
                If Stable_time(testFile.LastWriteTime, keepfiledays * 86400) Then
                    testFile.Delete()
                End If
            Next
        Catch ex As Exception
            Errorsub("cleanlocalactivedirs: " & ex.Message)
        End Try
    End Function
    Function Filesinafolder(ByVal path As String, ByVal subfolders As Boolean, ByVal type As String) As ReadOnlyCollection(Of String)
        Try
            If subfolders Then
                Filesinafolder = My.Computer.FileSystem.GetFiles(path, FileIO.SearchOption.SearchAllSubDirectories, type)
            Else
                Filesinafolder = My.Computer.FileSystem.GetFiles(path, FileIO.SearchOption.SearchTopLevelOnly, type)
            End If
        Catch ex As Exception
            Errorsub("filesinafolder: " & ex.ToString)
            Filesinafolder = Nothing
        End Try
    End Function

    Public Function Stable_time(ByVal t1string As String, ByVal stabletime As Int64) As Boolean
        Try
            Dim t1 As DateTime = Convert.ToDateTime(t1string)
            Dim now As DateTime = DateTime.Now
            t1 = t1.AddSeconds(stabletime)
            Stable_time = False
            If DateTime.Compare(now, t1) >= 0 Then
                Stable_time = True
            End If
        Catch ex As Exception
            Errorsub("stable_time: " & ex.Message)
        End Try

    End Function
    Public Function Makedir(ByVal dir As String) As Boolean
        Try
            Dim driveNames As New List(Of String)
            If My.Computer.FileSystem.DirectoryExists(dir) Then
                Return True
            Else
                My.Computer.FileSystem.CreateDirectory(dir)
                Return True
            End If
        Catch ex As Exception
            Return False
            Errorsub("make dir failed: " & dir & ": " & ex.Message)
        End Try
    End Function
    Sub Errorsub(ByVal errorstring As String)
        Try
            EventLog.WriteEntry("GeneraAdGateway " & errorstring, EventLogEntryType.Warning)
        Catch ex As Exception
        End Try
    End Sub


#End Region
#Region "Timers"
    Private Sub StartTimers()
        Try
            AddHandler Timer1.Elapsed, AddressOf Timer1_Tick
            Timer1.Interval = 10000
            Timer1.Enabled = True
            Timer1.Start()
        Catch ex As Exception
            Errorsub("timers: " & ex.ToString)
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal obj As Object, ByVal e As ElapsedEventArgs)
        Gateway()
    End Sub

#End Region


    Function Rename_picture(ByVal testfile As FileInfo, ByVal pathtohighrespics As String, ByVal gateway_dir As String) As Boolean
        EventLog.WriteEntry("pathtohighrespics: " & pathtohighrespics & vbNewLine & "testfile: " & testfile.FullName & vbNewLine & "gateway_dir: " & gateway_dir)
        Dim tempname As String = ""
        Dim found As Integer = 0
        Dim comslit_vars As String() = Nothing
        Dim Count As Integer = 1
        Dim ver_number As Integer = 1
        Dim test As Boolean = False
        Try
            If My.Computer.FileSystem.FileExists(pathtohighrespics & testfile.Name) Then
                Do Until test
                    test = False
                    If My.Computer.FileSystem.FileExists(pathtohighrespics & Filenameonly(testfile) & ".V" & ver_number.ToString & testfile.Extension) Then
                        ver_number = ver_number + 1
                        test = False
                        If ver_number > 1000 Then Exit Do
                    Else
                        test = True
                    End If
                Loop
                EventLog.WriteEntry("testfile.MoveTo: " & pathtohighrespics & Filenameonly(testfile) & ".V" & ver_number.ToString & testfile.Extension)
                testfile.MoveTo(pathtohighrespics & Filenameonly(testfile) & ".V" & ver_number.ToString & testfile.Extension)
            Else
                EventLog.WriteEntry("testfile.MoveTo: " & pathtohighrespics & testfile.Name)
                testfile.MoveTo(pathtohighrespics & testfile.Name)
            End If
            Rename_picture = True

        Catch ex As Exception
            Rename_picture = False
            Errorsub("rename_picture: " & ex.ToString)
        Finally
        End Try
    End Function
End Class
