Imports Osi.AlChavo.AT.DataProviders
Imports System.Reflection
'Imports System.Security.Cryptography
Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Globalization
Imports System.Threading

Public Class ATPunchClock
    Inherits System.Web.UI.Page

    'Global Enumeration to identify punch in/out
    Public Enum PunchOrder
        PunchOut = 0
        PunchIn = 1
    End Enum

    'Global enumeration to identify the punch type; not in use currently
    Public Enum PunchType
        Regular = 1
        Sick = 2
        Vacations = 3
    End Enum

    Public HoraAhora As DateTime
    Public APunchOrder As PunchOrder
    Public ApunchType As PunchType

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            SetSecurity()
            LoadCombo()
            InitGeoAndCamera()
            ppadtxt.Value = ""
        End If
        'lblTituloFecha.Text = DateTime.Now.ToString("D", System.Globalization.CultureInfo.CurrentCulture)
    End Sub

    ''' <summary>
    ''' Sets the page security
    ''' </summary>
    Public Sub SetSecurity()
        Osi.AlChavo.WebUI.ApplicationMethods.SECValidateViewPermissionRedirect(Osi.AlChavo.WebUI.ApplicationMethods.SECValidatePermission(SecurityManagement.DataProviders.Security.Actions.View))
    End Sub

    ''' <summary>
    ''' This method gets the type of punch to be saved from the page
    ''' </summary>
    ''' <returns>IF it is a punch IN (true) or not</returns>
    Protected Function MountPunch() As Boolean

        'Check values
        Dim PassPunchOrder As Boolean
        Dim PassPunchType As Boolean
        PassPunchOrder = Integer.TryParse(hidPunchOrder.Value, APunchOrder)
        'PassPunchType = Integer.TryParse(hidPunchType.Value, ApunchType)
        PassPunchType = True
        ApunchType = PunchType.Regular
        Return (PassPunchOrder And PassPunchType)


    End Function

    ''' <summary>
    ''' This method runs the general methods to the punch IN process, validate info, get punch, do punch
    ''' </summary>
    ''' <param name="sender">LnkBtnBottomPunchIN web object;Punch IN button</param>
    ''' <param name="e">LnkBtnBottomPunchIN on click argument</param>
    Protected Sub LnkBtnBottomPunchIN_Click(sender As Object, e As EventArgs) Handles LnkBtnBottomPunchIN.Click

        ResetErrs()

        If Not ValidateInputs() Then
            'Bad Inputs or Date
            Return
        End If

        If Not MountPunch() Then
            lblInvalidPunch.Visible = True
            Return
        End If

        If Not DoPunch(APunchOrder, ApunchType) Then
            'could not punch
            Return
        End If

    End Sub

    ''' <summary>
    ''' This method runs the general methods to the punch OUT process, validate info, get punch, do punch
    ''' </summary>
    ''' <param name="sender">LnkBtnBottomPunchOUT web object;Punch OUT button</param>
    ''' <param name="e">LnkBtnBottomPunchOUT on click argument</param>
    Protected Sub LnkBtnBottomPunchOUT_Click(sender As Object, e As EventArgs) Handles LnkBtnBottomPunchOUT.Click

        ResetErrs()

        If Not ValidateInputs() Then
            Return
        End If

        If Not MountPunch() Then
            lblInvalidPunch.Visible = True
            Return
        End If

        If Not DoPunch(APunchOrder, ApunchType) Then
            'could not punch
            Return
        End If

    End Sub

    ''' <summary>
    ''' Loads the department combobox 
    ''' Related Stored Proc: ATGetDpts
    ''' </summary>
    Protected Sub LoadCombo()

        Dim dp As New Osi.AlChavo.AT.DataProviders.AttendanceDataProvider
        Dim lst As List(Of ATGetDpts1_Result) = dp.GetCompanyDepartments(Session("CompanyID"))
        Dim dt As New DataTable
        Dim DefaultDept As New ATGetDpts1_Result()
        If Session("Language").ToString.Contains("es") Then
            DefaultDept.Description = "Departamento Defecto"
        Else
            DefaultDept.Description = "Default Department"
        End If
        DefaultDept.EntryNum = -1
        lst.Insert(0, DefaultDept)
        dt = ConvertToDataTable(lst)
        DdlDepartments.DataSource = dt
        DdlDepartments.DataTextField = "Description"
        DdlDepartments.DataValueField = "EntryNum"
        'DdlDepartments.Items.Insert(0, New Telerik.Web.UI.RadComboBoxItem("Default Department", -1))
        DdlDepartments.DataBind()

        dt.Dispose()
        DefaultDept = Nothing
        dp = Nothing

    End Sub

    ''' <summary>
    ''' This method does the punch process
    ''' Related Stored Procedures
    ''' Check Pin: Stored Proc: ATCheckPinByCompID
    ''' Get geolocation info
    ''' Get day of the week for the payroll start: ATCheckPayrollStart
    ''' Check timecard existance: ATGetTimeCardByDateSpan
    ''' Retrieves timecard: ATGetLastTimeCardByEmpEntryNum
    ''' Create a timecard if it does not exist for the week: ATCreateTimeCard
    ''' Check if the employee is set for autopunch and has done them: ATGetEmployeeAutoPunch
    ''' Punch: ATPunch
    ''' Notifications
    ''' </summary>
    ''' <param name="Punchorder">Punch IN or OUT</param>
    ''' <param name="Punchtype">Always regular punch</param>
    ''' <returns>True on punch success, false otherwise</returns>
    Protected Function DoPunch(Punchorder As PunchOrder, Punchtype As PunchType) As Boolean

        Dim AreValidated As Boolean = True
        'Check Pin
        Dim pinnumber As String = Osi.AlChavo.SecurityManagement.DataProviders.Security.EncriptAT(ppadtxt.Value)
        Dim cardDate As DateTime
        Dim dp As New AttendanceDataProvider
        Dim pinlist As List(Of ATCheckPinByCompID_Result) = dp.CheckPinByCompId(Session("CompanyID"), pinnumber)
        Dim PunchLatitude As Double? = 0, PunchLongitude As Double? = 0, IPAddress As String

        If pinlist.Count = 0 Then
            'Pin not found return
            lblWrongPIN.Visible = True
            Return False
        End If

        If Not checkGeo(pinlist) Then
            Return False
        End If

        'Check Previous punch type and department with this if Punch Out    Se va obviar estas validaciones al momento de ponchar por el proposito de expedicion
        'If Punchorder = ATPunchClock.PunchOrder.PunchOut Then

        '    Dim LastPunchList As List(Of ATGetLastPunchByEmpEntryNum_Result) = dp.GetLastPunchByEntryNumDate(Session("CompanyID"), pinlist(0).employeeentrynum, HoraAhora)

        '    If LastPunchList.Count > 0 Then

        '        If LastPunchList(0).isPunchIn Then

        '            If LastPunchList(0).Sort <> Punchtype Or LastPunchList(0).DepartmentID <> Integer.Parse(DdlDepartments.SelectedValue) Then
        '                'Warn different types or departments on PunchIn PunchOut combo
        '                lblDPunches.Visible = True
        '                Return False
        '            End If

        '        End If

        '    End If

        'End If

        'Get payroll start day
        Dim startday As Integer
        Dim CheckStartList As List(Of Integer?) = dp.CheckPayrollStart(Session("CompanyID"))

        If CheckStartList.Count = 0 Then
            'Start day config not found return
            lblCompanyPayrollStart.Visible = True
            Return False
        End If

        startday = CheckStartList(0)

        'Check timecard existance
        Dim startdate As New DateTime
        Dim enddate As DateTime = HoraAhora
        Dim NombreEmpleado As String = ""

        startdate = GetDateOfLast(startday - 1)


        Dim timecard As List(Of ATGetTimeCardByDateSpan_Result) = dp.GetCardByTimeSpan(Session("CompanyID"), startdate, enddate, pinlist(0).employeeentrynum)
        Dim CreatedCardList As List(Of ATGetLastTimeCardByEmpEntryNum_Result)

        If timecard.Count = 0 Then
            'No timecard; Create one
            Dim InsertResult As Integer = dp.CreateTimeCard(Session("CompanyID"), pinlist(0).employeeid, pinlist(0).employeeentrynum)

            If InsertResult > 0 Then

                CreatedCardList = dp.GetCardByEmpEntryNum(Session("CompanyID"), pinlist(0).employeeentrynum)

                If CreatedCardList.Count > 0 Then

                    cardDate = CreatedCardList(0).CardDate
                    NombreEmpleado = CreatedCardList(0).nombre

                Else

                    'Could not find card return
                    lblCouldNotPunch.Visible = True
                    Return False

                End If

            Else

                'Could not create card return
                lblCouldNotPunch.Visible = True
                Return False

            End If

        Else
            cardDate = timecard(0).carddate
            NombreEmpleado = timecard(0).nombre
        End If

        If Not (Double.TryParse(hdLat.Value, PunchLatitude) And Double.TryParse(hdLong.Value, PunchLongitude)) Then
            PunchLatitude = 0
            PunchLongitude = 0
        End If

        IPAddress = Request.UserHostAddress

        'Autopunch feature May 2015 
        Dim EmployeeAutoPunchInfo As List(Of ATGetEmployeeAutoPunch_Result) = dp.ATGetEmployeeAutoPunch(Session("CompanyID"), pinlist(0).employeeid, pinlist(0).employeeentrynum, HoraAhora)
        Dim IsAutoPunch As Boolean = EmployeeAutoPunchInfo(0).IsAutoPunch
        Dim CanPunch As Boolean = EmployeeAutoPunchInfo(0).CanPunch

        If Session("UserGuidID") Is Nothing Or Session("CompanyID") Is Nothing Then
            System.Web.Security.FormsAuthentication.SignOut()
            System.Web.Security.FormsAuthentication.RedirectToLoginPage()
        End If

        If Not IsAutoPunch Then
            'Normal Punch
            Dim PunchResult As Integer = dp.Punch(Session("CompanyID"), pinlist(0).employeeid, pinlist(0).employeeentrynum, HoraAhora,
                         cardDate, DdlDepartments.SelectedValue, "Web", Punchtype, Punchorder, Session("UserGuidID"), hdPhoto.Value _
                         , PunchLatitude, PunchLongitude, IPAddress)

            If PunchResult = 0 Then

                lblCouldNotPunch.Visible = True
                Return False

            End If

        Else

            If CanPunch Then
                'Run autopunch
                dp.ATAutoPunch(Session("CompanyID"), pinlist(0).employeeid, pinlist(0).employeeentrynum, HoraAhora,
                         cardDate, DdlDepartments.SelectedValue, "Web", Punchtype, New Guid(Session("UserGuidID").ToString), hdPhoto.Value _
                         , PunchLatitude, PunchLongitude, IPAddress)
            Else
                'Autopunch done already today
                lblAutopunchDone.Visible = True
            End If

        End If

        If DdlDepartments.Items.Count > 0 Then
            DdlDepartments.SelectedIndex = 0
        End If

        lblPunchName.Text = ""
        lblPunchName.Text = NombreEmpleado '& "<br/>"
        lblPunchName.Visible = True
        'lblPunchComplete.Visible = True

        'SavePhoto(hdPhoto.Value)
        ppadtxt.Value = ""

        'Notifications August 2015
        Dim NotificationTask = Osi.AlChavo.WebUI.ATClock.CheckNotifications(hdArrivalNot.Value, hdLateNot.Value, DdlDepartments.SelectedValue, pinlist(0).employeeentrynum,
               Session("CompanyID"), HoraAhora, NombreEmpleado, New Guid(Session("UserGuidID").ToString()), hdAbsentNOT.Value, hdMissingNOT.Value, hdEarlyArrivalNOT.Value, pinlist(0).employeeid)

        dp = Nothing
        Return AreValidated
    End Function

    ''' <summary>
    ''' Punch time validation
    ''' </summary>
    ''' <returns>True if date is valid, false otherwise.</returns>
    Protected Function ValidateInputs() As Boolean

        Dim AreValidated As Boolean = True
        'Validate Date
        If Not DateTime.TryParse(hidTime.Value, HoraAhora) Then
            'Invalid Date
            lblBadDate.Visible = True
            AreValidated = False
        End If

        

        Return AreValidated
    End Function

    ''' <summary>
    ''' Retrieves geolocation settings: ATGetCompanyAppDefs
    ''' Check geo global settings.
    ''' Retrieves individual filter settings: ATGetMobileLocations
    ''' Checks individual settings
    ''' </summary>
    ''' <param name="pinlist">Object holds table system information from the person registering the punch.</param>
    ''' <remarks>The geolocation GLOBAL settings , if enabled must be met to pass validation unless
    ''' the individual has personalized settings from the filter page which he must pass one of them to be
    ''' able to punch. You can disable the global settings which applies to everyone and use the personalized
    ''' settings for case scenarios.</remarks>
    ''' <returns>True if  it passes the geolocation validation settings: <example>punch within distance</example>.</returns>
    Public Function checkGeo(ByRef pinlist As List(Of ATCheckPinByCompID_Result)) As Boolean
        Dim dp As New AttendanceDataProvider
        Dim GeoSetting As String = "0"
        Dim CompanyLat As Double, CompanyLong As Double
        Dim AppSetupList As List(Of ATGetCompanyAppDefs_Result) = dp.ATGetAppSetupDefs(Session("CompanyID"))
        Dim latitude2 As Double, longitude2 As Double, MinimunDistance As Double
        Dim DistanceMeasure As String = ""
        Dim EnforcedCount As Integer = 0, WithinBounds As Boolean = False
        Dim PassedLocationCheck As Boolean = True

        If AppSetupList.Count > 0 Then
            For Each item As ATGetCompanyAppDefs_Result In AppSetupList
                If item.NAME = "Geo" Then
                    GeoSetting = item.ValueAlpha
                ElseIf item.NAME = "CompanyLat" Then
                    CompanyLat = item.ValueNumeric
                ElseIf item.NAME = "CompanyLong" Then
                    CompanyLong = item.ValueNumeric
                ElseIf item.NAME = "GPSDist" Then
                    MinimunDistance = item.ValueNumeric
                    DistanceMeasure = item.ValueAlpha
                End If
            Next
        End If


        AppSetupList = Nothing

        'check lat y long 
        If Not (Double.TryParse(hdLat.Value, latitude2) And Double.TryParse(hdLong.Value, longitude2)) Then
            If hdGeoEnable.Value = "1" Then
                lblNoGeoLocationAvailable.Visible = True
            End If
            latitude2 = 0
            longitude2 = 0
        End If

        'Global Settings 
        If GeoSetting = "0" Then
            'Return True
        Else
            'Check Global Settings
            
            If Distance(CompanyLat, latitude2, CompanyLong, longitude2, DistanceMeasure) > MinimunDistance Then
                lblMinimunDist.Visible = True
                PassedLocationCheck = False
            Else
                Return True
            End If
        End If

        'Get Locations for this person loop/if enforced check
        Dim PunchLocations As List(Of ATGetMobileLocations_Result) = dp.GetMobileLocations(Session("CompanyID"), pinlist(0).employeeentrynum, pinlist(0).employeeid)

        If PunchLocations.Count > 0 Then
           

            For A As Integer = 0 To PunchLocations.Count - 1 Step 1

                If PunchLocations(A).Enforced Then
                    EnforcedCount += 1
                End If

                If PunchLocations(A).Device = "Mobile" Then
                    
                    If (Distance(PunchLocations(A).Latitude, latitude2, PunchLocations(A).Longitude, longitude2, DistanceMeasure) > MinimunDistance) Then
                        'lblMinimunDist.Visible = True
                    Else
                        WithinBounds = True
                        Exit For
                    End If

                ElseIf PunchLocations(A).Device = "IP" Then

                    If SameIPSubNet(PunchLocations(A).IP, Request.UserHostAddress) Then
                        WithinBounds = True
                        Exit For
                    End If

                End If
            Next

            If (EnforcedCount > 0 And Not WithinBounds) Then
                lblMinimunDist.Visible = True
                Return False
            Else
                If WithinBounds Then
                    lblMinimunDist.Visible = False
                    Return True
                End If
                If EnforcedCount = 0 Then
                    Return PassedLocationCheck
                End If
            End If

        End If

        dp = Nothing

        If PassedLocationCheck Then
            lblMinimunDist.Visible = False
        End If

        Return PassedLocationCheck

    End Function

    ''' <summary>
    ''' Compares the given IP parameters for matching. The matching can be exact <example> ( SetIP = 192.168.0.1 , UserIp = 192.168.0.1 ) </example>
    ''' Or it can match if UserIP falls withing a given SetIP subnet <example> ( SetIP = 192.168.0.* , UserIP = 192.168.0.77 ). </example>
    ''' </summary>
    ''' <param name="SetIP" >IP set to match from table definition (tblatpunchlocations.IP).</param>
    ''' <param name="UserIP" >IP to compare with the definition.</param>
    ''' <returns>True if IP is valid, false otherwise.</returns>
    Public Function SameIPSubNet(ByVal SetIP As String, ByVal UserIP As String) As Boolean
        Dim AllowedIpParts As String() = SetIP.Split(".")
        Dim ipParts As String() = UserIP.Split(".")
        If (AllowedIpParts.Length > 0 And ipParts.Length > 0) And (AllowedIpParts.Length = ipParts.Length) Then
            For i As Integer = 0 To AllowedIpParts.Length - 1 Step 1
                ' Compare if not wildcard
                If (AllowedIpParts(i) <> "*") Then
                    ' Compare IP address part
                    If (ipParts(i) <> AllowedIpParts(i).ToLower()) Then
                        Return False
                    End If
                End If
            Next
        Else
            Return False
        End If
        Return True

    End Function

    ''' <summary>
    ''' Initializes the global app settings for the clock.
    ''' Get application setup list: ATGetCompanyAppDefs
    ''' </summary>
    Public Sub InitGeoAndCamera()
        Dim dp As New AttendanceDataProvider
        Dim AppSetupList As List(Of ATGetCompanyAppDefs_Result) = dp.ATGetAppSetupDefs(Session("CompanyID"))

        If AppSetupList.Count > 0 Then
            For Each item As ATGetCompanyAppDefs_Result In AppSetupList
                If item.NAME = "Geo" Then
                    If item.ValueAlpha = "1" Then
                        hdGeoEnable.Value = 1
                    Else
                        hdGeoEnable.Value = 0
                    End If
                ElseIf item.NAME = "Camera" Then
                    If item.ValueAlpha = "1" Then
                        hdCamEnable.Value = 1
                    Else
                        hdCamEnable.Value = 0
                    End If
                ElseIf item.NAME = "ArrivalNotifications" Then
                    If item.ValueAlpha = "1" Then
                        hdArrivalNot.Value = 1
                    Else
                        hdArrivalNot.Value = 0
                    End If
                ElseIf item.NAME = "LateNotification" Then
                    If item.ValueAlpha = "1" Then
                        hdLateNot.Value = 1
                    Else
                        hdLateNot.Value = 0
                    End If
                ElseIf item.NAME = "LateNotificationMins" Then
                    If Not item.ValueAlpha Is Nothing Then
                        hdLateMins.Value = item.ValueAlpha
                    ElseIf Not item.ValueNumeric Is Nothing Then
                        hdLateMins.Value = item.ValueNumeric.ToString
                    Else
                        hdLateMins.Value = "8"
                    End If
                ElseIf item.NAME = "MissingPunchs" Then
                    If Not item.ValueAlpha Is Nothing Then
                        hdMissingNOT.Value = item.ValueAlpha
                    Else
                        hdMissingNOT.Value = 0
                    End If
                ElseIf item.NAME = "Absences" Then
                    If Not item.ValueAlpha Is Nothing Then
                        hdAbsentNOT.Value = item.ValueAlpha
                    Else
                        hdAbsentNOT.Value = 0
                    End If
                ElseIf item.NAME = "EarlyArrival" Then
                    If Not item.ValueAlpha Is Nothing Then
                        hdEarlyArrivalNOT.Value = item.ValueAlpha
                    Else
                        hdEarlyArrivalNOT.Value = 0
                    End If
                End If
            Next
        End If

        dp = Nothing
        AppSetupList = Nothing

    End Sub

    ''' <summary>
    ''' Calculates the distance between 2 given latitudes and longitudes
    ''' </summary>
    ''' <param name="latitude1" >First latitude in decimal format (degrees: Ex: 18.0 )</param>
    ''' <param name="latitude2">Second latitude in decimal format</param>
    ''' <param name="longitude1" >First longitude in decimal format (degree Ex: 66.0 )</param>
    ''' <param name="longitude2" >Second longitude in decimal format.</param>
    ''' <returns>Distance between the points in the scale setup in tblatapplicationsetupdefs.name valuealpha .</returns>
    Public Function Distance(ByVal latitude1 As Double, ByVal latitude2 As Double, ByVal longitude1 As Double, ByVal longitude2 As Double, ByVal Measure As String) As Double
        'Dim a_distance As Double = 0
        'var R = 6371; // En Kilometros.
        Dim R As Double = 2.093 * Math.Pow(10, 7) 'En pies.
        'var R = 3959.191; // En millas.
        If Measure = "Feet" Then
            R = 2.093 * Math.Pow(10, 7) 'En pies.
        Else
            R = 3959.191 '; // En millas.
        End If
        Dim dLat As Double = (latitude2 - latitude1) * Math.PI / 180
        Dim dLon As Double = (longitude2 - longitude1) * Math.PI / 180
        Dim a As Double = Math.Sin(dLat / 2) * Math.Sin(dLat / 2) +
        Math.Cos(latitude1 * Math.PI / 180) * Math.Cos(latitude2 * Math.PI / 180) *
        Math.Sin(dLon / 2) * Math.Sin(dLon / 2)
        Dim c As Double = 2 * Math.Atan2(Math.Sqrt(a), Math.Sqrt(1 - a))
        Dim d As Double = R * c

        'If d > 1 Then
        '    a_distance = d
        'ElseIf d <= 1 Then
        '    a_distance = d * 1000
        'End If

        Return d
    End Function

    ''' <summary >
    ''' Converts the given list template to a strong typed datatable.
    ''' </summary>
    ''' <param name="list">A List template to convert.</param>
    ''' <returns>The List template in datatable format.</returns>
    Public Function ConvertToDataTable(Of T)(ByVal list As IList(Of T)) As DataTable

        Dim table As New DataTable()
        Dim fields() As PropertyInfo = GetType(T).GetProperties()

        For Each field As PropertyInfo In fields
            table.Columns.Add(field.Name, field.PropertyType)
        Next

        For Each item As T In list

            Dim row As DataRow = table.NewRow()
            For Each field As PropertyInfo In fields

                row(field.Name) = field.GetValue(item)

            Next
            table.Rows.Add(row)

        Next

        Return table
    End Function

    ''' <summary >
    ''' Calculates the last time as date when the given day of the week as integer happened.
    ''' Ex: dia = 0 (Sunday) if the punch was on date 2015-09-03 calculates to 2015-08-30 (previous Sunday).
    ''' </summary>
    ''' <param name="dia">Number representing the day of the week to look for starting from 0 (Sunday) to 6 (Saturday).</param>
    ''' <returns>The previous date when that day of the week happened.</returns>
    Public Function GetDateOfLast(ByVal dia As Integer) As DateTime

        Dim counter As Double = 0

        While DateAdd(DateInterval.Day, counter, DateTime.Now).DayOfWeek <> dia
            counter = counter - 1
        End While

        Return DateAdd(DateInterval.Day, counter, DateTime.Now)

    End Function

    ''' <summary >
    ''' Calculates the last time as date when the given day of the week as integer happened.
    ''' Ex: dia = 0 (Sunday) and offesetdate = 2015-09-03 calculates to 2015-08-30 (previous Sunday).
    ''' </summary>
    ''' <param name="dia">Number representing the day of the week to look for starting from 0 (Sunday) to 6 (Saturday).</param>
    ''' <param name="offesetdate" >Date from which to start counting backwards.</param>
    ''' <returns>The previous date when that day of the week happened starting from offesetdate.</returns>
    Public Function GetDateOfLastStartingFromDate(ByVal dia As Integer, ByVal offesetdate As DateTime) As DateTime

        Dim counter As Double = 0

        While DateAdd(DateInterval.Day, counter, offesetdate).DayOfWeek <> dia
            counter = counter - 1
        End While

        Return DateAdd(DateInterval.Day, counter, offesetdate)

    End Function

    ''' <summary >
    ''' Resets the validation labels to hidden.
    ''' </summary>
    Public Sub ResetErrs()

        lblBadDate.Visible = False
        lblInvalidPunch.Visible = False
        lblWrongPIN.Visible = False
        lblCompanyPayrollStart.Visible = False
        lblCouldNotPunch.Visible = False
        lblPunchName.Visible = False
        lblPunchComplete.Visible = False
        lblDPunches.Visible = False
        lblMinimunDist.Visible = False
        lblAutopunchDone.Visible = False

    End Sub

    ''' <summary >
    ''' Takes a base64 string representing a photo and saves it to disk.
    ''' </summary>
    ''' <param name="Photo">A base64 representing a photo.</param>
    ''' <returns>1.</returns>
    ''' <remarks>Not in use; photos are being saved as string in tblatPunchPhotos with relation to its punch (tblatpunches).</remarks>
    Public Function SavePhoto(ByVal Photo As String)
        ' Convert Base64 String to byte[]
        Try

            Dim imageBytes As Byte() = Convert.FromBase64String(Photo)
            Dim ms As MemoryStream = New MemoryStream(imageBytes, 0, imageBytes.Length, True)

            ' Convert byte[] to Image
            ms.Write(imageBytes, 0, imageBytes.Length)
            Dim bmp As New Bitmap(ms)
            bmp.Save("C:\temp\Images\" & Session("CompanyID") & "\test.jpeg", ImageFormat.Jpeg)

        Catch exception As Exception
            Console.WriteLine(exception.InnerException)
        End Try

        Return 1

    End Function

    Protected Overrides Sub InitializeCulture()
        Dim Culture As String = Convert.ToString(Session("Language"))
        Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(Culture)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(Culture)
        MyBase.InitializeCulture()
    End Sub
End Class