' ==========================================================+
' ===================== Main module ========================+
' ==========================================================+

Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Kelvin Kellner
' Student ID: 190668940
' Date: 07-30-2021
' Program title: CP212 A5 Fictional Speedrunning Community
' Module description: used to launch the Main form dashboard
'                     and start the program
'===========================================================+

Sub Start()
    ' Display the appropriate Form depending on whether or not the User has Logged In yet
    If frmMain.userPlayerID = 0 Then
        ' Prompt them to Log In or Sign Up
        frmNotLoggedIn.Show
    Else
        ' Show Main program dashboard
        frmMain.Show
    End If
End Sub

' Utility Function: Converts time as single to formatted time string
Public Function TimeToString(time As Single) As String
        TimeToString = Format(time - (time Mod 3600), "00") & ":" & Format(((time Mod 3600) - (time Mod 60)) / 60, "00") & ":" & Format(time Mod 60, "00") & ":" & Format((time * 1000) Mod 1000, "00")
End Function

' Utility Function: Converts date as string to formatted date string
Public Function DateToString(datestr As String) As String
        DateToString = Format(CDate(datestr), "short date")
End Function

' Utility Function: Saves me time, prints array to VBA Worksheet
' Source: https://newtonexcelbach.com/2012/03/12/writing-arrays-to-the-worksheet-vba-function/
Function CopyToRange(VBAArray As Variant, RangeName As String, Optional NumRows As Long = 0, _
                    Optional NumCols As Long = 0, Optional ClearRange As Boolean = True, _
                    Optional RowOff As Long = 0, Optional ColOff As Long = 0) As Long
    Dim DataRange As Range
 
    On Error GoTo RtnError
    If TypeName(VBAArray) = "Range" Then VBAArray = VBAArray.Value2
    If NumRows = 0 Then NumRows = UBound(VBAArray)
    If NumCols = 0 Then NumCols = UBound(VBAArray, 2)
    Set DataRange = Range(RangeName)
    If ClearRange = True Then DataRange.Offset(RowOff, ColOff).ClearContents
    DataRange.Resize(NumRows, NumCols).name = RangeName
    Range(RangeName).Offset(RowOff, ColOff).Value = VBAArray
    Set DataRange = Nothing
    CopyToRange = 0
    Exit Function
     
RtnError:
    CopyToRange = 1
End Function



' ==========================================================+
' ======================= frmAddGame =======================+
' ==========================================================+

Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Kelvin Kellner
' Student ID: 190668940
' Date: 07-30-2021
' Program title: CP212 A5 Fictional Speedrunning Community
' Form description: Allows for creation of new Games
'===========================================================+

Private Sub btnAddGame_Click()
    ' Check for empty text fields
    If txtTitle.Value = Empty Or txtStudio.Value = Empty Or txtYear.Value = Empty Then
        MsgBox "Please fill out all fields.", vbCritical, "Empty Fields"
    Else
        With frmMain.rs
            ' Select any matching entries from rs
            .Open "SELECT * FROM Games WHERE Title = '" & txtTitle.Value & "' AND Studio = '" & txtStudio.Value & "' AND ReleaseYear = " & txtYear.Value, frmMain.cn, adOpenDynamic, adLockOptimistic
            ' If duplicate Game
            If Not .EOF Then
                MsgBox "That game is already in our database.", vbCritical, "Duplicate Game"
            Else
                ' Add fields from textboxes
                .AddNew
                frmMain.rs("Title").Value = txtTitle.Value
                frmMain.rs("Studio").Value = txtStudio.Value
                frmMain.rs("ReleaseYear").Value = txtYear.Value
                .Update ' update rs
            End If
            .Close ' close rs
            ' Unload Forms
            Unload Me
            Unload frmSelectGame
            ' Re-show Select Game form, initializes again
            frmSelectGame.Show
        End With
    End If
End Sub



' ==========================================================+
' ====================== frmEditGame =======================+
' ==========================================================+

Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Kelvin Kellner
' Student ID: 190668940
' Date: 07-30-2021
' Program title: CP212 A5 Fictional Speedrunning Community
' Form description: Allows for editing existing Games
'===========================================================+

Private Sub UserForm_Initialize()
    ' Fill game data
    With frmMain.rs
        .Open "SELECT * FROM Games WHERE GameID = " & frmMain.currentGameID, frmMain.cn
        If Not .EOF Then
            txtTitle.Text = .Fields("Title")
            txtStudio.Text = .Fields("Studio")
            txtYear.Text = .Fields("ReleaseYear")
        End If
        .Close ' close rs
    End With
End Sub

Private Sub btnSave_Click()
    If txtTitle.Value = Empty Or txtStudio.Value = Empty Or txtYear.Value = Empty Then
        MsgBox "Please fill out all fields.", vbCritical, "Empty Fields"
    Else
        ' Reflect changes in Games database
        With frmMain.rs
            ' Select Game from Games database
            .Open "SELECT * FROM Games WHERE GameID = " & frmMain.currentGameID, frmMain.cn, adOpenKeyset, adLockOptimistic
            ' If Game Found
            If Not .EOF Then
                ' Fill fields with Game info from textboxes
                .Fields("Title").Value = txtTitle.Text
                .Fields("Studio").Value = txtStudio.Text
                .Fields("ReleaseYear").Value = txtYear.Text
                .Update ' update rs
            Else
                ' Print message if not found (should not happen unless there is an error)
                MsgBox "Game could not be found in database.", vbCritical, "Game Not Found"
            End If
            .Close ' close rs
        End With
        Unload Me
        Unload frmSelectGame
        frmSelectGame.Show
    End If
End Sub

Private Sub lblDeleteGame_Click()
    ' Yes or No popup to prevent misclicks!
    If MsgBox("This action cannot be undone! Are you sure?", vbYesNo) = vbNo Then Exit Sub
    
    ' Delete Game from rs
    With frmMain.rs
        ' Select Game by Game ID
        .Open "SELECT * FROM Games WHERE GameID = " & frmMain.currentGameID, frmMain.cn, adOpenKeyset, adLockOptimistic
        If Not .EOF Then
            ' Delete the record from rs
            .Delete
            .UpdateBatch
        End If
        .Close ' close rs
    End With
    ' Reset current Game ID
    frmMain.currentGameID
    ' Unload Forms
    Unload Me
    Unload frmSelectGame
    ' Restart Select Games Form
    frmSelectGame.Show
End Sub



' ==========================================================+
' ======================== frmLogIn ========================+
' ==========================================================+

Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Kelvin Kellner
' Student ID: 190668940
' Date: 07-30-2021
' Program title: CP212 A5 Fictional Speedrunning Community
' Form description: Allows for Logging into existing accounts
'===========================================================+

' On Submit
Private Sub btnLogIn_Click()
    Dim msg As String, exists As Boolean
    ' Retrieve text from textbox
    msg = txtUsername.Text
    
    ' Check for empty textbox
    If msg = Empty Then
        ' Error when User does not enter a username
        MsgBox "Please enter a username.", vbCritical, "Blank Username"
    Else
        exists = False ' default to False
        ' Search through all Players and see if User account exists
        With frmMain.rs
            .Open "SELECT * FROM Players", frmMain.cn
            ' Loop until we find a match or reach end of rs query
            Do Until .EOF Or exists = True
                ' If username in rs matches textbox text
                If .Fields("Username") = msg Then
                    ' Set the userPlayerID variable
                    frmMain.userPlayerID = .Fields("PlayerID")
                    frmMain.currentPlayerID = .Fields("PlayerID")
                    exists = True
                End If
                ' Move to next Player in rs
                .MoveNext
            Loop
            .Close ' close rs
        End With
        
        ' Perform appropriate action for success and failure
        If exists = False Then
            ' Account does not exist
            MsgBox "That username could not be found in our database.", vbCritical, "Account Does Not Exist"
        Else
            ' Close the Log In Forms and open Main Form dashboard
            Unload Me
            Unload frmNotLoggedIn
            frmMain.Show
        End If
    End If
End Sub



' ==========================================================+
' ======================== frmMain =========================+
' ==========================================================+

Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Kelvin Kellner
' Student ID: 190668940
' Date: 07-30-2021
' Program title: CP212 A5 Fictional Speedrunning Community
' Form description: Main dashboard for the program
'===========================================================+

Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset

Public userPlayerID As Integer
Public currentPlayerID As Integer
Public currentGameID As Integer
Public nextForm As String

Const dbName = "speedrunning.accdb"

' On Form Close, disconnect from database
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Close connection to database
    cn.Close
End Sub

Private Sub btnPlayers_Click()
    ' Open form
    frmPlayers.Show
End Sub

Private Sub btnMyProfile_Click()
    ' Open form
    frmMyProfile.Show
End Sub

Private Sub btnLogOut_Click()
    ' Clear currently Logged In Player
    userPlayerID = 0
    ' Close dashboard
    Unload Me
    ' Restart program
    Main.Start
End Sub

Private Sub btnNewAttempt_Click()
    ' If a Game has not been selected so far then open Select Game Form first
    If currentGameID = 0 Then
        ' Set nextForm to indicate with form to open on successful Continue in Select Game Form
        ' set nextForm String so Select Game Form knows where to continue to
        nextForm = "frmNewAttempt"
        frmSelectGame.Show
    ' Otherwise simply open the New Speedrun Attempt Form
    Else
        frmNewAttempt.Show
    End If
End Sub

Private Sub btnRuns_Click()
    ' If a Game has not been selected so far then open Select Game Form first
    If currentGameID = 0 Then
        ' Set nextForm to indicate with form to open on successful Continue in Select Game Form
        nextForm = "frmRuns"
        frmSelectGame.Show
    Else
        frmRuns.Show
    End If
End Sub

' On Form Intialize, establish database connection
Private Sub UserForm_Initialize()
    ' Open connection to database
    With cn
        .ConnectionString = "Data Source=" & ThisWorkbook.Path & "\" & dbName
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    
    userPlayerID = 0
    currentPlayerID = 0
    currentGameID = 0
End Sub



' ==========================================================+
' ====================== frmMyProfile ======================+
' ==========================================================+

Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Kelvin Kellner
' Student ID: 190668940
' Date: 07-30-2021
' Program title: CP212 A5 Fictional Speedrunning Community
' Form description: Profile page for the User, allows them
'                   to view and edit their account details
'===========================================================+

Dim editing As Boolean, changed As Boolean

Private Sub btnEdit_Click()
    If editing Then
        ' Lock fields and hide delete option
        editing = False
        lblDeleteAccount.Visible = False
        txtLocation.Locked = True
        txtBio.Locked = True
        txtUsername.Enabled = True
        fraMyProfile.Height = 234
        btnEdit.Caption = "Edit"
        ' Reflect changes in Player database
        With frmMain.rs
            ' Select user from Player database
            .Open "SELECT * FROM Players WHERE PlayerID = " & frmMain.userPlayerID, frmMain.cn, adOpenKeyset, adLockOptimistic
            changed = False
            ' If Player Found
            If Not .EOF Then
                If .Fields("Username").Value <> txtUsername.Text Or .Fields("Location").Value <> txtLocation.Text Or .Fields("Bio").Value <> txtBio.Text Then
                    ' Update information in rs using textboxes
                    .Fields("Username").Value = txtUsername.Text
                    .Fields("Location").Value = txtLocation.Text
                    .Fields("Bio").Value = txtBio.Text
                    .Update
                    changed = True
                End If
            Else
                ' Print message if not found (should not happen unless there is an error)
                MsgBox "Player could not be found in database.", vbCritical, "Player Not Found"
            End If
            .Close ' close rs
            ' Don't bother reloading if no changes are made
            If changed Then
                ' Unload Forms
                Unload Me
                Unload frmPlayers
                ' Show Players Form
                frmPlayers.Show
            End If
        End With
    Else
        ' Unlock fields and make options visible
        editing = True
        lblDeleteAccount.Visible = True
        txtLocation.Locked = False
        txtBio.Locked = False
        txtUsername.Enabled = False
        fraMyProfile.Height = 258
        btnEdit.Caption = "Save Changes"
    End If
End Sub

Private Sub lblDeleteAccount_Click()
    ' Yes or No popup to prevent misclicks!
    If MsgBox("This action cannot be undone! Are you sure?", vbYesNo) = vbNo Then Exit Sub
    
    ' Delete account from rs
    With frmMain.rs
        ' Select Player by Player ID
        .Open "SELECT * FROM Players WHERE PlayerID = " & frmMain.userPlayerID, frmMain.cn, adOpenKeyset, adLockOptimistic
        If Not .EOF Then
            ' Delete the record from rs
            .Delete
            .UpdateBatch
        End If
        .Close ' close rs
    End With
    
    ' Unload Forms
    Unload Me
    Unload frmPlayers
    Unload frmMain
    ' Start program over
    Main.Start
End Sub

Private Sub btnBack_Click()
    ' Unload Form
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    editing = False ' default value
    ' Call Sub to grab user Player data from rs
    Call frmPlayers.FillTextboxesFromPlayerID(frmMain.userPlayerID, txtUsername, txtLocation, txtBio)
    ' Call Sub to fill Runs using rs
    Call frmPlayerProfile.FillPlayerRunsFromPlayerID(frmMain.userPlayerID, lstHeaders, lstRuns)
End Sub



' ==========================================================+
' ====================== frmNewAttempt =====================+
' ==========================================================+

Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Kelvin Kellner
' Student ID: 190668940
' Date: 07-30-2021
' Program title: CP212 A5 Fictional Speedrunning Community
' Form description: Timer for attempting new speedrunning
'                   Runs for any particular Game
'===========================================================+

Dim running As Boolean
Dim startTime As Single, currentTime As Single, elapsed As Single
Dim h As Integer, m As Integer, s As Integer, ms As Integer

Private Sub btnSelectGame_Click()
    ' Show Form
    frmMain.nextForm = "frmNewAttempt"
    Unload Me ' so it will re-initialize after it is finished
    frmSelectGame.Show
End Sub

Private Sub btnReset_Click()
    ' Reset Timer
    btnReset.Visible = False
    With btnStartStop
        .Enabled = True
        .Locked = False
    End With
    ' Reset Text
    h = 0: m = 0: s = 0: ms = 0
    lblH.Caption = "00": lblM.Caption = "00": lblS.Caption = "00": lblMS.Caption = "000"
End Sub

Private Sub btnStartStop_Click()
    ' Toggle Timer Start/Stop
    If running Then
        running = False
        ' Formatting
        btnReset.Visible = True
        With btnStartStop
            .Caption = "Start Timer"
            .ForeColor = RGB(0, 102, 0)
            .BackColor = RGB(204, 255, 204)
            .Enabled = False
            .Locked = True
        End With
    Else
        startTime = Timer()
        running = True
        ' Formatting
        btnReset.Visible = False
        With btnStartStop
            .Caption = "Stop Timer"
            .ForeColor = RGB(102, 0, 0)
            .BackColor = RGB(255, 204, 204)
            .Enabled = True
            .Locked = False
        End With
        ' Run event loop to update text
        Do While running
            PauseForSecs (0.01)
            currentTime = Timer()
            ' Use data to update variables
            h = elapsed - (elapsed Mod 3600)
            m = (elapsed Mod 3600 - s) / 60
            elapsed = currentTime - startTime
            s = elapsed Mod 60
            ms = elapsed * 1000 Mod 1000
            ' Update labels to match variables
            lblH.Caption = Format(h, "00")
            lblM.Caption = Format(m, "00")
            lblS.Caption = Format(s, "00")
            lblMS.Caption = Format(ms, "000")
        Loop
        ' Users can choose not to publish false positives :)
        If MsgBox("Would you like to save this run?", vbYesNo) = vbNo Then Exit Sub ' leave sub if No is pressed
        ' Create new Run and save to Runs database
        With frmMain.rs
            ' Open with special writing permissions, thanks Google :)
            .Open "Runs", frmMain.cn, adOpenKeyset, adLockOptimistic, adCmdTable
            ' Move to end of rs
            Do Until .EOF
                .MoveNext
            Loop
            ' Create new record
            .AddNew
            ' Add all required fields
            frmMain.rs("GameID").Value = frmMain.currentGameID
            frmMain.rs("PlayerID").Value = frmMain.userPlayerID
            frmMain.rs("Time").Value = elapsed
            frmMain.rs("RunDate").Value = Now
            .Update ' update rs
            .Close ' close rs
        End With
        MsgBox "Run saved successfully!", vbOKOnly
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Intialize variables
    running = False
    h = 0: m = 0: s = 0: ms = 0
    ' Formatting
    btnStartStop.Caption = "Start Timer"
    btnStartStop.ForeColor = RGB(0, 102, 0)
    btnStartStop.BackColor = RGB(204, 255, 204)
    ' Call Sub to fill Game data textboxes from rs
    Call frmSelectGame.FillTextBoxesByGameID(frmMain.currentGameID, txtTitle, txtStudio, txtYear)
End Sub

' Little utility function for pausing up to 10ms per cycle (Excel default sleep functions only wait for 1s or greater)
' Only useful for the aesthetics not for the functionality, makes the Timer update really fast, gives you speedrunner vibes!
' Time is actually recorded with Timer() function, this is only used for updating visuals :)
' Source: https://bytes.com/topic/visual-basic/answers/738464-vba-pausing-tenths-seconds-milliseconds
Public Function PauseForSecs(ByVal Delay As Double)
    Dim dblEndTime As Double
    dblEndTime = Timer + Delay
    Do While Timer < dblEndTime
      DoEvents
    Loop
End Function



' ==========================================================+
' ===================== frmNotLoggedIn =====================+
' ==========================================================+

Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Kelvin Kellner
' Student ID: 190668940
' Date: 07-30-2021
' Program title: CP212 A5 Fictional Speedrunning Community
' Form description: Used to launch Log In and Sign Up Forms
'===========================================================+

Private Sub btnLogIn_Click()
    ' Show Form
    frmLogIn.Show
End Sub

Private Sub lblSignUp_Click()
    ' Show Form
    frmSignUp.Show
End Sub



' ==========================================================+
' ==================== frmPlayerProfile ====================+
' ==========================================================+

Option Explicit

Private Sub fraPlayer_Click()

End Sub

' ==== CP212 Windows Application Programming ===============+
' Name: Kelvin Kellner
' Student ID: 190668940
' Date: 07-30-2021
' Program title: CP212 A5 Fictional Speedrunning Community
' Form description: Displays Profile data for another Player
'===========================================================+

Private Sub UserForm_Initialize()
    ' Call Sub to grab Player data from rs
    Call frmPlayers.FillTextboxesFromPlayerID(frmMain.currentPlayerID, txtUsername, txtLocation, txtBio)
    ' Call Sub to fill Runs using rs
    Call FillPlayerRunsFromPlayerID(frmMain.currentPlayerID, lstHeaders, lstRuns)
End Sub

Private Sub btnBack_Click()
    ' Unload Form
    Unload Me
End Sub

Public Sub FillPlayerRunsFromPlayerID(playerID As Integer, listHeaders As MSForms.ListBox, listRuns As MSForms.ListBox)
    ' Fill Runs table ' Fill Headers if not already filled
    If listHeaders.ListCount = 0 Then
        listHeaders.AddItem
        listHeaders.List(0, 0) = "Year"
        listHeaders.List(0, 1) = "Game Title"
        listHeaders.List(0, 2) = "Time"
        listHeaders.List(0, 3) = "Date"
    End If
    
    ' Fill listbox with all Runs from this Player in order of descending date (newest first)
    Dim count As Integer, i As Integer
    ' Clear all contents of listbox
    listRuns.Clear
    count = 0 ' use count to track row in listbox
    With frmMain.rs
        ' Open rs with runs from Player by ID and order by Run Date
        .Open "SELECT * FROM Runs WHERE PlayerID = " & playerID & " ORDER BY RunDate DESC", frmMain.cn
        ' Check for empty rs, meaning no Runs (in search or at all)
        If .EOF Then
            ' Print message to listbox
            listRuns.AddItem
            listRuns.List(0, 0) = "No Runs Yet."
            listRuns.Locked = True ' prevent selecting
        Else
            ' Fill listbox with all Runs in rs
            Do While Not .EOF
                listRuns.AddItem
                listRuns.List(count, 0) = .Fields("GameID")
                listRuns.List(count, 1) = .Fields("GameID")
                listRuns.List(count, 2) = Main.TimeToString(.Fields("Time"))
                listRuns.List(count, 3) = Main.DateToString(.Fields("RunDate"))
                .MoveNext ' move to next record in rs
                count = count + 1
            Loop
            listRuns.Locked = False ' allow selecting
        End If
        .Close ' close rs
    End With
    ' Replace all Game ID with Game Title and Release Year using rs
    With frmMain.rs
        ' Select all Players
        .Open "SELECT * FROM Games"
        For i = 0 To count - 1
            ' If the ID was found in rs
            .Filter = "GameID = " & listRuns.List(i, 0)
            If Not .EOF Then
                listRuns.List(i, 0) = .Fields("ReleaseYear")
                listRuns.List(i, 1) = .Fields("Title")
            Else
                listRuns.List(i, 0) = ""
                listRuns.List(i, 1) = "Deleted game."
            End If
            .Filter = "" ' clear filter, important!!
        Next
        .Close ' close rs
    End With
End Sub



' ==========================================================+
' ======================= frmPlayers =======================+
' ==========================================================+

Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Kelvin Kellner
' Student ID: 190668940
' Date: 07-30-2021
' Program title: CP212 A5 Fictional Speedrunning Community
' Form description: Displays a list of all Players
'===========================================================+

Private Sub UserForm_Initialize()
    ' Fill Headers if not already filled
    If lstHeaders.ListCount = 0 Then
        lstHeaders.AddItem
        lstHeaders.List(0, 0) = "Username"
        lstHeaders.List(0, 1) = "Location"
    End If
    
    ' Fill listbox with all Players to start
    FillPlayersListBox "SELECT * FROM Players"
End Sub

Public Sub btnSearch_Click()
    ' If search bar is not empty
    If Not txtSearch.Text = Empty Then
        ' Then try to search for that Player in database and fill listbox with results
        ' Searches for Username contains or Location contains search textbox text
        FillPlayersListBox "SELECT * FROM Players WHERE Username LIKE '%" & txtSearch.Text & "%' OR Location LIKE '%" & txtSearch.Text & "%'"
    Else
        ' If search bar is empty, then reset Players listbox to include all Players (same as btnClear)
        UserForm_Initialize
    End If
End Sub

Private Sub btnViewProfile_Click()
    ' Check if Player list is actually empty
    If lstPlayers.ListCount = 0 Or lstPlayers.Locked = True Then
        MsgBox "Please try another search.", vbCritical, "No Players Found"
    Else
        ' Find PlayerID for selected Player from rs
        With lstPlayers
            ' If there is a valid entry selected from listbox
            If .ListIndex > -1 Then
                If .Selected(.ListIndex) Then
                    ' Use a SQL Query to find that Player in rs
                    frmMain.rs.Open "SELECT * FROM Players WHERE Username = '" & .List(.ListIndex, 0) & "'", frmMain.cn
                    ' Prevent bad rs errors
                    If Not frmMain.rs.EOF Then
                        ' Update currentPlayerID to match selected Player's PlayerID
                        frmMain.currentPlayerID = frmMain.rs.Fields("PlayerID")
                    Else
                        MsgBox "Could not retrieve player from database.", vbCritical, "Database Error"
                    End If
                    frmMain.rs.Close ' close rs
                    ' Hide current Form
                    Unload Me
                    ' Show the Player Profile Form or if user is selected show the My Profile Form
                    If True Then 'frmMain.currentPlayerID <> frmMain.userPlayerID Then
                        frmPlayerProfile.Show
                    Else
                        frmMyProfile.Show
                    End If
                Else
                    MsgBox "You must select a player to do that.", vbCritical, "No Player Selected"
                End If
            Else
                MsgBox "You must select a player to do that.", vbCritical, "No Player Selected"
            End If
        End With
    End If
End Sub

Private Sub btnBack_Click()
    ' Close Form
    Unload Me
End Sub

Private Sub btnClear_Click()
    ' Call initialize to re-fill all Players
    UserForm_Initialize
End Sub

Private Sub FillPlayersListBox(sqlString As String)
    Dim count As Integer
    ' Clear all contents of listbox
    lstPlayers.Clear
    count = 0 ' use count to track row in listbox
    
    With frmMain.rs
        ' Open rs using SQL String
        .Open sqlString, frmMain.cn
        ' Check for empty rs, meaning no Players (in search or at all)
        If .EOF Then
            ' Print message to listbox
            lstPlayers.AddItem
            lstPlayers.List(0, 0) = "No Players Found."
            lstPlayers.Locked = True ' prevent selecting
        Else
            ' Fill listbox with all Players in rs
            Do While Not .EOF
                lstPlayers.AddItem
                lstPlayers.List(count, 0) = .Fields("Username")
                lstPlayers.List(count, 1) = .Fields("Location")
                .MoveNext ' move to next record in rs
                count = count + 1
            Loop
            lstPlayers.Locked = False ' allow selecting
        End If
        .Close ' close rs
    End With
End Sub

Public Sub FillTextboxesFromPlayerID(playerID As Integer, boxUsername As MSForms.TextBox, boxLocation As MSForms.TextBox, boxBio As MSForms.TextBox)
    ' Select Player from rs by PlayerID
    With frmMain.rs
        .Open "SELECT * FROM Players WHERE PlayerID = " & playerID
        ' If Player Found
        If Not .EOF Then
            ' Fill textboxes with user info
            boxUsername.Text = .Fields("Username")
            boxLocation.Text = .Fields("Location")
            boxBio.Text = .Fields("Bio")
        Else
            ' Print message if not found (should not happen unless there is an error)
            MsgBox "Player could not be found in database.", vbCritical, "Player Not Found"
        End If
        .Close ' close rs
    End With
End Sub



' ==========================================================+
' ======================= frmReport ========================+
' ==========================================================+

Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Kelvin Kellner
' Student ID: 190668940
' Date: 07-30-2021
' Program title: CP212 A5 Fictional Speedrunning Community
' Form description: Reports Word document report of runs
'===========================================================+

Private Sub btnBack_Click()
    ' Close Form
    Unload Me
End Sub

Private Sub btnGenerate_Click()
    ' Generate Word document report of Runs for a Game
    Dim wdApp As Word.Application
    Set wdApp = New Word.Application
    Dim s As Shape
    Dim currGame As Integer: currGame = -1
    Dim count As Integer: count = 0
    Dim i As Integer, j As Integer
    Dim currentRank As Integer, cols As Integer
    Dim sqlString As String
    Dim allGames As Boolean
    
    ' For Run info
    If optAllGames.Value Then
        allGames = True
        sqlString = "SELECT * FROM Runs ORDER BY GameID, Time"
    Else
        allGames = False
        sqlString = "SELECT * FROM Runs WHERE GameID = " & frmMain.currentGameID & " ORDER BY Time"
    End If
    ' Intiliaze array and fill with Run info
    Dim Run()
    With frmMain.rs
        .Open sqlString, frmMain.cn
        Do Until .EOF
            count = count + 1
            ' Prepare Array for all the Run info!
            ReDim Preserve Run(9, count)
            Run(1, count) = .Fields("RunID")
            Run(2, count) = .Fields("GameID")
            ' I know fields 3-5 are wasteful if there is only one game, but I do not want to spend the time to make a better solution right now, put my time somewhere else!
            Run(3, count) = "Game not found."
            Run(4, count) = "" ' Studio
            Run(5, count) = "" ' ReleaseYear
            Run(6, count) = .Fields("PlayerID")
            Run(7, count) = "Deleted user."
            Run(8, count) = .Fields("Time")
            Run(9, count) = .Fields("RunDate")
            .MoveNext
        Loop
    End With
    frmMain.rs.Close ' close rs
    
    ' For Game info
    If allGames Then
        sqlString = "SELECT * FROM Games"
    Else
        sqlString = "SELECT * FROM Games WHERE GameID = " & frmMain.currentGameID
    End If
    ' Fill array with Game info
    With frmMain.rs
        .Open sqlString, frmMain.cn
        Do Until .EOF
            For i = 1 To count
                ' Fill placeholder fields
                If Run(2, i) = .Fields("GameID") Then
                    Run(3, i) = .Fields("Title")
                    Run(4, i) = .Fields("Studio")
                    Run(5, i) = .Fields("ReleaseYear")
                End If
            Next
            .MoveNext
        Loop
    End With
    frmMain.rs.Close ' close rs
    
    ' For Player info
    With frmMain.rs
        .Open "SELECT * FROM Players", frmMain.cn
        Do Until .EOF
            For i = 1 To count
                ' Fill username
                If Run(6, i) = .Fields("PlayerID") Then
                    Run(7, i) = .Fields("Username")
                End If
            Next
            .MoveNext
        Loop
    End With
    frmMain.rs.Close ' close rs
    
    ' Number of columns for chat
    If allGames Then
        cols = 7
    Else
        ' No need to put Game info if there is only data for 1 game
        cols = 4
    End If
    ' Write from Run Array to Word document
    With wdApp
        .Visible = True
        .Activate
        .Documents.Add
        
        ' Create title and subheadings
        With .Selection
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Paragraphs.SpaceAfter = 0
            .Font.Italic = True
            .Font.name = "franklin gothic demi"
            .Font.Size = 16
            .TypeText ("Speedrun Report" & vbNewLine)
            .Paragraphs.SpaceAfter = 1
            .Font.Italic = False
            .Font.name = "franklin gothic heavy"
            If allGames Then
                .TypeText ("All Games" & vbNewLine)
            Else
                .TypeText (Run(3, 1) & " (" & Run(5, 1) & ") by " & Run(4, 1) & vbNewLine) ' Game (Year) by Studio
            End If
            .Font.Size = 6
            .TypeText vbNewLine ' empty space
            .Font.Italic = True
            .Font.name = "franklin gothic book"
            .Font.Size = 12
            .TypeText ("Exported on " & Format(Now, "short date") & vbNewLine)
            .Font.Italic = False
            .Font.name = "calibri"
            .Font.Size = 11
            .TypeParagraph
        End With
        
        ' Create table of Run data
        With .Selection
            .Tables.Add _
                    Range:=wdApp.Selection.Range, _
                    NumRows:=count + 1, NumColumns:=cols, _
                    DefaultTableBehavior:=wdWord9TableBehavior, _
                    AutoFitBehavior:=wdAutoFitContent
            .Rows(1).Range.Font.Bold = True ' bold headers
            ' Set up headers
            If allGames Then
                .TypeText Text:="Game Title"
                .MoveRight Unit:=wdCell
                .TypeText Text:="Studio"
                .MoveRight Unit:=wdCell
                .TypeText Text:="Release Year"
                .MoveRight Unit:=wdCell
            End If
            .TypeText Text:="Rank"
            .MoveRight Unit:=wdCell
            .TypeText Text:="Username"
            .MoveRight Unit:=wdCell
            .TypeText Text:="Time"
            .MoveRight Unit:=wdCell
            .TypeText Text:="Date"
            .MoveRight Unit:=wdCell
            ' Print data from Run
            currentRank = 1
            For i = 1 To count
                ' No need to print Game info if report is only for one Game
                If allGames Then
                    ' Reset Rank to 1 for each new game
                    If currGame = -1 Then
                        currGame = Run(2, i) ' Game ID
                    ElseIf currGame <> Run(2, i) Then
                        currGame = Run(2, i)
                        currentRank = 1
                    Else
                        currentRank = currentRank + 1
                    End If
                    ' Write Game info
                    .TypeText Text:=Run(3, i)
                    .MoveRight Unit:=wdCell
                    .TypeText Text:=Run(4, i)
                    .MoveRight Unit:=wdCell
                    .TypeText Text:=Run(5, i)
                    .MoveRight Unit:=wdCell
                Else
                    currentRank = i
                End If
                ' Write Run and Player info
                .TypeText Text:=currentRank ' Rank
                ' Colour 1st, 2nd, and 3rd differently for the ENHANCED USER EXPERIENCE ;)))
                If currentRank = 1 Then
                    .Shading.BackgroundPatternColor = wdColorLightYellow
                ElseIf currentRank = 2 Then
                    .Shading.BackgroundPatternColor = wdColorGray20
                ElseIf currentRank = 3 Then
                    .Shading.BackgroundPatternColor = wdColorTan
                End If
                .MoveRight Unit:=wdCell
                .TypeText Text:=Run(7, i) ' Username
                .MoveRight Unit:=wdCell
                .TypeText Text:=Main.TimeToString(CSng(Run(8, i))) ' Time
                .MoveRight Unit:=wdCell
                .TypeText Text:=Main.DateToString(CStr(Run(9, i))) ' Date
                ' Leave table instead of creating new row
                If i <> count Then
                    .MoveRight Unit:=wdCell
                Else
                    .MoveDown Unit:=wdLine
                    .TypeText Text:=vbNewLine
                End If
            Next
            
            ' Generate Plot chart for current Game and copy to Word document
            If Not allGames Then
                frmRuns.PlotChart
                .ParagraphFormat.Alignment = wdAlignParagraphLeft
                For Each s In Application.Sheets(frmRuns.shtName).Shapes
                    s.Copy
                    .PasteSpecial _
                    Link:=False, _
                    DataType:=wdPasteEnhancedMetafile, _
                    Placement:=wdInLine, _
                    DisplayAsIcon:=False
                Next s
            End If
        End With
    End With
End Sub

Private Sub UserForm_Click()

End Sub



' ==========================================================+
' ======================== frmRuns =========================+
' ==========================================================+

Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Kelvin Kellner
' Student ID: 190668940
' Date: 07-30-2021
' Program title: CP212 A5 Fictional Speedrunning Community
' Form description: Displays a list of all speedrun Runs for
'                   a particular Game
'===========================================================+

Public shtName As String

Private Sub btnSelectGame_Click()
    ' Show Form
    frmMain.nextForm = "frmRuns"
    Unload Me ' so it will re-initialize after it is finished
    frmSelectGame.Show
End Sub

Private Sub btnBack_Click()
    ' Close Form
    Unload Me
End Sub

Private Sub btnReport_Click()
    ' Show Form for generating Report
    frmReport.Show
End Sub

Private Sub btnPlot_Click()
    ' Call function to plot chart of Runs in chronological order with Player Username
    PlotChart
End Sub


Private Sub btnViewProfile_Click()
    ' Check if Runs list is actually empty
    If lstRuns.ListCount = 0 Or lstRuns.Locked = True Then
        MsgBox "You must select a run to do that.", vbCritical, "No Runs Found"
    Else
        ' Find PlayerID for selected Player from rs
        With lstRuns
            ' If there is a valid entry selected from listbox
            If .ListIndex > -1 Then
                If .Selected(.ListIndex) Then
                    ' Use a SQL Query to find that Player in rs
                    frmMain.rs.Open "SELECT * FROM Players WHERE Username = '" & .List(.ListIndex, 1) & "'", frmMain.cn
                    ' Prevent bad rs errors
                    If Not frmMain.rs.EOF Then
                        ' Update currentPlayerID to match selected Player's PlayerID
                        frmMain.currentPlayerID = frmMain.rs.Fields("PlayerID")
                    Else
                        MsgBox "Could not retrieve player from database.", vbCritical, "Database Error"
                    End If
                    frmMain.rs.Close ' close rs
                    ' Hide current Form
                    Me.Hide
                    ' Show the Player Profile Form or if user is selected show the My Profile Form
                    If frmMain.currentPlayerID = frmMain.userPlayerID Then
                        frmMyProfile.Show
                    Else
                        frmPlayerProfile.Show
                    End If
                Else
                    MsgBox "You must select a run to do that.", vbCritical, "No Run Selected"
                End If
            Else
                MsgBox "You must select a run to do that.", vbCritical, "No Run Selected"
            End If
        End With
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Fill Game info from rs using Sub
    Call frmSelectGame.FillTextBoxesByGameID(frmMain.currentGameID, txtTitle, txtStudio, txtYear)
    
    ' Fill Runs table ' Fill Headers if not already filled
    If lstHeaders.ListCount = 0 Then
        lstHeaders.AddItem
        lstHeaders.List(0, 0) = "Rank"
        lstHeaders.List(0, 1) = "Player"
        lstHeaders.List(0, 2) = "Time"
        lstHeaders.List(0, 3) = "Date"
    End If
    
    ' Fill listbox with all Runs from this Game in order of ascending time
    FillRunsListBox "SELECT * FROM Runs WHERE GameID = " & frmMain.currentGameID & " ORDER BY Time ASC"
End Sub

Private Sub FillRunsListBox(sqlString As String)
    Dim count As Integer, i As Integer
    ' Clear all contents of listbox
    lstRuns.Clear
    count = 0 ' use count to track row in listbox
    With frmMain.rs
        ' Open rs using SQL String
        .Open sqlString, frmMain.cn
        ' Check for empty rs, meaning no Runs (in search or at all)
        If .EOF Then
            ' Print message to listbox
            lstRuns.AddItem
            lstRuns.List(0, 0) = "No Runs Yet."
            lstRuns.Locked = True ' prevent selecting
        Else
            ' Fill listbox with all Runs in rs
            Do While Not .EOF
                lstRuns.AddItem
                lstRuns.List(count, 0) = count + 1
                lstRuns.List(count, 1) = .Fields("PlayerID")
                lstRuns.List(count, 2) = Main.TimeToString(.Fields("Time"))
                lstRuns.List(count, 3) = Main.DateToString(.Fields("RunDate"))
                .MoveNext ' move to next record in rs
                count = count + 1
            Loop
            lstRuns.Locked = False ' allow selecting
        End If
        .Close ' close rs
    End With
    ' Replace all Player ID with Player Username using rs
    With frmMain.rs
        ' Select all Players
        .Open "SELECT * FROM Players"
        For i = 0 To count - 1
            ' If the ID was found in rs
            .Filter = "PlayerID = " & lstRuns.List(i, 1)
            If Not .EOF Then
                lstRuns.List(i, 1) = .Fields("Username")
            Else
                lstRuns.List(i, 1) = "Deleted user."
            End If
            .Filter = "" ' clear filter, important!!
        Next
        .Close ' close rs
    End With
End Sub

Public Sub PlotChart()
    ' Great line chart of Run data for Game by Player
    Dim sht As Worksheet
    Dim rng As Range
    Dim i As Integer, j As Integer, k As Integer
    Dim count As Integer: count = 0
    Dim playerCount As Integer: playerCount = 0
    Dim newPlayer As Boolean
    Dim cht As Chart
    
    shtName = txtTitle.Text & " Speedruns Plotted"
    
    ' Fill Arrays with Run and Player data
    Dim Run()
    Dim Player()
    With frmMain.rs
        .Open "SELECT * FROM Runs WHERE GameID = " & frmMain.currentGameID & " ORDER BY RunDate", frmMain.cn
        Do Until .EOF
            count = count + 1
            ReDim Preserve Run(4, count)
            Run(1, count) = .Fields("PlayerID")
            newPlayer = True
            For i = 1 To playerCount
                If Player(i) = Run(1, count) Then
                    newPlayer = False
                End If
            Next
            If newPlayer Then
                playerCount = playerCount + 1
                ReDim Preserve Player(playerCount)
                Player(playerCount) = Run(1, count)
            End If
            Run(2, count) = "Deleted user."
            Run(3, count) = .Fields("Time")
            Run(4, count) = .Fields("RunDate")
            .MoveNext
        Loop
    End With
    frmMain.rs.Close ' close rs
    
    ' Replace PlayerID with Username
    With frmMain.rs
        .Open "SELECT * FROM Players"
        Do Until .EOF
            For i = 1 To count
                If Run(1, i) = .Fields("PlayerID") Then
                    Run(2, i) = .Fields("Username")
                End If
            Next
            .MoveNext
        Loop
    End With
    frmMain.rs.Close ' close rs
    
    ' Delete old worksheet if it exists
    For Each sht In ThisWorkbook.Worksheets
        If sht.name = shtName Then
            Application.DisplayAlerts = False
            Sheets(shtName).Delete
            Application.DisplayAlerts = True
        End If
    Next sht
    
    ' Create new worksheet for table and chart
    Sheets.Add(After:=Sheets(Sheets.count)).name = shtName
    Set sht = Sheets(shtName)
    'Activate sheet
    sht.Activate
    'Select cell A1 in active worksheet
    Range("A1").Select
    'Zoom to first cell
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    
    ' Print Run Array to Worksheet
    Const cols = 4
    With sht.Range("A1")
        ' Fill Headers
        .Offset(0, 0).Value = "Speedrun #"
        .Offset(0, 1 + playerCount).Value = "Date"
        .Offset(0, 0).Font.Bold = True
        .Offset(0, 1 + playerCount).Font.Bold = True
        For j = 1 To playerCount
            .Offset(0, j).Font.Bold = True
        Next
        ' Fill Data
        For i = 1 To count
            For j = 1 To playerCount
                If Run(1, i) = Player(j) Then
                    .Offset(0, j).Value = Run(2, i) ' Username
                    k = j
                End If
            Next
            .Offset(i, 0).Value = i ' #
            .Offset(i, 1 + playerCount).Value = Main.DateToString(CStr(Run(4, i))) ' Date
            .Offset(i, k).Value = Run(3, i) ' Time
        Next
        .Offset(0, 2 + playerCount).Select ' Select Cell for Chart Left
    End With
    sht.Columns.AutoFit ' AutoFit Column width
    
    ' Create line chart
    Set cht = sht.Shapes.AddChart(xlLine, ActiveCell.Left, ActiveCell.Top).Chart
    With cht
        .SetSourceData Source:=sht.Range("B1", Cells(count, 1 + playerCount))
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = txtTitle.Text & " Speedruns"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Time"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Characters.Text = "Speedrun #"
    End With
End Sub



' ==========================================================+
' ====================== frmSelectGame =====================+
' ==========================================================+

Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Kelvin Kellner
' Student ID: 190668940
' Date: 07-30-2021
' Program title: CP212 A5 Fictional Speedrunning Community
' Form description: Used in various locations throughout the
'                   program to select a Game, edit or add a
'                   new Game to the list of Games
'===========================================================+

Private Sub btnBack_Click()
    ' Close Form
    Unload Me
End Sub

Private Sub btnClear_Click()
    ' Call initialize to re-fill all Games
    UserForm_Initialize
End Sub

Private Sub btnEditGame_Click()
    ' Check if Games list is actually empty
    If lstGames.ListCount = 0 Or lstGames.Locked = True Then
        MsgBox "Please try another search.", vbCritical, "No Games Found"
    Else
        ' Find Game ID for selected Game from rs
        With lstGames
            ' If there is a valid entry selected from listbox
            If .ListIndex > -1 Then
                If .Selected(.ListIndex) Then
                    ' Use a SQL Query to find that Game in rs
                    frmMain.rs.Open "SELECT * FROM Games WHERE Title = '" & .List(.ListIndex, 0) & "' AND Studio = '" & .List(.ListIndex, 1) & "' AND ReleaseYear = " & .List(.ListIndex, 2), frmMain.cn
                    ' Prevent bad rs errors
                    If Not frmMain.rs.EOF Then
                        ' Update currentGameID to match selected Game's GameID
                        frmMain.currentGameID = frmMain.rs.Fields("GameID")
                    Else
                        ' Should not happen unless error
                        MsgBox "Could not retrieve game from database.", vbCritical, "Database Error"
                    End If
                    frmMain.rs.Close ' close rs
                    ' Show the Edit Game Form or if user is selected show the My Profile Form
                    frmEditGame.Show
                Else
                    MsgBox "You must select a game to do that.", vbCritical, "No Game Selected"
                End If
            Else
                MsgBox "You must select a game to do that.", vbCritical, "No Game Selected"
            End If
        End With
    End If
End Sub

Private Sub btnNewGame_Click()
    ' Open Form
    frmAddGame.Show
End Sub

Private Sub btnNext_Click()
    ' Check if Games list is actually empty
    If lstGames.ListCount = 0 Or lstGames.Locked = True Then
        MsgBox "Please try another search, or add a new game.", vbCritical, "No Games Found"
    Else
        ' Find GameID for selected Game from rs
        With lstGames
            ' If there is a valid entry selected from listbox
            If .ListIndex > -1 Then
                If .Selected(.ListIndex) Then
                    ' Use an SQL Query to find that Game in rs
                    frmMain.rs.Open "SELECT * FROM Games WHERE Title = '" & .List(.ListIndex, 0) & "' AND Studio = '" & .List(.ListIndex, 1) & "' AND ReleaseYear = " & .List(.ListIndex, 2), frmMain.cn
                    ' Prevent bad rs errors
                    If Not frmMain.rs.EOF Then
                        ' Update currentGameID to match selected Game's GameID
                        frmMain.currentGameID = frmMain.rs.Fields("GameID")
                    Else
                        MsgBox "Could not retrieve game from database.", vbCritical, "Database Error"
                    End If
                    frmMain.rs.Close ' close rs
                    ' Hide current Form
                    Me.Hide
                    ' Show the appropriate Form
                    If frmMain.nextForm = "frmNewAttempt" Then
                        frmMain.nextForm = ""
                        frmNewAttempt.Show
                    ElseIf frmMain.nextForm = "frmRuns" Then
                        frmMain.nextForm = ""
                        frmRuns.Show
                    End If
                Else
                    MsgBox "Please select a game to continue.", vbCritical, "No Game Selected"
                End If
            Else
                MsgBox "Please select a game to continue.", vbCritical, "No Game Selected"
            End If
        End With
    End If
End Sub

Private Sub btnSearch_Click()
    ' If search bar is not empty
    If Not txtSearch.Text = Empty Then
        ' Then try to search for that Game in database and fill listbox with results
        ' Searches for Title contains or Studio contains search textbox text
        FillGamesListBox "SELECT * FROM Games WHERE Title LIKE '%" & txtSearch.Text & "%' OR Studio LIKE '%" & txtSearch.Text & "%'"
    Else
        ' If search bar is empty, then reset Games listbox to include all Games (same as btnClear)
        UserForm_Initialize
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Fill Headers if not already filled
    If lstHeaders.ListCount = 0 Then
        lstHeaders.AddItem
        lstHeaders.List(0, 0) = "Title"
        lstHeaders.List(0, 1) = "Studio"
        lstHeaders.List(0, 2) = "Year"
    End If
    
    ' Fill listbox with all Games to start
    FillGamesListBox "SELECT * FROM Games"
End Sub

Private Sub FillGamesListBox(sqlString As String)
    Dim count As Integer
    ' Clear all contents of listbox
    lstGames.Clear
    count = 0 ' use count to track row in listbox
    With frmMain.rs
        ' Open rs using SQL String
        .Open sqlString, frmMain.cn
        ' Check for empty rs, meaning no Games (in search or at all)
        If .EOF Then
            ' Print message to listbox
            lstGames.AddItem
            lstGames.List(0, 0) = "No Games Found."
            lstGames.Locked = True ' prevent selecting
        Else
            ' Fill listbox with all Games in rs
            Do While Not .EOF
                lstGames.AddItem
                lstGames.List(count, 0) = .Fields("Title")
                lstGames.List(count, 1) = .Fields("Studio")
                lstGames.List(count, 2) = .Fields("ReleaseYear")
                .MoveNext ' move to next record in rs
                count = count + 1
            Loop
            lstGames.Locked = False ' allow selecting
        End If
        .Close ' close rs
    End With
End Sub

Public Sub FillTextBoxesByGameID(GameID As Integer, boxTitle As MSForms.TextBox, boxStudio As MSForms.TextBox, boxYear As MSForms.TextBox)
    ' Open rs for the Game by GameID
    With frmMain.rs
        .Open "SELECT * FROM Games WHERE GameID = " & GameID, frmMain.cn
        ' If Game was found
        If Not .EOF Then
            ' Fill textboxes with data from rs
            boxTitle.Text = .Fields("Title")
            boxStudio.Text = .Fields("Studio")
            boxYear.Text = .Fields("ReleaseYear")
        Else
            ' Print message if not found (should not happen unless there is an error)
            MsgBox "Game could not be found in database.", vbCritical, "Game Not Found"
        End If
        .Close ' close rs
    End With
End Sub



' ==========================================================+
' ======================= frmSignUp ========================+
' ==========================================================+

Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Kelvin Kellner
' Student ID: 190668940
' Date: 07-30-2021
' Program title: CP212 A5 Fictional Speedrunning Community
' Form description: Allows for creation of new Player accounts
'===========================================================+

' On Submit
Private Sub btnSignUp_Click()
    ' Create new Player in Access database if appropriate
    Dim msg(1 To 3) As String, taken As Boolean
    ' Retrieve text from textboxes
    msg(1) = txtUsername.Text
    msg(2) = txtLocation.Text
    msg(3) = txtBio.Text
    
    ' Check for empty username textbox
    If msg(1) = Empty Then
        ' Error when User does not enter a username
        MsgBox "Please enter a username.", vbCritical, "Blank Username"
    Else
        taken = False ' default to False
        ' Search through all Players and see if User account already taken
        With frmMain.rs
            ' Open with special writing permissions, thanks Google :)
            .Open "Players", frmMain.cn, adOpenKeyset, adLockOptimistic, adCmdTable
            ' Loop until we find that username is already taken or reach end of rs query
            Do Until .EOF Or taken = True
                ' If username in rs matches username textbox text
                If .Fields("Username") = msg(1) Then
                    ' Set taken to True
                    taken = True
                End If
                ' Move to next Player in rs
                .MoveNext
            Loop
            
             ' Perform appropriate action for taken or available
            If taken = True Then
                ' Account with that username already taken
                MsgBox "An account with that username already exists in our database.", vbCritical, "Username Already Taken"
            Else
                ' Is not taken... now we can create the Player's account
                ' Create new record in rs and populate with values as text from textboxes
                .AddNew
                frmMain.rs("Username").Value = msg(1)
                frmMain.rs("Location").Value = msg(2)
                frmMain.rs("Bio").Value = msg(3)
                .Update
                frmMain.userPlayerID = frmMain.rs("PlayerID").Value
            End If
            .Close ' close rs
        End With
        
        ' Close form after successful account creation
        If taken = False Then
            ' Close the Sign Up Forms and open Main Form dashboard
            Unload Me
            Unload frmNotLoggedIn
            frmMain.Show
        End If
    End If
End Sub
