Imports System.Threading
Imports System.Net.Sockets
Imports System.IO

Public Class frmGame
    '------------------------------------------------------------
    '-                File Name : frmGame.vb                    - 
    '-                Part of Project: Assignment11             -
    '------------------------------------------------------------
    '-                Written By: Benjamin R. Neeb              -
    '-                Written On: April 20, 2021                -
    '------------------------------------------------------------
    '- File Purpose:                                            -
    '-                                                          -
    '- This file contains the all code for the application. All -
    '- networking and user interaction are done here.           -
    '------------------------------------------------------------
    '- Program Purpose:                                         -
    '-                                                          -
    '- The purpose of this program is to allow a user to choose -
    '- to either host or join a game of Memory. Once the user   -
    '- selects his role, he can begin playing. Score is kept    -
    '- and once all the cards have been flipped, the player     -
    '- with the most pairs wins!                                -
    '------------------------------------------------------------
    '- Global Variable Dictionary (alphabetically):             -
    '- aConnection:     Socket for connection.                  -
    '- Client:          TcpClient object represents a client    -
    '- blnConnected:    Boolean to keep track of connection     -
    '-                  status.                                 -
    '- GetDataThread:   Thread for listening to incoming        -
    '-                  connections.                            -
    '- arrLetters():    Char array to hold randomized letters.  -
    '- NetStream:       NetworkStream to transfer data.         -
    '- NetWriter:       Writer for writing data accros network. -
    '- NetReader:       Reader for reading data accros network. -
    '- intP1Score:      Integer to keep track of P1's score.    -
    '- intP2Score:      Integer to keep track of P2's score.    -
    '- blnPlayerCtr:    Boolean to track number of consequetive -
    '-                  turns a player has taken.               -
    '- blnPlayerMe:     Boolean to determine which player       -
    '-                  (1 or 2) is assigned to me.             -
    '- blnPlayerTurn:   Boolean to track player's turns         -
    '- Server:          TcpListener object represents our       -
    '-                  server                                  -
    '------------------------------------------------------------

    '---------------------------------------------------------------------------------------
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '---------------------------------------------------------------------------------------

    'The number of cards on the game board
    Const intCARDNUMBER As Integer = 36

    '---------------------------------------------------------------------------------------
    '--- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES ---
    '--- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES ---
    '--- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES ---
    '---------------------------------------------------------------------------------------

    'This TcpListener object represents our server
    Dim Server As TcpListener

    'This TcpClient object represents a client
    Dim Client As TcpClient

    'We need to create a Socket object to associate with our 
    'server.  Our socket will run on port 1000 by default.
    Dim aConnection As Socket

    'We need a NetworkStream through which data is transferred
    'between the client and the server
    Dim NetStream As NetworkStream

    'These are the objects that we will use for reading and 
    'writing data across the network stream
    Dim NetWriter As BinaryWriter
    Dim NetReader As BinaryReader

    'We will have to start up a thread that specifically listens
    'for network stream traffic coming to our server from the
    'client.  This is the thread object we will use for that 
    'purpose.
    Dim GetDataThread As Thread

    'Boolean to keep track of our connection
    Dim blnConnected As Boolean = False

    'Dim dctLetters As Dictionary(Of Integer, Char) = New Dictionary(Of Integer, Char)
    Dim arrLetters(35) As Char

    'Tracks how many consecutive turns a player has taken
    Dim blnPlayerTurn As Boolean = False

    'Tracks which player's turn it is
    Dim blnPlayerCtr As Boolean = False

    Dim blnPlayerMe As Boolean

    'Score keepers
    Dim intP1Score = 0
    Dim intP2Score = 0

    '-----------------------------------------------------------------------------------
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '-----------------------------------------------------------------------------------

    Private Sub frmGame_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '------------------------------------------------------------
        '-            Subprogram Name: frmGame_Load                 -
        '------------------------------------------------------------
        '-                Written By: Benjamin R. Neeb              -
        '-                Written On: April 17, 2021                -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This subroutine is called on form load. This sub sets    -
        '- CheckForIllegalCrossThreadCalls to False.                -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- sender – Identifies which particular control that raised -
        '-          the click event                                 - 
        '- e – Holds the EventArgs object sent to the routine       -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        '.NET gets mad if there are threads used across controls, 
        'so disable it...  Again, not procedurally the correct way 
        'to do this, but okay for this program
        CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub rdoServer_CheckedChanged(sender As Object, e As EventArgs) Handles rdoServer.CheckedChanged
        '------------------------------------------------------------
        '-            Subprogram Name: rdoServer_CheckedChanged     -
        '------------------------------------------------------------
        '-                Written By: Benjamin R. Neeb              -
        '-                Written On: April 17, 2021                -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This subroutine is called when the server radiobutton is -
        '- checked. It shows only the options available to the      -
        '- server role.                                             -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- sender – Identifies which particular control that raised -
        '-          the click event                                 - 
        '- e – Holds the EventArgs object sent to the routine       -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        'Change controls
        btnStart.Text = "Start Server"
        btnStop.Text = "Stop Server"

        'Show controls
        grpGame.Visible = True
    End Sub

    Private Sub rdoClient_CheckedChanged(sender As Object, e As EventArgs) Handles rdoClient.CheckedChanged
        '------------------------------------------------------------
        '-            Subprogram Name: rdoClient_CheckedChanged     -
        '------------------------------------------------------------
        '-                Written By: Benjamin R. Neeb              -
        '-                Written On: April 17, 2021                -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This subroutine is called when the client radiobutton is -
        '- checked. It shows only the options available to the      -
        '- client role.                                             -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- sender – Identifies which particular control that raised -
        '-          the click event                                 - 
        '- e – Holds the EventArgs object sent to the routine       -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        'Change controls
        btnStart.Text = "Start Client"
        btnStop.Text = "Stop Client"

        'Hide controls
        grpGame.Visible = False

    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        '------------------------------------------------------------
        '-            Subprogram Name: btnStart_Click               -
        '------------------------------------------------------------
        '-                Written By: Benjamin R. Neeb              -
        '-                Written On: April 17, 2021                -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This subroutine is called when the user clicks the start -
        '- button. It first checks the game's configuration and     -
        '- attempts to make a connection.                           -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- sender – Identifies which particular control that raised -
        '-          the click event                                 - 
        '- e – Holds the EventArgs object sent to the routine       -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        Try
            If rdoServer.Checked Then
                'Assign player
                If rdoPlayer1.Checked Then
                    blnPlayerMe = True
                Else
                    blnPlayerMe = False
                End If

                'Assign turn
                If rdoFirst1.Checked Then
                    blnPlayerTurn = True
                Else
                    blnPlayerTurn = False
                End If

                CreatePairs()
                StartConnection(True)
            Else
                StartConnection(False)
            End If

            'Disable controls
            grpNetwork.Enabled = False
            grpGame.Enabled = False
            btnStart.Enabled = False

            'Enable controls
            btnStop.Enabled = True

        Catch ex As Exception
            MessageBox.Show("Could not start")
        End Try
    End Sub

    Private Sub btnStop_Click(sender As Object, e As EventArgs) Handles btnStop.Click
        '------------------------------------------------------------
        '-            Subprogram Name: btnStop_Click                -
        '------------------------------------------------------------
        '-                Written By: Benjamin R. Neeb              -
        '-                Written On: April 17, 2021                -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This subroutine is called when the user clicks the stop  -
        '- button. It terminates any active connections and resets  -
        '- the game for future use.                                 -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- sender – Identifies which particular control that raised -
        '-          the click event                                 - 
        '- e – Holds the EventArgs object sent to the routine       -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        Try
            txtLog.Text &= "Leaving Game..." & vbCrLf

            NetWriter.Write("~~END~~")

        Catch ex As Exception
            'We don't need to do anything, but there was a problem
            'communicating with the server
        End Try

        Try
            'Destroy all of the objects that we created
            NetWriter.Close()
            NetReader.Close()
            NetStream.Close()

            If rdoServer.Checked Then
                Server.Stop()
            Else
                Client.Close()
            End If

            NetWriter = Nothing
            NetReader = Nothing
            NetStream = Nothing
            Server = Nothing
            Client = Nothing

            'Clear char array
            Array.Clear(arrLetters, 0, 35)

            'Wipe button text
            'Check cards
            For Each ctrl As Control In Controls
                If TypeOf ctrl Is Button Then
                    Dim btn As Button = ctrl

                    'Check if it's a letter card
                    If Not (btn.Name = "btnStart" Or btn.Name = "btnStop") Then
                        'Empty button text
                        btn.Text = ""

                        'Disable button
                        btn.Enabled = False

                        'Reset color
                        btn.BackColor = Color.Transparent
                    End If
                End If
            Next

            Try
                GetDataThread.Abort()
            Catch Ex As Exception
                'We don't care since we are trying to stop the thread
            End Try

        Catch Ex As Exception
            'We don't have to do anything since we are leaving anyway

        Finally
            txtLog.Text &= "Disconnected...client closed" & vbCrLf
        End Try

        blnConnected = False

        'Disable controls
        btnStop.Enabled = False

        'Enable controls
        grpNetwork.Enabled = True
        grpGame.Enabled = True
        btnStart.Enabled = True
    End Sub

    Private Sub CreatePairs()
        '------------------------------------------------------------
        '-            Subprogram Name: CreatePairs                  -
        '------------------------------------------------------------
        '-                Written By: Benjamin R. Neeb              -
        '-                Written On: April 17, 2021                -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This subroutine creates the randomized letter pairs for  -
        '- the game by creating them within a List of Char and then -
        '- populating the global array arrLetters().                -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- lstLetters:      List of Char to hold initial letters.   -
        '- random:          Random integer instance.                -
        '- intRandom:       Random integer for assigning random     -
        '-                  letter.                                 -
        '------------------------------------------------------------

        'Create list for placing letters
        Dim lstLetters As List(Of Char) = New List(Of Char)

        'Populate list of letters (half as many cards)
        For i As Integer = 0 To intCARDNUMBER / 2 - 1
            lstLetters.Add(Chr(i + 65))
            lstLetters.Add(Chr(i + 65))
        Next

        'Declare random object here for efficiency
        Dim random As New Random

        For i As Integer = 0 To 35

            'Create random number
            Dim intRandom As Integer = random.Next(0, lstLetters.Count - 1)

            'Add random letter to dictionary
            'dctLetters.Add(i, lstLetters(intRandom))
            arrLetters(i) = lstLetters(intRandom)

            'Remove that letter
            lstLetters.RemoveAt(intRandom)
        Next
    End Sub

    Private Sub StartConnection(blnServer As Boolean)
        '------------------------------------------------------------
        '-            Subprogram Name: StartConnection              -
        '------------------------------------------------------------
        '-                Written By: Benjamin R. Neeb              -
        '-                Written On: April 17, 2021                -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This subroutine begins the actual connection. The        -
        '- boolean passed as a parameter determines if we are       -
        '- connecting either the server or client.                  -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- blnServer:       Boolean to determine role.              -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- strArray:        String that holds the randomized chars  -
        '-                  that will be passed to the client.      -
        '------------------------------------------------------------

        Try
            txtLog.Text &= "Attempting connection..." & vbCrLf

            If blnServer Then
                'Create the server and point it at the port from 
                'the textbox value that the user entered.
                Server = New TcpListener(Net.IPAddress.Parse(txtAddress.Text), CInt(txtPort.Text))
                Server.Start()

                'We wait here until we get a connection request from a
                'client...  The server is running and when we get a connection,
                'we will accept it and place it in the Socket object we created.
                txtLog.Text &= "Listening for client connection..." & vbCrLf
                Application.DoEvents()
                aConnection = Server.AcceptSocket()
                txtLog.Text &= "...client connection accepted" & vbCrLf

                'If we get to this point with no exceptions, then we have
                'accepted a request from a client. Now we need to get the 
                'NetworkStream that is associated with our Socket object.
                NetStream = New NetworkStream(aConnection)
            Else
                'Create the client and point it at the server's address
                'and port from the textbox values that the user entered.
                'We will get an exception here if the server is not already
                'up and running.
                Client = New TcpClient()
                Client.Connect(txtAddress.Text, CInt(txtPort.Text))

                'If we get to this point with no exceptions, then we have
                'requested a connection to the server and it was accepted.
                'Now we need to get the NetworkStream that is associated 
                'with our TcpClient.
                NetStream = Client.GetStream()
            End If

            'The last major setup piece that we need to do is to
            'create objects for transferring data across the 
            'NetworkStream.  Bind our Reader and Writer to the
            'NetworkStream object
            NetWriter = New BinaryWriter(NetStream)
            NetReader = New BinaryReader(NetStream)

            txtLog.Text &= "Network stream and reader/writer objects created" & vbCrLf

            'Set up our thread to listen for data arriving
            txtLog.Text &= "Preparing thread to watch for data..." & vbCrLf
            GetDataThread = New Thread(AddressOf GetData)
            GetDataThread.Start()

            'Send randomized letters to client
            Dim strArray As String = New String(arrLetters)
            NetWriter.Write(strArray)

            'Send information about players
            NetWriter.Write("S-" & "," & blnPlayerMe & "," & blnPlayerTurn)

            blnConnected = True

        Catch IOException As IOException
            txtLog.Text &= "Error in setting up Connection -- Closing" & vbCrLf

            btnStop_Click("", Nothing)
        Catch SocketEx As SocketException
            txtLog.Text &= "Server error. Unfound or already exists" & vbCrLf

            btnStop_Click("", Nothing)
        End Try
    End Sub

    Sub GetData()
        '------------------------------------------------------------
        '-            Subprogram Name: GetData                      -
        '------------------------------------------------------------
        '-                Written By: Benjamin R. Neeb              -
        '-                Written On: April 17, 2021                -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This subroutine is a separate thread that recieves any   -
        '- and all data transmitted from the other player.          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- strData:         String that holds the data recieved.    -
        '------------------------------------------------------------

        'This is the routine that we spin off into its own thread to 
        'listen for and retrieve network traffic.

        'This is a string that we use to pull the data off of the network stream
        Dim strData As String

        txtLog.Text &= "Data watching thread active" & vbCrLf

        'Here's the main listening loop that will continue until we
        'receive the ~~END~~ message or the connection abruptly stops.
        Try
            Do
                'Pull data from the network into our string
                strData = NetReader.ReadString

                ProcessData(strData)
            Loop While (strData <> "~~END~~") 'And aConnection.Connected FIND ME commented this out...

            btnStop_Click("", Nothing)

            'Errors can occur if we try to write and it's not there
        Catch IOEx As IOException
            txtLog.Text &= "Closing connection..." & vbCrLf
            btnStop_Click("", Nothing)
        End Try
    End Sub

    Private Sub ProcessData(strData As String)
        '------------------------------------------------------------
        '-            Subprogram Name: ProcessData                  -
        '------------------------------------------------------------
        '-                Written By: Benjamin R. Neeb              -
        '-                Written On: April 17, 2021                -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- Once the data is recieved from the GetData thread, this  -
        '- sub is called to assign and handle the different.        -
        '- operations.                                              -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- strData:     String recieved from other player.          -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- strSplit:        String array formed from splitting      -
        '-                  received data.                          -
        '------------------------------------------------------------

        'Check if we're recieving the random letters from server
        If strData.Length = intCARDNUMBER And rdoClient.Checked Then

            'Assign array the string value
            For i As Integer = 0 To intCARDNUMBER - 1
                arrLetters(i) = strData(i)
            Next

            'Alert of succesful operation
            txtLog.Text &= "Random cards synced from server." & vbCrLf
        ElseIf strData.StartsWith("S-") Then

            'Split data
            Dim strSplit() As String = Split(strData, ",")

            'Assign player information
            If rdoClient.Checked Then
                'What player am I?
                blnPlayerMe = Not (Boolean.Parse(strSplit(1)))

                'Who turn is it?
                blnPlayerTurn = Boolean.Parse(strSplit(2))
            End If

            'Set cards enabled/disabled
            TurnChange(blnPlayerTurn)
        ElseIf strData.Contains("T-") Then
            'Split data
            Dim strSplit() As String = Split(strData, ",")

            'Who turn is it?
            blnPlayerTurn = Boolean.Parse(strSplit(1))
            TurnChange(blnPlayerTurn)

            'Decrement(make false) turns
            blnPlayerCtr = 0
        ElseIf strData.StartsWith("btn") Then
            'Perform card operations
            CardOps("Player", strData)
        End If

        'Log info
        If rdoServer.Checked Then
            txtLog.Text &= "From Client: " & strData & vbCrLf
        Else
            txtLog.Text &= "From Server: " & strData & vbCrLf
        End If

        'Catch-up GUI
        Application.DoEvents()
    End Sub

    Private Sub CardClick(sender As Object, e As EventArgs) Handles btn0.Click, btn1.Click, btn2.Click, btn3.Click, btn4.Click, btn5.Click, btn6.Click, btn7.Click, btn8.Click, btn9.Click, btn10.Click, btn11.Click, btn12.Click, btn13.Click, btn14.Click, btn15.Click, btn16.Click, btn17.Click, btn18.Click, btn19.Click, btn20.Click, btn21.Click, btn22.Click, btn23.Click, btn24.Click, btn25.Click, btn26.Click, btn27.Click, btn28.Click, btn29.Click, btn30.Click, btn31.Click, btn32.Click, btn33.Click, btn34.Click, btn35.Click

        If blnConnected Then
            'All buttons are named btn#. Get the number of the button by doing substring 3
            Dim strButton As String = CType(sender, Button).Name.ToString

            'If we are flipping for the other player's click
            CardOps(sender, strButton)

        End If
    End Sub

    Private Sub CardOps(sender As Object, strButton As String)
        '------------------------------------------------------------
        '-            Subprogram Name: CardOps                      -
        '------------------------------------------------------------
        '-                Written By: Benjamin R. Neeb              -
        '-                Written On: April 17, 2021                -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This subroutine handles the click event of all the card  -
        '- buttons.                                                 -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- sender – Identifies which particular control that raised -
        '-          the click event                                 - 
        '- e – Holds the EventArgs object sent to the routine       -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- intAns:      Integer for user response.                  -
        '- btn:         Button instance to gather name and text of  -
        '-              sender object.                              -
        '- blnMatched:  Boolean to determine if a match was made.   -
        '- blnWon:      Boolean to determine if all pairs have been -
        '-              matched.                                    -
        '- strWon:      String message used to congratulate winner. -
        '------------------------------------------------------------

        Dim blnMatched As Boolean = False
        Dim blnWon As Boolean = True

        'Show card letter value
        Me.Controls(strButton).Text = arrLetters(strButton.Substring(3))

        'Check cards
        For Each ctrl As Control In Controls
            If TypeOf ctrl Is Button Then
                Dim btn As Button = ctrl

                'Check for matching card
                If btn.Text = Me.Controls(strButton).Text And btn.Name <> Me.Controls(strButton).Name Then
                    'Set flag to true
                    blnMatched = True

                    'Call sub to take care of matching
                    Matched(btn, Me.Controls(strButton))
                End If

                'Check if the player has won
                If btn.Name <> "btnStart" And btn.Name <> "btnStop" And btn.BackColor <> Color.LightGreen Then
                    blnWon = False
                End If
            End If
        Next

        'If it's my turn, show the card to the other player
        If blnPlayerTurn = blnPlayerMe Then
            NetWriter.Write(strButton)

            If blnMatched Then
                UpdateScore(blnPlayerTurn)
            End If

            'If not, check if we've reached our max turns
            If blnPlayerCtr = False Then
                'Unless there was a match
                If Not blnMatched Then
                    'If not, increment (make true)
                    blnPlayerCtr = True
                End If
            Else
                'Who turn is it?
                TurnChange(Not blnPlayerTurn)

                'Decrement(make false) turns
                blnPlayerCtr = 0

                'Send information to start timer, change turns, and decrement turns
                NetWriter.Write("T-" & "," & blnPlayerTurn)
            End If
        ElseIf blnMatched Then
            UpdateScore(Not blnPlayerMe)
        End If

        If blnWon Then
            Dim strWon As String

            If intP1Score > intP2Score Then
                lblMessage.Text = "Player 1 has won!"
                strWon = "Congratulations Player 1!"
            ElseIf intP1Score < intP2Score Then
                lblMessage.Text = "Player 2 has won!"
                strWon = "Congratulations Player 2!"
            Else
                lblMessage.Text = "Tie Game!"
                strWon = "Congratulations to both Players!"
            End If

            Dim intAns As Integer = MsgBox(strWon & vbCrLf & "Do you want to restart the game?", MsgBoxStyle.OkCancel, "Game Complete")
            If MsgBoxResult.Ok Then
                btnStop_Click(sender, Nothing)
            End If
        End If
    End Sub

    Private Sub Matched(btnFirst As Button, btnSecond As Button)
        '------------------------------------------------------------
        '-            Subprogram Name: Matched                      -
        '------------------------------------------------------------
        '-                Written By: Benjamin R. Neeb              -
        '-                Written On: April 17, 2021                -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This subroutine handles the buttons when the user        -
        '- matches a pair.                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- btnFirst:    First button matched.                       -
        '- btnSecond:   Second button matched.                      -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        'Disable buttons
        btnFirst.Enabled = False
        btnSecond.Enabled = False

        'Change button colors
        btnFirst.BackColor = Color.LightGreen
        btnSecond.BackColor = Color.LightGreen

        'Reset turn counter
        blnPlayerCtr = False

        'Let the GUI catch up
        Application.DoEvents()
    End Sub

    Private Sub UpdateScore(blnUpdate)
        '------------------------------------------------------------
        '-            Subprogram Name: UpdateScore                  -
        '------------------------------------------------------------
        '-                Written By: Benjamin R. Neeb              -
        '-                Written On: April 17, 2021                -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This subroutine handles the score updates when a user    -
        '- makes a match.                                           -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- blnUpdate:       Holds which user's score to update.     -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        'Who's score are we incremengint?
        If blnUpdate = True Then
            'Increment P1's score
            intP1Score += 1
        Else
            'Increment P2's score
            intP2Score += 1
        End If

        'Update score label
        lblScore.Text = "Player 1: " & intP1Score & " --- Player 2: " & intP2Score
    End Sub

    Private Sub TurnChange(blnNewTurn As Boolean)
        '------------------------------------------------------------
        '-            Subprogram Name: TurnChange                   -
        '------------------------------------------------------------
        '-                Written By: Benjamin R. Neeb              -
        '-                Written On: April 17, 2021                -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This subroutine handles the swapping of turns between    -
        '- the players.                                             -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- btnNewTurn:  Holds which user's turn to change.          -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- btn:         Button instance to gather name and text of  -
        '-              sender object.                              -
        '------------------------------------------------------------

        'Make sure everything in the GUI is caught-up
        Application.DoEvents()

        'Pause for a couple seconds so the user can see the flipped cards
        Thread.Sleep(1500)

        'Set player turn
        blnPlayerTurn = blnNewTurn

        For Each ctrl As Control In Controls
            If TypeOf ctrl Is Button Then
                Dim btn As Button = ctrl

                'If the button is not the start or stop button, perform action
                If btn.Name <> "btnStart" And btn.Name <> "btnStop" And btn.BackColor <> Color.LightGreen Then
                    If blnPlayerTurn = blnPlayerMe Then
                        'Enable card
                        btn.Enabled = True

                        'Change back color
                        btn.BackColor = Color.LightBlue

                        'Hide text
                        btn.Text = ""
                    Else
                        'Disable card
                        btn.Enabled = False

                        'Change back color
                        btn.BackColor = Color.LightCoral

                        'Hide text
                        btn.Text = ""
                    End If
                End If
            End If
        Next
    End Sub
End Class