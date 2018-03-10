Imports System.Security
Imports System.Security.Cryptography
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Threading
Imports Transitions

Public Class betsayda

#Region "Animations"
    Dim linear As New Transitions.TransitionType_Linear(1000)
    Dim tlinear As New Transitions.Transition(linear)
    Dim EiEo As New Transitions.TransitionType_EaseInEaseOut(1000)
    Dim tEiEo As New Transitions.Transition(EiEo)

   
    Public Sub Shake(ByVal control As Control, ByVal speed As Integer, ByVal repeat As Integer, ByVal type As Integer)
        Dim Shaker As New Shake_Animation(control, repeat)
        If type = 1 Then
            Shaker.moveReapeRing = Shake_Animation.MOVE_LEFT_RIGHT
        Else
            Shaker.moveReapeRing = Shake_Animation.MOVE_UP_DOWN
        End If
        Shaker.timeForPauseBetweenMove = speed
        Shaker.startShake()
    End Sub
    Public Sub Forward(ByRef control As Control, ByRef start_values As Integer, ByRef end_values As Integer)
        transition_animator.Reset()
        transition_animator.StartValue = start_values
        transition_animator.EndValue = end_values
        s_control = control
        transition_animator.Start()
        transition_location_timer.Interval = 1
        transition_location_timer.Start()
    End Sub
    Public Sub Back()
        transition_animator.ToggleForward()
        transition_animator.Start()
        transition_location_timer.Interval = 1
        transition_location_timer.Start()
    End Sub

    Public Function IsEnterInsideBox(ByVal target As Control, ByVal mover As Control) As Boolean
        Dim bool As Boolean = Nothing
        Dim x As Integer
        Dim y As Integer
        Dim xr As Integer = target.Location.X
        Dim yr As Integer = target.Location.Y
        Dim xb As Integer = mover.Location.X
        Dim yb As Integer = mover.Location.Y
        Dim hb As Integer = mover.Size.Height
        Dim wb As Integer = mover.Size.Width
        Dim hr As Integer = target.Size.Height
        Dim wr As Integer = target.Size.Width
        x = xr + Val(wr) - wb
        y = yr + Val(hr) - hb
        If xb >= xr And xb <= x And yb >= yr And yb <= y Then
            Return True
        Else
            Return False
        End If
        Return bool
    End Function

    Public Function IsEnterBox(ByVal target As Control, ByVal mover As Control) As Boolean
        Dim bool As Boolean = Nothing
        Dim x As Double
        Dim y As Double
        Dim r As Double
        Dim t As Double
        Dim xr As Integer = target.Location.X
        Dim yr As Integer = target.Location.Y
        Dim xb As Integer = mover.Location.X
        Dim yb As Integer = mover.Location.Y
        Dim hb As Integer = mover.Size.Height
        Dim wb As Integer = mover.Size.Width
        Dim hr As Integer = target.Size.Height
        Dim wr As Integer = target.Size.Width
        x = xr + Val(wr) - wb
        y = yr + Val(hr) - hb
        x = x + wb
        y = y + hb
        If xb >= (xr - wb) And xb <= x And yb >= (yr - hb) And yb <= y Then
            Return True
        Else
            Return False
        End If
        Return bool
    End Function

    Function GetLocation(ByVal e As Control) As Point
        Dim i As Point
        i.X = Cursor.Position.X - e.Location.X
        i.Y = Cursor.Position.Y - e.Location.Y
        Return i
    End Function


    Function SetLocation(ByVal i As Point, ByVal e As System.Drawing.Point) As Point
        Dim p As New Point
        p = e
        p.X = p.X - i.X
        p.Y = p.Y - i.Y
        Return p
    End Function

    Public Sub LabelTransition(ByVal lbl As Control, ByVal str As String, ByVal col As Color)
        If lbl.Text = str Then
            Exit Sub
        Else
            With tlinear
                .add(lbl, "Text", str)
                .add(lbl, "ForeColor", col)
                Try
                    .run()
                Catch
                End Try
            End With
        End If
    End Sub

    Sub SwapObject(ByVal r As Control, ByVal l As Control, ByVal _LEFT As Integer)
        Dim ctrlOnScreen As Control
        Dim ctrlOffScreen As Control
        If l.Left = _LEFT Then
            ctrlOnScreen = l
            ctrlOffScreen = r
        Else
            ctrlOnScreen = r
            ctrlOffScreen = l
        End If
        ctrlOnScreen.SendToBack()
        ctrlOffScreen.BringToFront()
        tEiEo.add(ctrlOnScreen, "Left", -1 * ctrlOnScreen.Width)
        tEiEo.add(ctrlOffScreen, "Left", _LEFT)

        Try
            tEiEo.run()
        Catch
        End Try
    End Sub

    Sub ReSize(ByVal e As Control, ByVal type As String, ByVal w As Integer)
        tEiEo.add(e, type, w)
        Try
            tEiEo.run()
        Catch
        End Try
    End Sub

    Sub ChangeColor(ByVal e As Control, ByVal type As String, ByVal col As Color)
        tEiEo.add(e, type, col)
        Try
            tEiEo.run()
        Catch
        End Try
    End Sub
    'The Flash transition animates the property to the destination value
    'and back again. You specify how many flashes to show and the length
    'of each flash...
    Sub FlashColor(ByVal e As Control, ByVal type As String, ByVal col As Color, ByVal i As Integer, ByVal l As Integer)
        Dim flash As New Transitions.TransitionType_Flash(i, l)
        Dim tflash As New Transitions.Transition(flash)
        tflash.add(e, type, col)
        Try
            tflash.run()
        Catch
        End Try
    End Sub

    ' e is the size of destination, i is size of 
    Sub BounceBack(ByVal lenght As Control, ByVal target As Control, ByVal type As String, ByVal ctr As Integer)
        Dim b As New Transitions.TransitionType_Bounce(ctr)
        Dim tbounce As New Transitions.Transition(b)
        Dim i As Integer
        If type = "Top" Then
            i = lenght.Height - target.Height
        ElseIf type = "Left" Then
            i = lenght.Width - target.Width
        Else
            Exit Sub
        End If
        tbounce.add(target, type, i)
        Try
            tbounce.run()
        Catch
        End Try
    End Sub

    Sub ThrowBack(ByVal lenght As Control, ByVal target As Control, ByVal type As String, ByVal ctr As Integer)
        Dim tc As New Transitions.TransitionType_ThrowAndCatch(ctr)
        Dim ttc As New Transitions.Transition(tc)
        ttc.add(target, type, 0)
        Try
            ttc.run()
        Catch
        End Try
    End Sub

    Sub DropMe(ByVal lenght As Control, ByVal target As Control, ByVal type As String, ByVal ctr As Integer)

        Dim drop As New Transitions.TransitionType_Drop(ctr)
        Dim tdrop As New Transitions.Transition(drop)
        Dim i As Integer
        If type = "Top" Then
            i = lenght.Height - target.Height
        ElseIf type = "Left" Then
            i = lenght.Width - target.Width
        Else
            Exit Sub
        End If
        tdrop.add(target, type, i)
        Try
            tdrop.run()
        Catch
        End Try

    End Sub

    'Sub DropIt(ByVal lenght As Control, ByVal target As Control)

    'Dim elements As IList(Of TransitionElement) = New List(Of TransitionElement) From {New TransitionElement(40, 100, 1), New TransitionElement(65, 70, 2), New TransitionElement(80, 100, 1), New TransitionElement(90, 92, 2), New TransitionElement(100, 100, 1)}
    ' Dim t As New Transitions.Transition(elements)
    'Dim destinationValue As Integer = ((lenght.Height - target.Height) - 10)
    '    Transitions.Transition.run(target, "Top", destinationValue, New TransitionType_UserDefined(elements, &H7D0))
    ' Dim transition As New Transition(New TransitionType_Linear(&H7D0))
    '     transition.add(target, "Left", (target.Left + 400))
    ' Dim transition2 As New Transition(New TransitionType_EaseInEaseOut(&H7D0))
    '   transition2.add(target, "Top", &H13)
    '   transition2.add(target, "Left", 6)
    '   transition.runChain(New Transition() {transition, transition2})


    'End Sub
    Dim bool_0 As Boolean = False
    Dim int_0 As Integer
    Dim int_1 As Integer
    Dim int_2 As Integer
    Dim int_3 As Integer

    Public Sub ZoomIn(ByVal pic As Control, ByVal by As Integer)
        If Not bool_0 = False Then
            Dim num5 As Integer
            Dim num6 As Integer
            Do While (num5 = num6)
                num6 = 1
                Dim num7 As Integer = num5
                num5 = 1
                If (1 > num7) Then
                    Return
                End If
            Loop
        Else
            int_0 = pic.Height
            int_1 = pic.Width
            Dim num3 As Integer = Convert.ToInt32(CDbl((Math.Round(CDbl(((by * 0.01) * int_0))) * 0.5)))
            Dim num4 As Integer = Convert.ToInt32(CDbl((Math.Round(CDbl(((by * 0.01) * int_1))) * 0.5)))
            Dim num As Integer = (int_0 + (num3 * 2))
            Dim num2 As Integer = (int_1 + (num4 * 2))
            int_2 = num3
            int_3 = num4
            pic.Width = num2
            pic.Height = num
            pic.Top = (pic.Top - int_2)
            pic.Left = (pic.Left - int_3)
            bool_0 = False
        End If
    End Sub

    Public Sub ZoomOut(ByVal pic As Control)
        If bool_0 = True Then
            Dim num As Integer
            Dim num2 As Integer
            Do While (num = num2)
                num2 = 1
                Dim num3 As Integer = num
                num = 1
                If (1 > num3) Then
                    Return
                End If
            Loop
        Else
            pic.SuspendLayout()
            pic.Width = int_1
            pic.Left = (pic.Left + int_3)
            pic.Height = int_0
            pic.Top = (pic.Top + int_2)
            pic.ResumeLayout()
            bool_0 = False
        End If
    End Sub


#End Region

#Region "EncDec"
    Public Function Encrypt(ByVal plainText As String, ByVal passPhrase As String) As String
        Try
            Dim saltValue As String = "mySaltValue"
            Dim hashAlgorithm As String = "SHA1"
            Dim passwordIterations As Integer = 2
            Dim initVector As String = "@1B2c3D4e5F6g7H8"
            Dim keySize As Integer = 256

            Dim initVectorBytes As Byte() = Encoding.ASCII.GetBytes(initVector)
            Dim saltValueBytes As Byte() = Encoding.ASCII.GetBytes(saltValue)

            Dim plainTextBytes As Byte() = Encoding.UTF8.GetBytes(plainText)


            Dim password As New PasswordDeriveBytes(passPhrase, saltValueBytes, hashAlgorithm, passwordIterations)

            Dim keyBytes As Byte() = password.GetBytes(keySize \ 8)
            Dim symmetricKey As New RijndaelManaged()

            symmetricKey.Mode = CipherMode.CBC

            Dim encryptor As ICryptoTransform = symmetricKey.CreateEncryptor(keyBytes, initVectorBytes)

            Dim memoryStream As New MemoryStream()
            Dim cryptoStream As New CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write)

            cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length)
            cryptoStream.FlushFinalBlock()
            Dim cipherTextBytes As Byte() = memoryStream.ToArray()
            memoryStream.Close()
            cryptoStream.Close()
            Dim cipherText As String = Convert.ToBase64String(cipherTextBytes)
            Return cipherText
        Catch
            MsgBox("Rename passphase to kitchanismo")
        End Try
    End Function


    Public Function Decrypt(ByVal cipherText As String, ByVal passPhrase As String) As String
        Try
            Dim saltValue As String = "mySaltValue"
            Dim hashAlgorithm As String = "SHA1"

            Dim passwordIterations As Integer = 2
            Dim initVector As String = "@1B2c3D4e5F6g7H8"
            Dim keySize As Integer = 256

            Dim initVectorBytes As Byte() = Encoding.ASCII.GetBytes(initVector)
            Dim saltValueBytes As Byte() = Encoding.ASCII.GetBytes(saltValue)
            Dim cipherTextBytes As Byte() = Convert.FromBase64String(cipherText)
            Dim password As New PasswordDeriveBytes(passPhrase, saltValueBytes, hashAlgorithm, passwordIterations)


            Dim keyBytes As Byte() = password.GetBytes(keySize \ 8)
            Dim symmetricKey As New RijndaelManaged()
            symmetricKey.Mode = CipherMode.CBC
            Dim decryptor As ICryptoTransform = symmetricKey.CreateDecryptor(keyBytes, initVectorBytes)
            Dim memoryStream As New MemoryStream(cipherTextBytes)
            Dim cryptoStream As New CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read)
            Dim plainTextBytes As Byte() = New Byte(cipherTextBytes.Length - 1) {}
            Dim decryptedByteCount As Integer = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length)

            memoryStream.Close()
            cryptoStream.Close()
            Dim plainText As String = Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount)
            Return plainText
        Catch
            MsgBox("Rename passphase to kitchanismo")
        End Try
    End Function
#End Region

#Region "FormFade"
    Public Sub FormFade(ByVal e As System.Object, ByVal isleep As Integer)
        e.Opacity = 0.99
        Dim num As Integer = 100 '&H62
        Do
            e.Opacity = (CDbl(num) / 100)
            Thread.Sleep(isleep)
            num = (num + -2)
        Loop While (num >= 8)
    End Sub
#End Region

#Region "ProtectApp"
    Public Sub ProtectApp(ByVal procname As String, ByVal x As Object)
        Dim TargetProcess As Process() = Process.GetProcessesByName(procname)
        If TargetProcess.Length = 0 Then
            Interaction.MsgBox("Please Do NOT Rename this Application!", MsgBoxStyle.Exclamation, Nothing)
            x.close()
        End If
    End Sub

#End Region

#Region "GenerateWord"
    Public WordList As New List(Of String)
    Private RandomClass As New Random()
    Public Sub ChooseWord(ByVal Word As Label, ByVal WordLength As Integer, ByVal DictionaryWords As String)
        Try
            Dim TextFile As New StreamReader(DictionaryWords)
            While TextFile.EndOfStream = False
                WordList.Add(TextFile.ReadLine())
            End While

            For Counter As Integer = WordList.Count - 1 To 0 Step -1
                If WordList(Counter).Length <> WordLength Then
                    WordList.RemoveAt(Counter)
                End If
            Next
            Word.Text = WordList(RandomClass.Next(0, WordList.Count))
        Catch ex As Exception
            MessageBox.Show("Error:" & ControlChars.NewLine & ex.Message, " - Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
        End Try
    End Sub


    Public Function ChooseWordRandomNum(ByVal DictionaryWords As String) As String
        Dim _word As String
        Dim ioFile As New StreamReader(DictionaryWords)
        Dim lines As New List(Of String)
        Dim rnd As New Random()
        Dim line As Integer
        While ioFile.Peek <> -1
            lines.Add(ioFile.ReadLine())
        End While
        line = rnd.Next(lines.Count + 1)
        _word = (lines(line).Trim())
        ioFile.Close()
        ioFile.Dispose()
        Return _word
    End Function

    Public Function ScrambleWord(ByVal s As String) As String
        Dim rndm As New Random
        Dim sb As New StringBuilder
        Dim chrWord As List(Of Char) = s.ToList
        For i As Integer = 0 To chrWord.Count - 1
            Dim temp As Integer = rndm.Next(0, chrWord.Count)
            sb.Append(chrWord(temp))
            chrWord.RemoveAt(temp)
        Next
        Return sb.ToString
    End Function


#End Region

#Region "FindButtonName"
    Private AllFormButtons As New List(Of Button)
    Public Function FindButtonByName(ByVal buttonName As String) As Button
        If AllFormButtons.Exists(Function(b) String.Compare(b.Name, buttonName, True) = 0) Then
            Return AllFormButtons.Where(Function(b) String.Compare(b.Name, buttonName, True) = 0).First()
        Else
            Return Nothing
        End If
    End Function
    Public Sub GatherControls(ByVal parentControl As Control)
        If TypeOf parentControl Is Button Then
            AllFormButtons.Add(DirectCast(parentControl, Button))
        End If
        For Each ctrl As Control In parentControl.Controls
            GatherControls(ctrl)
        Next
    End Sub
#End Region

#Region "RandomNumber"
    Public Function RanMinMax(ByVal min As Integer, ByVal max As Integer) As Integer
        Dim res As Integer
        Dim rnd As New Random
        res = rnd.Next(min, max + 1)
        Return res
    End Function

    Public Sub DisplayRandomInListbox(ByVal picLength As Integer)
        Dim allNumbers As New List(Of Integer)
        Dim randomNumbers As New List(Of Integer)
        Dim rand As New Random
        For i As Integer = 1 To picLength Step 1
            allNumbers.Add(i)
        Next i
        For i As Integer = 1 To picLength Step 1
            Dim selectedIndex As Integer = rand.Next(0, (allNumbers.Count))
            Dim selectedNumber As Integer = allNumbers(selectedIndex)
            randomNumbers.Add(selectedNumber)
            allNumbers.Remove(selectedNumber)
            lbran.Items.Add(selectedNumber)
        Next i
    End Sub
    Dim lbran As New ListBox
    Dim ctr As Integer = 0
    Dim bool As Boolean = True
    Public Function GetRandomItemInListbox(ByVal lb As ListBox) As String
        Dim cnt As Integer = lb.Items.Count
        Dim str As String
        Dim i As Integer
        If cnt = 0 Then
            MsgBox("Listbox has O item. Add an item.")
            Return "None"
            Exit Function
        End If
        DisplayRandomInListbox(cnt)
        If ctr = cnt Then
            lbran.Items.Clear()
            DisplayRandomInListbox(cnt)
            ctr = 0
        End If
        i = lbran.Items.Item(ctr) - 1
        str = lb.Items.Item(i).ToString
        ctr += 1
        Return str
    End Function
#End Region



End Class






