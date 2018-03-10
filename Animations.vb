Class Transition_Animation
    Private StartTime As Date
    Private AddPlayTime As Integer
    Private AddPosition As Single
    Sub New()
    End Sub
    Sub New(ByVal Length%)
        _Length = Length
    End Sub
    Sub New(ByVal Length%, ByVal EaseMode As Easing, ByVal Transition As Object)
        _Length = Length
        Me.EaseMode = EaseMode
        Me.Transition = Transition
    End Sub
#Region " Playing "
    Public Sub Start()
        AddPlayTime = PlayTime
        StartTime = Now
    End Sub

    Public Sub Pause()
        AddPlayTime = PlayTime
        StartTime = Nothing
    End Sub

    Public Sub Reset()
        StartTime = Nothing
        AddPlayTime = 0

        PlayForward = True
        AddPosition = 0
    End Sub

    Public Sub ToggleRunning()
        IsRunning = Not IsRunning
    End Sub

    Public Sub ToggleForward()
        PlayForward = Not PlayForward
    End Sub
#End Region
#Region " Properties "
    Private _Length = 1000
    Public ReadOnly Property Length As Integer
        Get
            Return _Length
        End Get
    End Property

    Public ReadOnly Property PlayTime() As Integer
        Get
            Return If(IsRunning, AddPlayTime + (Now - StartTime).TotalMilliseconds, AddPlayTime)
        End Get
    End Property

    Public Property Position As Single
        Get
            Return Math.Max(0, Math.Min(1, GetPosition() / Length))
        End Get
        Set(ByVal value As Single)
            AddPosition = If(PlayForward, PlayTime - Length * value, (PlayTime + Length * value) / 2)
        End Set
    End Property

    Private Function GetPosition() As Single
        Return If(PlayForward, PlayTime - AddPosition, AddPosition * 2 - PlayTime)
    End Function

    Public Property IsRunning As Boolean
        Get
            Return StartTime <> Nothing
        End Get
        Set(ByVal value As Boolean)
            If value Then Start() Else Pause()
        End Set
    End Property

    Private _PlayForward As Boolean = True
    Public Property PlayForward As Boolean
        Get
            Return _PlayForward
        End Get
        Set(ByVal value As Boolean)
            If value = _PlayForward Then Exit Property
            AddPosition = If(value, PlayTime - If(Position > 0, GetPosition, 0), If(Position = 1, (PlayTime + Length) / 2, GetPosition() + AddPosition / 2))
            _PlayForward = value
        End Set
    End Property

    Public Property StartValue As Single = 0
    Public Property EndValue As Single = 1
    Public Property EaseMode As Easing = 0
    Public Property Transition As Calc = Transitions.Linear

    Public ReadOnly Property Value As Single
        Get
            Return Value(Position)
        End Get
    End Property

    Public ReadOnly Property Value(ByVal Position As Single) As Single
        Get
            Dim p = Position

            Select Case EaseMode
                Case Is = Easing.EaseIn : Value = Transition.Invoke(p)
                Case Is = Easing.EaseOut : Value = 1 - Transition.Invoke(1 - p)
                Case Is = Easing.EaseInOut : Value = If(p <= 0.5, Transition.Invoke(2 * p) / 2, (2 - Transition.Invoke(2 * (1 - p))) / 2)
            End Select

            Return StartValue + Value * EndValue
        End Get
    End Property
#End Region
    Public Enum Easing
        EaseIn
        EaseOut
        EaseInOut
    End Enum
    Public Delegate Function Calc(ByVal x!) As Single
    Public NotInheritable Class Transitions

        Public Shared ReadOnly Linear = New Calc(Function(x!) x)
        Public Shared ReadOnly Quadrantic = New Calc(Function(x!) Math.Pow(x, 2))
        Public Shared ReadOnly Quintic = New Calc(Function(x!) Math.Pow(x, 5))
        Public Shared ReadOnly Circular = New Calc(Function(x!) 1 - Math.Sin(Math.Acos(x)))
        Public Shared ReadOnly Jumpback = New Calc(Function(x!) Math.Pow(x, 2) * ((1.5 + 1) * x - 1.5))
        Public Shared ReadOnly Elastic = New Calc(Function(x!) Math.Pow(2, 10 * (x - 1)) * Math.Cos(20 * Math.PI * 1.5 / 3 * x))
        Public Shared ReadOnly Bounce = New Calc(Function(x!)
                                                     Dim a As Double = 0, b As Double = 1

                                                     Do
                                                         If x >= (7 - 4 * a) / 11 Then Return -Math.Pow((11 - 6 * a - 11 * x) / 4, 2) + Math.Pow(b, 2)
                                                         a += b : b /= 2
                                                     Loop
                                                 End Function)
    End Class
End Class
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'
'Transition Class By BlackCap 
'http://www.hackforums.net/member.php?action=profile&uid=316828
'All Credits Go To Him <3

Class Shake_Animation



    Private frm As Control
    Private howMuch As Integer
    Private th As Threading.Thread
    Private timeForPause As Integer
    Public Shared MOVE_UP_DOWN As Integer = 0
    Public Shared MOVE_LEFT_RIGHT As Integer = 1

    Private direction As Integer

    Public Property timeForPauseBetweenMove As Integer
        Get
            Return timeForPause
        End Get
        Set(ByVal value As Integer)
            timeForPause = value
        End Set
    End Property

    Sub New(ByVal frm As Control, ByVal howMuchToReapeR As Integer)
        Me.frm = frm
        howMuch = howMuchToReapeR
    End Sub

    Public Property moveReapeRing As Integer
        Get
            Return direction
        End Get
        Set(ByVal value As Integer)
            direction = value
        End Set
    End Property

    Sub shake()

        Dim tempLoc As Point = frm.Location
        Dim startLoc As Point = New Point(frm.Location.X, frm.Location.Y)

        Select Case moveReapeRing

            Case MOVE_LEFT_RIGHT

                For a As Integer = howMuch To 0 Step -1

                    Dim poss As New ff(AddressOf formPosition)
                    frm.Invoke(poss, New Point(startLoc.X - a, startLoc.Y))
                    Threading.Thread.Sleep(timeForPauseBetweenMove)

                    frm.Invoke(poss, New Point(startLoc.X + a, startLoc.Y))
                    Threading.Thread.Sleep(timeForPauseBetweenMove)

                Next

            Case MOVE_UP_DOWN

                For a As Integer = howMuch To 0 Step -1

                    Dim poss As New ff(AddressOf formPosition)
                    frm.Invoke(poss, New Point(startLoc.X, startLoc.Y - a))
                    Threading.Thread.Sleep(timeForPauseBetweenMove)

                    frm.Invoke(poss, New Point(startLoc.X, startLoc.Y + a))
                    Threading.Thread.Sleep(timeForPauseBetweenMove)

                Next

        End Select

        Dim pos As New ff(AddressOf formPosition)
        frm.Invoke(pos, startLoc)

    End Sub

    Public Delegate Sub ff(ByVal p As Point)

    Sub formPosition(ByVal p As Point)
        frm.Location = p
    End Sub

    Sub startShake()

        Try
            th = New System.Threading.Thread(AddressOf shake)
            th.Start()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class
'I Am Not Sure Who Created This Part Of Code
'I Had Found It On The Internet Long Time Ago!
'I Am Really Sorry!


Module Animations
    Public WithEvents transition_location_timer As New Timer
    Public transition_animator As New Transition_Animation(2000, Transition_Animation.Easing.EaseInOut, Transition_Animation.Transitions.Jumpback)
    Public s_control As Control
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
    Private Sub info_loc_changer() Handles transition_location_timer.Tick
        s_control.Location = New Point(transition_animator.Value, s_control.Location.Y)
        If transition_animator.Value = 1 Then
            transition_location_timer.Stop()
        End If
    End Sub

 
End Module
