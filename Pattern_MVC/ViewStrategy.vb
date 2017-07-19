''' <summary>
''' The ViewStrategy class is composed of two holders for Behaviours that must be set before they can be used
''' </summary>
''' <remarks>Inheritors must assign specific implementations to Behaviour A and Behaviour B</remarks>
Public MustInherit Class ViewStrategy
    Inherits System.Windows.Forms.Form
    'We need a holder for each Behaviour
    Private MyBehaviourA As IBehaviourA
    ''' <summary>
    ''' Set which BehaviourA to use
    ''' </summary>
    ''' <param name="NewBehaviour">IBehaviourA</param>
    ''' <remarks></remarks>
    Protected Sub SetBehaviourA(ByVal NewBehaviour As IBehaviourA)
        MyBehaviourA = NewBehaviour
    End Sub
    ''' <summary>
    ''' Execute the Behaviours
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub RunMyBehaviours()
        Try
            MyBehaviourA.UseBehaviour()
        Catch ex As Exception
            Throw New Exception("You must set Behaviours A and B before running")
        End Try
    End Sub
End Class
