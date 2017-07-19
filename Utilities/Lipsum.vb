Imports System.Text

Public Class Lipsum

#Region "Singleton"

    Public Shared ReadOnly Instance As Lipsum = New Lipsum
    ''' <summary>
    ''' Prevent anyone from instantiating this class by making New private
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub New()

    End Sub

#End Region

    Private loremipsum As String() = {"Lorem ipsum dolor sit amet, consectetur adipiscing elit. Quisque ac urna est. Nulla magna dui, ullamcorper sit amet bibendum ut, ullamcorper eget lorem. Ut sem tellus, semper malesuada neque a, lacinia tempus erat. Aenean fermentum augue sed urna consequat ultricies. Pellentesque molestie porttitor tortor. Fusce porttitor nunc erat, id luctus est tincidunt ac. Quisque purus ipsum, mollis sed est sit amet, accumsan vulputate est. Vivamus quis dui tellus. ", _
                                      "Quisque aliquam nulla id tellus congue, vitae accumsan ante molestie. Etiam iaculis elit vel massa lacinia, in consequat enim sodales. Aliquam eu leo cursus, convallis tortor feugiat, tempor arcu. Donec erat purus, dictum a eleifend ut, ultricies sit amet felis. Donec venenatis ornare massa quis tempor. Etiam justo metus, porta id fermentum quis, luctus lacinia neque. Mauris vel lectus facilisis, consequat justo sit amet, eleifend nisi. Suspendisse potenti. Nam consectetur, diam nec aliquam lacinia, erat justo suscipit elit, eget tempus orci neque vel leo. Suspendisse luctus nisl a tincidunt hendrerit. Donec tellus augue, tincidunt ac dictum sed, hendrerit ac eros. Proin risus diam, rhoncus vitae tellus at, sagittis commodo dui. ", _
                                      "Proin tempus elit sed purus pellentesque, id sollicitudin mauris fermentum. Integer tincidunt tellus nec pretium luctus. Morbi ultricies felis ac gravida porttitor. Phasellus auctor nibh justo, ac vehicula nunc dapibus sollicitudin. Aliquam vitae fringilla felis. Proin et neque eu diam sagittis iaculis quis nec erat. Nam accumsan nibh ligula, vitae ultrices neque euismod ut. ", _
                                      "Sed condimentum est non feugiat viverra. In justo est, gravida non augue in, porttitor auctor nibh. Aliquam sed elit sit amet orci pellentesque laoreet eget luctus risus. Ut ultricies justo in mauris faucibus, non scelerisque felis placerat. Nunc imperdiet massa sollicitudin, auctor nibh a, sagittis ligula. Curabitur mattis, nulla euismod pellentesque eleifend, urna tellus euismod neque, ut luctus orci ipsum sit amet odio. Integer varius erat et tortor molestie rhoncus id id felis. Nulla sed eros vestibulum sem scelerisque posuere. Ut posuere sapien justo, et congue ipsum vehicula in. Mauris auctor et ligula et placerat. Quisque quam sapien, luctus at interdum at, pharetra in magna. Maecenas ut ipsum odio. Fusce luctus, sem at sodales tincidunt, tellus felis viverra ante, sit amet elementum magna purus a lorem. Praesent malesuada venenatis auctor. ", _
                                      "Sed imperdiet odio in nunc volutpat elementum. Etiam id tellus volutpat, faucibus turpis eu, viverra eros. Curabitur ultricies libero nec vulputate cursus. Vestibulum posuere pretium nibh. Sed a massa ut est molestie congue at in diam. In hac habitasse platea dictumst. Curabitur posuere mauris sed arcu fringilla, fermentum dignissim metus dignissim. Cras turpis libero, euismod id nulla pulvinar, aliquam faucibus ligula. Pellentesque ultricies purus vitae nulla venenatis cursus. Integer nisl augue, congue id dolor nec, sodales pretium dui. Phasellus sollicitudin justo vel ligula suscipit mattis. Mauris et lacus sed metus pretium laoreet. Proin mi lacus, viverra id nisl a, auctor hendrerit mi. Nulla facilisi. Morbi pellentesque diam eu orci tincidunt rhoncus. Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia Curae; "}

    Public Function Paragraphs(Optional ByVal count As Integer = 5) As String
        Dim sb As New StringBuilder
        Dim index As Integer
        For counter = 0 To count - 1
            index = counter
            If counter > 4 Then index = counter - 5
            sb.Append(loremipsum(index))
        Next
        Return sb.ToString
    End Function

End Class
