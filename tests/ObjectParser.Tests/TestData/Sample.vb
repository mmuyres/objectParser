Public Class Invoice
    Private m_total As Decimal

    Public Sub AddLine(ByVal amount As Decimal, ByVal qty As Integer)
        m_total = m_total + amount * qty
    End Sub

    Public Function Total() As Decimal
        Return m_total
    End Function

    Private Sub Reset()
        m_total = 0
    End Sub
End Class

Public Class Customer
    Public Function DisplayName(ByVal first As String) As String
        Return first.Trim()
    End Function
End Class
