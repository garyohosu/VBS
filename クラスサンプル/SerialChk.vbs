Option Explicit

Call Main()

Sub Main()
    Dim MSerial
    Set MSerial = New Serial

    MSerial.Header = "AM1"
    MSerial.Size = 11

    MSerial.SerialNumber = "AM200000001"
    msgbox(MSerial.check)

    MSerial.SerialNumber = "AM10000001"
    msgbox(MSerial.check)

    MSerial.SerialNumber = "AM100000001"
    msgbox(MSerial.check)

End Sub

Class Serial

    Private m_SerialNumber
    Private m_Header
    Private m_Size

    Public Property Get SerialNumber
        SerialNumber = m_SerialNumber
    End Property

    Public Property Let SerialNumber(vData)
        m_SerialNumber = vData
    End Property

    Public Property Get Header
        Header = m_Header
    End Property

    Public Property Let Header(vData)
        m_Header = vData
    End Property

    Public Property Get Size
        Size = m_Size
    End Property

    Public Property Let Size(vData)
        m_Size = vData
    End Property

    Public Function check
        if m_Header = left(m_SerialNumber,len(m_Header)) and m_Size = len(m_SerialNumber) then
            check = true
        else
            check = false
        end if
    End Function

    Private Sub Class_Initialize()
        m_Header = "8SK"
        m_Size = 11
    End Sub

    Private Sub Class_Terminate()
        MsgBox ("èIóπ")
    End Sub

End Class
