Option Explicit

Call Main()




Sub Main()

	' ArrayListÇÃçÏê¨
	Dim myArrayList
	Set myArrayList = CreateObject("System.Collections.ArrayList")

    Dim MSerial1
    Set MSerial1 = New Serial

    MSerial1.Header = "AM1"
    MSerial1.Size = 11

    MSerial1.SerialNumber = "AM200000001"
    myArrayList.add MSerial1

    Dim MSerial2
    Set MSerial2 = New Serial

    MSerial2.Header = "AM1"
    MSerial2.Size = 11

    MSerial2.SerialNumber = "AM10000002"
    myArrayList.add MSerial2

    Dim MSerial3
    Set MSerial3 = New Serial

    MSerial3.Header = "AM1"
    MSerial3.Size = 11

    MSerial3.SerialNumber = "AM100000003"
    myArrayList.add MSerial3

    dim elem

    For Each elem In myArrayList
        msgbox(elem.SerialNumber)
        msgbox(elem.check)
    next

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
