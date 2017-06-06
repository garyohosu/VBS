Option Explicit

Call Main()

Sub Main()

	dim Alarm

	set Alarm = new ClsAlarm
	
	Alarm.add "","10:14","打合せ10:14　MEC第二会議室"
	Alarm.add "","10:16","打合せ10:16　MEC第二会議室"
	Alarm.start

end sub

class ClsAlarm
    Private m_list

    Private Sub Class_Initialize()
        set m_list = new ArrayList
    End Sub
	
	public sub add(xdate,xtime,xmsg)
		dim ad

		set ad = new ClsAlarmData

		if xdate<>"" then
			ad.AlarmDate = xdate
		end if

		ad.AlarmTime = xtime
		ad.message = xmsg

		m_list.add(ad)

	end sub

	public sub start
		dim item

		do while 1
			for each item in m_list.item
				if item.Disable = false then
					item.checkTime
				end if
				WScript.Sleep 100
			next
			WScript.Sleep 100
		loop

	end sub

end class

class ClsAlarmData

    Private m_message
    Private m_AlarmTime
	Private m_AlarmDate
	Private m_Disable

    Private Sub Class_Initialize()
        m_Disable = false
    End Sub


    Public Property Get Message
        Message = m_message
    End Property

    Public Property Let Message(vData)
        m_message = vData
    End Property

    Public Property Get AlarmTime
        AlarmTime = m_AlarmTime
    End Property

    Public Property Let AlarmTime(vData)
        m_AlarmTime = vData
    End Property

    Public Property Get AlarmDate
        AlarmDate = m_AlarmDate
    End Property

    Public Property Let AlarmDate(vData)
        m_AlarmDate = vData
    End Property

    Public Property Get Disable
        Disable = m_Disable
    End Property

    Public Property Let Disable(vData)
        m_Disable = vData
    End Property

	public sub checkTime
		dim strDate
		dim strTime

		'strDateに本日日付が yyyy/mm/dd 形式でセットされます。
		strDate = FormatDateTime(Now, 1)

		'strDateに現在の時刻が hh:mm 形式でセットされます。
		strTime = FormatDateTime(Now, 4)

		if m_AlarmDate="" or m_AlarmDate = strDate then '日付指定なし　または本日
			if strTime >= m_AlarmTime then
				msgbox m_message, vbSystemModal
				m_Disable = True
			end if
		end if
	end sub

end class

'動的配列版ArrayList
class ArrayList

	private m_Item()
	private m_count

	public sub Add(x)
		ReDim Preserve m_item(m_count)
		If IsObject(x) Then
			set m_item(m_count) = x
		else
			m_item(m_count) = x
		end if
		m_count = m_count + 1
	end sub

	public sub Change(i,x)
		If IsObject(x) Then
			set m_item(i) = x
		else
			m_item(i) = x
		end if
	end sub

	public function Count
		Count = m_count
	end function

	public function Clear
		m_count=0
		Erase m_item
	end function

	public function Item
		Item = m_Item
	end function

	public function Items(n)
		If IsObject(m_Item(n)) Then
			set Items = m_Item(n)
		else
			Items = m_Item(n)
		end if
	end function

end class


