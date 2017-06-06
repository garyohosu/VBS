Option Explicit

dim al

set al = new ArrayList

al.Add("test1")
al.Add("test2")
al.Add("test3")

al.Insert 1,"test Ins"

dim i

msgbox(al.Count)

for i = 0 to al.Count - 1
	msgbox(al.Item(i))
next

al.Clear

dim t1 

set t1 = new ClsTest
t1.mes = "Test2_1"

al.Add(t1)

dim t2 

set t2 = new ClsTest
t2.mes = "Test2_2"


al.Add(t2)

dim t3 

set t3 = new ClsTest
t3.mes = "Test2_3"

al.Add(t3)

for i = 0 to al.Count - 1
	al.Item(i).say
next


Class ClsTest
	public mes
	public sub say
		msgbox(mes)
	end sub
end Class



Class ArrayList

	public nextData

	public data

	private m_size

	private sub class_initialize()
		me.nextData = Null
	end sub

    Private Sub Class_Terminate()
        'MsgBox (data & "‰ğ•ú")
    End Sub

	public function Clear

		dim item
		dim cnt
		dim prev

		do

			set item = me
			set prev = me

			cnt = 0

			do while not isNull(item.nextData)
				cnt = cnt + 1
				set prev = item
				set item = item.nextData
			loop
			if cnt > 0 then
				set item = nothing
				prev.nextData = Null
			end if
		loop while cnt > 0
		me.nextData = Null
		m_size = 0
	end function



	public sub Add(x)

		dim item
		dim newItem

		set newItem = new ArrayList

		If IsObject(x) Then
			set newItem.data = x
		else
			newItem.data = x
		end if

		newItem.nextData = null

		set item = me

		do while not isNull(item.nextData)
			set item = item.nextData
		loop

		set item.nextData = newItem
		
		m_size = m_size + 1

	end sub

	public function Count

		Count = m_size

	end function

	public function Item(n)
		dim I
		dim itm

		set itm = me

		for I = 0 to n
			if not isNull(itm.nextData) then
				set itm = itm.nextData
			end if
		next
		
		If IsObject(itm.data) Then
			set item = itm.data
		else		
			Item = itm.data
		end if
	end function

	public sub Insert(n,x)

		dim I
		dim itm
		dim insData
		dim prev

		if n >= m_size then
			msgbox("ˆø”‚ª‘½‚«‚·‚¬‚Ü‚·")
			exit sub
		end if

		set insData = new ArrayList

		If IsObject(x) Then
			set insData.data = x
		else		
			insData.data = x
		end if

		set itm = me
		set prev = me

		for I = 0 to n
			if not isNull(itm.nextData) then
				set prev = itm
				set itm = itm.nextData
			end if
		next

		set prev.nextData = insData
		set insData.nextData = itm
		m_size = m_size + 1

	end sub

end class
