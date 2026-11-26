<%
Class XoiseRecordset
	Private m_Sql
	Private m_Array
	Private m_Array_Length
	Private m_Index
	Private m_Column_Array
	Private m_Conn
	Private m_Page_Size
	Private m_Absolute_Page
	Private m_Pages
	Private m_Max_Record

	Public  Property Get Item(strName)
		col_index = GetItemIndex(strName)
		Item = m_Array(col_index, m_Index)
	End Property

	Public   sub Itemm(strName,uu)
			col_index = GetItemIndex(strName)
			 m_Array(col_index, m_Index)=uu
	End sub


	Public Property Let SQL(ByVal strSQL)
		m_Sql = strSQL
		Call Open()
	End Property

	Public Property Let Conn(ByRef objConn)
		Set m_Conn = objConn
		Call Open()
	End Property

	Public Property Get EOF()
		If m_Index >= m_Max_Record Or m_Array_Length < 0 Then
			EOF = True
		Else
			If m_Pages Then
				If m_Index < RecordCount() Then
					EOF = False
				Else
					EOF = True
				End If
			Else
				EOF = False
			End If
		End If
	End Property

	Public Property Get Row(intIndex)
		Row = m_Array(intIndex - 1)
	End Property

	Public Property Get BOF()
		BOF = False
	End Property

	Public Property Get GetRows()
		GetRows = m_Array
	End Property

	Public Property Get AbsolutePosition()
		AbsolutePosition = m_Index + 1
	End Property

	Public Property Let AbsolutePosition(intIndex)
		m_Index = intIndex - 1
	End Property

	Public Property Get AbsolutePage()
		AbsolutePage = m_Absolute_Page
	End Property

	Public Property Let AbsolutePage(intIndex)
		m_Pages = True
		m_Absolute_Page = intIndex
		m_Index = m_Page_Size * m_Absolute_Page - m_Page_Size
		m_Max_Record = m_Index + m_Page_Size
	End Property

	Public Property Get PageSize()
		PageSize = m_Page_Size
	End Property

	Public Property Let PageSize(intPageSize)
		m_Pages = True
		m_Page_Size = intPageSize
		m_Index = m_Page_Size * m_Absolute_Page - m_Page_Size
		m_Max_Record = m_Index + m_Page_Size
	End Property

	Public Property Get PageCount()
		cnt = Int(RecordCount() / m_Page_Size)
		If cnt < RecordCount() / m_Page_Size Then
			cnt = cnt + 1
		End If
		PageCount = cnt
	End Property

	Public Property Get RecordCount()
		RecordCount = m_Array_Length + 1
	End Property

	Public Sub Move(intPos)
		intPos = intPos - 1
		If IsNumeric(intPos) Then
			If intPos >= 0 And intPos < m_Array_Length Then
				m_Index = intPos
			End If
		End If
	End Sub

	Public Sub MoveNext()
		m_Index = m_Index + 1
	End Sub

	Public Sub MovePrevious()
		m_Index = m_Index - 1
	End Sub

	Public Sub MoveFirst()
		m_Index = m_Page_Size * m_Absolute_Page - m_Page_Size
	End Sub

	Public Sub MoveLast()
		If m_Pages Then
			m_Index = (m_Page_Size * m_Absolute_Page - m_Page_Size) + m_Page_Size - 1
		Else
			m_Index = m_Array_Length
		End If
	End Sub

	Public Sub Open()
		If IsObject(m_Conn) And m_sql <> "" Then
			m_Column_String = ReplaceRegExp(m_Sql, "SELECT ([^\[]*) FROM(.*)", "$1")
			m_Column_Array = Split(m_Column_String, ",")
			If IsArray(m_Column_Array) Then
				For i = 0 To UBound(m_Column_Array)
					m_Column_Array(i) = Trim(Replace(Replace(ReplaceRegExp(m_Column_Array(i), "([^\[]*) AS ([^\[]*)", "$2"), "[", ""), "]", ""))
				Next
			End If

			m_Array = GetRSArray()

			If IsArray(m_Array) Then
				m_Array_Length = UBound(m_Array, 2)
				If Not m_Pages Then
					m_Max_Record = m_Array_Length + 1
				End If

			End If
		End If

	End Sub

	Private Sub Class_Initialize()
		m_Index = 0
		m_Array_Length = -1
		m_Absolute_Page = 1
		m_Page_Size = 0 ' 0 = Unlimited
		m_Pages = False
	End Sub

	Private Sub Class_Terminate()
		If IsArray(m_Array) Then
			Erase m_Array
		End If
		If IsObject(m_Conn) Then
			Set m_Conn = Nothing
		End If
	End Sub

	Private Function GetItemIndex(strName)
		For i = 0 To UBound(m_Column_Array)
			If CStr(Trim(strName)) = CStr(m_Column_Array(i)) Then
				GetItemIndex = i
				Exit For
			End If
		Next
	End Function



	Private Function GetRSArray()
		Set m_RS = m_Conn.Execute(m_Sql)
		If Not m_RS.EOF Then
			GetRSArray = m_RS.GetRows()
		Else
			GetRSArray = False
		End If
		Set m_RS = Nothing
	End Function

	Public Function ReplaceRegExp(ByVal strString, strPattern, strReplace)
		Dim RE: Set RE = New RegExp

		With RE
			.Pattern = strPattern
			.Global = True
			.IgnoreCase = True
			ReplaceRegExp = .Replace(strString, strReplace)
		End With
	End Function
End Class
%>