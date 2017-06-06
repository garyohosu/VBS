'CSV形式の１行から指定した列を取り出す(列番号は0から)
function csvRead(str,n)
	dim rdline
	dim ret
	rdline = split(str,",")
	ret = rdline(n)
end function
