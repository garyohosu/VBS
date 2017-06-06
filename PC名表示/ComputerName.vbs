         Set WshNetwork = WScript.CreateObject("WScript.Network")
         WScript.Echo "ドメイン = " & WshNetwork.UserDomain
         WScript.Echo "コンピュータ名 = " & WshNetwork.ComputerName
         WScript.Echo "ユーザー名 = " & WshNetwork.UserName
