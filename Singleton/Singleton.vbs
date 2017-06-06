private m_MySingleton

class MySingleton
  public function getInstance()
    if not isObject(m_MySingleton) then
      set m_MySingleton = new MySingleton
      msgbox("new")
    end if
    set getInstance = m_MySingleton
  end function
  public sub hello
    msgbox("hello")
  end sub
end class

'usage
set instance = (new MySingleton).getInstance()
instance.hello
set instance1 = (new MySingleton).getInstance()
instance1.hello
