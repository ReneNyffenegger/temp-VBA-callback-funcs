option explicit

implements cb

private chain as cb

' public property let cb_c (c_ as cb)
'   set c = c_
' end property
' 
' public property get cb_c as cb
'   set cb_c = c
' end property


public property set cb_chain(cb_ as cb) ' {
    set chain = cb_
end property ' }

public property get cb_chain() as cb ' {
    set cb_chain = chain
end property ' }

public function cb_go(arg as variant) as variant

    if not chain is nothing then
       arg = chain.go(arg)
    end if

'   set cc = me.go(v)

    debug.print arg

    cb_go = arg

end function
