option explicit

implements cb
private chain as cb

public macro_ as variant

public property set cb_chain(cb_ as cb) ' {
    set chain = cb_
end property ' }

public property get cb_chain() as cb ' {
    set cb_chain = chain
end property ' }

public function cb_go(arg as variant) as variant ' {
    cb_go = application.run(macro_, arg)
end function ' }
