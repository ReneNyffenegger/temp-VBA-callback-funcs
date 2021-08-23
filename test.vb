option explicit


' global debugPrint     as new cb_debugPrint
' global applicationRun as new cb_applicationRun


sub main() ' {

     dim c as new collection
 
     c.add "foo"
     c.add "bar"
     c.add "baz"
' 
' '   dim t as debugPrint
' 
    forEach c, debugPrint
    forEach c, debugPrint(applicationRun("addX"))
    forEach c, debugPrint(eval_(" @ & "" - "" & @"))
' 
'     forEach c, debugPrint(applicationRun("addX"))
' 
end sub ' }
' 
function addX(v as variant) ' {

    addX = v & "-X"

end function ' }
 
 function applicationRun(macro as variant) as variant ' {
 
     set applicationRun = new cb_applicationRun
     applicationRun.macro_ = macro
 
 end function ' }
 
function debugPrint(optional cb_ as cb) as cb ' {

    set debugPrint = new cb_debugPrint

    if not isMissing(cb_) then
    '  set debugPrint.cb_chain = cb_
       set debugPrint.chain = cb_
    end if 

end function ' }

function eval_(arg as variant) as cb ' {

  '
  ' need a cb_eval_ (not a cb)
  ' in order to set formula.
  '
    dim ret as new cb_eval_
'   set ret = new cb_eval_

    if varType(arg) = vbVariant or varType(arg) = vbString then ' {
       ret.formula = arg
    else
       debug.print "TODO, varType " & varType(arg)
    end if ' }

    set eval_ = ret

'   if not isMissing(cb_) then
    '  set debugPrint.cb_chain = cb_
'      set eval.chain = cb_
'   end if 
end function ' }

sub forEach(domain as variant, func as cb) ' {

    dim e as variant

    for each e in domain
        func.go(e)
    next e

end sub ' }
