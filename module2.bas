option explisit

public function istextobject(byval target_object as zcadentity) as boolean
  
  istextObject = false
  target_object.highlight true
  
  if not typeof target_object is zcadtext _
  and not typeof target_object is zcadmtext then
    istextobject = true
    target_object.highlight false
    thisdrawing.utility.prompt "matigai" & vbcrlf
  end if
  
end function

public sub resetcondition(byval target_object as zcadentity)
  
  if not target_object is nothing then target_object.highlight false
  
end sub
