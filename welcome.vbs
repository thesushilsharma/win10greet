 Dim speaks,sushil,greet
 Set sushil = Wscript.CreateObject("sapi.spvoice")
 if hour(time) < 5 then
 greet = "What happen " 
 ElseIf hour(time) >= 5 AND hour(time) < 12 then
 greet = "Good Morning "
 ElseIf hour(time) >= 12 AND hour(time) < 17 then
 greet = "Good afternoon Sushil "
 ElseIf hour(time) >= 17 AND hour(time) < 20 then
 greet = "Good evening Sushil "
 Else 
 greet = "Hacking is on "
 End if
 speaks = greet + "Sushil"
 sushil.Speak speaks

 sushil.speak "The current time is"
 if hour(time) > 12 then
 sushil.Speak hour(time)-12
 else
 if hour(time) = 0 then
 sushil.Speak "12"
 else
 sushil.Speak hour(time)
 end if
 end if
 if minute(time) < 10 then
 sushil.Speak "o"
 if minute(time) < 1 then
 sushil.Speak "clock"
 else
 sushil.Speak minute(time)
 end if
 else
 sushil.Speak minute(time)
 end if
 
 if hour(time) > 12 then
 sushil.Speak "P.M."
 else
 if hour(time) = 0 then
 if minute(time) = 0 then
 sushil.Speak "Midnight"
 else
 sushil.Speak "A.M."
 end if
 else
 if hour(time) = 12 then
 if minute(time) = 0 then
 sushil.Speak "Noon"
 else
 sushil.Speak "P.M."
 end if
 else
 sushil.Speak "A.M."
 end if
 end if
 end if
