Set Sapi = Wscript.CreateObject("SAPI.SpVoice")
set wshshell = wscript.CreateObject("wscript.shell")

dim Input

Sapi.speak "Please type, what you want to open?"
Input=inputbox ("Please type, what you want to open.")





if Input = "youtube" OR Input = "Youtube"then
Sapi.speak "Opening youtube"
wshshell.run "www.youtube.com"

else
if Input = "instructables" OR Input = "Instructables" then
Sapi.speak "Opening instructables"
wshshell.run "www.instructables.com"

else
if Input = "google" OR Input = "Cricket" then
Sapi.speak "Opening cricket"
wshshell.run "sports.html"

else
if Input = "command prompt" OR Input = "Command prompt" then
Sapi.speak "Opening command prompt"
wshshell.run "cmd"

else
if Input = "calculator" OR Input = "Calculator" then
Sapi.speak "Opening calculator"
wshshell.run "calc"

else
if Input = "notepad" OR Input = "Notepad" then
Sapi.speak "Opening notepad"
wshshell.run "notepad"

else
if Input = "Duolingo" OR Input = "duolingo" then
Sapi.speak "Opening Duolingo"
wshshell.run "www.Duolingo.com"

else
if Input = "Amazon" OR Input = "amazon" then
Sapi.speak "Opening amazon"
wshshell.run "www.Amazon.com"


else
if Input = "WhatsApp" OR Input = "whatsapp" then
Sapi.speak "Opening whatsapp"
wshshell.run "www.WhatsApp.com"


else
if Input = "Flipkart" OR Input = "flipkart" then
Sapi.speak "Opening flipkart"
wshshell.run "www.Flipkart.com"

else
if Input = "Wikipedia" OR Input = "wikipedia" then
Sapi.speak "Opening wikipedia"
wshshell.run "www.Wikipedia.org"

else
if Input = "Google Translate" OR Input = "google translate" then
Sapi.speak "opening google translate"
wshshell.run "https://translate.google.co.in/"  


else 
if Input = "Facebook" OR Input = "facebook" then
Sapi.speak "opening facebook"
wshshell.run "https://www.facebook.com/" 

else
if Input = "What is the weather?" OR Input = "what is the weather?" then
Sapi.speak "here is the weather"
wshshell.run "https://www.msn.com/en-in/weather/today/weather-today/we-city?el=KTS%2BrvcTczwMQhCds3kZaw%3D%3D&ocid=ansmsnweather" 
else



Sapi.speak "I don't recognize your input, Please try something else"
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
end if
