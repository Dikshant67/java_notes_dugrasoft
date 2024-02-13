Set objXMLDoc = CreateObject("Microsoft.XMLDOM") 
objXMLDoc.async = False 
objXMLDoc.load("C:\Users\dwagh\Downloads\VBScriptPrograms\XMLDemo\demoXML.xml")
msgbox "Hi"
Set Root = objXMLDoc.documentElement 
Set NodeList = Root.getElementsByTagName("interface") 
port = 0
For Each Elem In NodeList 
MsgBox "Port " & port & " has IP address of " & Elem.text
port = port + 1
Next