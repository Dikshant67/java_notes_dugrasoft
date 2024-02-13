'4. Get the error description & error number in below code ( use Err object)
    ' Dim arrCar(2)
    ' arrCar(0) = "Maruti"
    ' arrCar(1) = "Tata"
    ' arrCar(5) = "Mahindra"
	
option Explicit
on error resume Next	
 Dim arrCar(2)
  arrCar(0) = "Maruti"
  arrCar(1) = "Tata"
  arrCar(5) = "Mahindra"
	
	
msgbox "Error Number = "&Err.Number
msgbox "Error Description = "&Err.Description	
Err.clear
on error goto 0
	