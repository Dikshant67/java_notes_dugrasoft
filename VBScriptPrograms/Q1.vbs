'Prerquisite - "c:\MyFolder\Trainee_Names.xlsx"
'->Sheet1 - available data -> cell(2,1) = "Hrushikesh" , cell(4,1) = "Prajakta" , Cell(5,1) = "Ranajeet" , Cell(3,1)= "Siddharth"
'Write a vb script which will read above excel and create folders at "c:\MyFolder\" location as per Trainee Names
'Expected O/P : after executing .vbs 4 different folders should get created at "c:\MyFolder\" location with folder names Hrushikesh,Prajakta,Ranajeet,Siddharth

Option Explicit

Dim File_Obj

Dim xl_obj,xl_Workbook,xl_worksheet

Dim src,dest,NEW_PATH

Dim irow_count

Dim int_i:int_i=0

dest="c:\MyFolder\AniketGadage"

ForRead=1

src="c:\MyFolder\Trainee_Names.xlsx"

Set xl_Obj=CreateObject("excel.Application")
xl_Obj.visible=True
Set xl_Workbook=xl_obj.Workbooks.open(src)

Set xl_worksheet=xl_Workbook.worksheets(1)

xl_worksheet.cells(2,1).value = "Hrushikesh" 
xl_worksheet.cells(4,1).value = "Prajakta" 
xl_worksheet.cells(5,1).value  = "Ranajeet"  
xl_worksheet.cells(3,1).value = "Siddharth"

irow_count=xl_worksheet.usedrange.ROWS.Count

File_Obj.CreateFolder dest

For int_i=2 to irow_count
	NEW_PATH=dest+"\"+xl_worksheet.cells(int_i,1).value
	File_Obj.CreateFolder NEW_PATH
Next

set File_Obj=Nothing
Set Write_Obj=Nothing
sET Read_Obj=Nothing


' ========================================================================================
'2. (10 Marks)
'Prerquisite - "c:\MyFolder\MyDetails.txt" file is present
'Line 1 : Hi am Sandeep 
'Line 2 : I have total 4+ Years of experience in Manual testing
'Line 3 : I have handson experience of Retail and PLM domain
'Line 4 : I am associated with Expleo since past 2 years
'Line 5 : I am learning VB scripting
'write a vb script which will read above file and paste the content of this file in to other file but order of sentence will be reversed
'(last sentence should occure as first sentence in 2nd file,second last sentence at 2nd line... )
'Expected O/p "c:\MyFolder\CopyMyDetails.txt" with below lines
'Line 1 : I am learning VB scripting
'Line 2 : I am associated with Expleo since past 2 years
'Line 3 : I have handson experience of Retail and PLM domain
'Line 4 : I have total 4+ Years of experience in Manual testing
'Line 5 : Hi am Sandeep 

Option Explicit

Dim File_Obj,Write_Obj,Read_Obj
Dim src,dest
Dim ForRead,ForWrite
Dim Str
Dim arr(4)
Dim int_i:int_i=0

str=""

ForRead=1
ForWrite=2

src="c:\MyFolder\MyDetails.txt"
dest="c:\MyFolder\CopyMyDetails.txt"

Set File_Obj=CreateObject("scripting.FileSystemObject")


IF Not(File_Obj.FileExists(src)) Then
	File_Obj.CreateTextFile src
end IF

Set Read_Obj=File_Obj.OpenTextFile(src,ForRead)

IF Not(File_Obj.FileExists(dest)) Then
	File_Obj.CreateTextFile dest
end IF

Set Write_Obj=File_Obj.OpenTextFile(dest,ForWrite)

Do Until Read_Obj.AtEndOfStream

	arr(int_i)=Read_Obj.ReadLine()
	int_i=int_i+1
Loop

For int_i=UBound(arr) to 0 step -1
	Write_Obj.WriteLine(arr(int_i))
Next

set File_Obj=Nothing
Set Write_Obj=Nothing
sET Read_Obj=Nothing





' ==============================================================================================
'3. (10 Marks)
'arrFirstArray = Array(1,2,3,4,5)
'arrSecondArray = Array(3,4,5,6,7)
'write a vb script to get common values from both the arrays 
'Expected O/p - 3,4,5

Option Explicit

Dim First_Arr,Second_Arr,Result_Arr()
Dim int_i,int_j,int_k

int_k=0

First_Arr=Array(1,2,3,4,5)
Second_Arr=Array(3,4,5,6,7)

For int_i=0 to UBound(First_Arr)
	For int_j=0 to UBound(Second_Arr)
			If First_Arr(int_i)=Second_Arr(int_j) Then
				Redim Preserve Result_Arr(int_k)
				Result_Arr(int_k)=First_Arr(int_i)
				int_k=int_k+1
			end If
	Next
Next

MsgBox Join(Result_Arr,",")
' =========================================================================================
'4. (10 Marks)
'Write a program  to print below pattern with help of mininmum possible loops 
			
'			03
'		04	06	08
'	04	08	12	16	20



Option Explicit


Dim Row_count,col_count

Dim Int_i,Int_j,space1,ele_count
Dim Str


Str=""

Row_count=3
col_count=Row_count+2
space1=0
ele_count=1

For Int_i =1 TO Row_count
	Space1=Row_count-Int_i
	
	str=Str+Space(Space1)

	space1=Space1+1
	
	For int_j=1 to ele_count
			
		str=str+cstr(Space1*Int_i)+" "
		Space1 =Space1+1
	Next
	Str=Str&vbNewLine
	ele_count=ele_count+2	

Next

MsgBox Str
' ==========================================================================================
'5. (10 Marks)
'Write a vb script program to accept Car's brand name and car name  from user 
'based on brand and Car, Prize should get populated 
'(e.g TATA - Safari 25L,Nexon 15L,Punch 10L,Tigor 8L, Mahindra - XUV700 19 L,XUV500 15L,XUV300 12 L
' , Honda - Jazz 9L,Amaze 10L,City 12L, Skoda - Kushaq 17 L,Slavia 18L,Octavia 20 L) 


Option Explicit

Dim Car_Brand,Car_Name
Dim Concat_Name

Car_Brand=LCase(InputBox("Enter The Car Brand Name"))
Car_Name=LCase(InputBox("Enter The Car Name"))


Select Case Car_Brand

Case "tata"

	If Car_Name="safari" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 25L"
	ELSEIF Car_Name="nexon" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 15L"
	ELSEIF Car_Name="punch" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 10L"
	ELSEIF Car_Name="tigor" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 8L"
	ELSE
		MsgBox "NO SUCH BRAND AVILABLE"
	End IF

Case "mahindra"
	If Car_Name="xuv700" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 19L"
	ELSEIF Car_Name="xuv500" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 15L"
	ELSEIF Car_Name="xuv300" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 12L"
	ELSE
		MsgBox "NO SUCH BRAND AVILABLE"
	End IF

Case "honda"
	If Car_Name="jazz" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 9L"
	ELSEIF Car_Name="amaze" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 10L"
	ELSEIF Car_Name="city" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 12L"
	ELSE
		MsgBox "NO SUCH BRAND AVILABLE"
	End IF

Case "skoda"
	If Car_Name="kushaq" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 17L"
	ELSEIF Car_Name="slavia" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 18L"
	ELSEIF Car_Name="octavia" Then
		MsgBox Car_Brand&"  "&Car_Name&" Price = 20L"
	ELSE
		MsgBox "NO SUCH BRAND AVILABLE"
	End IF

End Select
' ============================================================================
'6. (10 Marks)
'Write a vb script to create a new excel and store below table.
'use minimum for loops
'1	2	3	4	5
'2	4	6	8	10
'3	6	9	12	15
'4	8	12	16	20
'5	10	15	20	25
'6	12	18	24	30
'7	14	21	28	35
'8	16	24	32	40
'9	18	27	36	45
'10	20	30	40	50




Option Explicit

Dim File_Obj

Dim xl_obj,xl_Workbook,xl_worksheet

Dim src,dest,NEW_PATH

Dim irow_count,icol_count

Dim int_i:int_i=0

dest="c:\MyFolder\AniketGadage"

ForRead=1
irow_count=10
icol_count=5

src="c:\MyFolder\Trainee_Names.xlsx"

Set xl_Obj=CreateObject("excel.Application")

xl_Obj.visible=True

Set xl_Workbook=xl_obj.Workbooks.open(src)

Set xl_worksheet=xl_Workbook.worksheets(1)

	for int_i=1 to icol_count
			for int_j=1 to irow_count
				xl_worksheet.cells(irow_count,icol_count).value = icol_count*irow_count
			Next
	Next

set File_Obj=Nothing
Set Write_Obj=Nothing
sET Read_Obj=Nothing






	
Dim
' ===============================================================================
'7. (10 Marks)
'Write a program to replace 3rd & 4th Expleo/expleo word with "Expleo India" . Make sure while showing output 
'msgbox should show entire string with replaced words, and not the half string from where you started replacing
'strOrgDetails = "Expleo office Pune. expleo head count is 2K. I work in Expleo. expleo is very nice organisation"


Option Explicit

Dim strOrgDetails,TEMP_LEFT,TEMP_RIGHT,ch
Dim Search_Count,int_i
strOrgDetails = "Expleo office Pune. expleo head count is 2K. I work in Expleo. expleo is very nice organisation"
Search_Count=0
TEMP_LEFT=""
TEMP_RIGHT=""

For int_i=1 to Len(strOrgDetails)
	ch=Mid(strOrgDetails,int_i,6)
	
	IF ch="Expleo" Or Ch="expleo" Then
		Search_Count=Search_Count+1
		IF Search_Count=3 OR Search_Count=4 Then
			TEMP_LEFT=Left(strOrgDetails,int_i-1)
			TEMP_RIGHT=Mid(strOrgDetails,int_i+6,Len(strOrgDetails)-Len(TEMP_LEFT)-6)
			strOrgDetails=TEMP_LEFT+"Expleo India"+TEMP_RIGHT
		end IF
	end if	
Next

MsgBOX strOrgDetails
' =====================================================================================
'8. (10 Marks)
'A. (5 Marks)
'Write a program to calculate total number of months difference between below dates
'dtDate1 = #07 Jan 2015#
'dtDate2 = #08 Jun 2016#
'dtDate3 = #10 Apr 2017#
'dtDate4 = #09 Sep 2019#

Option Explicit

Dim dtDate1,dtDate2,dtDate3,dtDate4

Dim Month_Sum:Month_Sum=0

dtDate1 = #07 Jan 2015#
dtDate2 = #08 Jun 2016#
dtDate3 = #10 Apr 2017#
dtDate4 = #09 Sep 2019#

Month_Sum=DateDiff("m",dtDate1,dtDate2)+DateDiff("m",dtDate2,dtDate3)+DateDiff("m",dtDate3,dtDate4)

MsgBox "Month_Sum = "&Month_Sum
' ====================================================================================
'5 Marks'
'B. Calculate length of the string without using inbuilt function



Option Explicit

Dim STR
Dim int_i,length_count
Dim ch
STR="Expleo office Pune."
length_count=0
	for int_i=1 to 100
	
		ch=Mid(STR,int_i,1)
		
		if ch="" Then
			Exit for
		end if
		length_count=length_count+1
	Next
	
	MsgBox "length Of String = "&length_count
	' ================================================================================
	'9. (10 Marks)
'Write a program to handle below erros for uninterrupted execution. get the error number and description as well for fail log

'Option Explicit
'intVal1  =  take input from user 'eg 2
'intVal2  =  take input from user 'eg 2
'intVal2 = intVal1 - intVal2 
'intVal2 = intVal1/intVal2
'msgbox intVal2



Option Explicit


Dim intVal1,intVal2

ON error Resume NEXT

	intVal1  =  Int(InputBox("ENTER THE NUMBER 1"))
	intVal2  =  Int(InputBox("ENTER THE NUMBER 2"))
	intVal2 = intVal1 - intVal2 
	intVal2 = intVal1/intVal2

	if Err.number<>0 Then
		msgbox "ERROR number = "&Err.number
		msgbox "ERROR description = "&Err.description
	Else
		msgbox "Code Run Sucessfull"
	end if

Err.clear

ON error Goto 0
' ===================================================================
'10. (10 Marks)
' a program to count the occurence of perticular sentence after 35th line in given text file 



Dim File_Obj,Read_Obj
Dim src
Dim ForRead,Search_Count,Position
Dim Str,Search_Text
Dim int_i:int_i=0

str=""

ForRead=1
Search_Count=0
Position=1
src="c:\MyFolder\sample.txt"

Set File_Obj=CreateObject("scripting.FileSystemObject")


IF Not(File_Obj.FileExists(src)) Then
	File_Obj.CreateTextFile src
end IF

Set Read_Obj=File_Obj.OpenTextFile(src,ForRead)

str=Read_Obj.Readall

Search_Text=InputBox("Enter the Search_Text")

For int_i=10 to Len(str)
	ch=Mid(str,int_i,Len(Search_Text))
	
	IF ch=Search_Text Then
		Search_Count=Search_Count+1
	end if	
Next

MsgBox Search_Count

set File_Obj=Nothing
Set Write_Obj=Nothing
sET Read_Obj=Nothing





