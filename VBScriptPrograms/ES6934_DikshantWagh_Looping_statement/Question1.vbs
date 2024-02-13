'                             Online VB Compiler.
'                 Code, Compile, Run and Debug VB program online.
' Write your code in this editor and press "Run" button to execute it.


Module VBModule
    Sub Main()
        'Console.WriteLine("Hello World")
        Dim a:a=5
        Dim i,j
        for i=0 to a step 1
           for j=1 to a step 1
               if j<=i then 
               Console.Write(i)
               end if
            Next   
           Console.WriteLine()   
              
        Next
         for i=4 to 0 step -1
           for j=0 to i step 1
               if i>j then 
               Console.Write(i)
               end if
            Next   
           Console.WriteLine()   
              
        Next
    End Sub
End Module
