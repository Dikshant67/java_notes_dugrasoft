'                             Online VB Compiler.
'                 Code, Compile, Run and Debug VB program online.
' Write your code in this editor and press "Run" button to execute it.


Module VBModule
    Sub Main()
   
        Dim i,j,k,m
        k=0
        Dim l:l=1
        for i=4 to 0 step -1
             l=l+1
            for k=0 to i step 1
                if k<i then 
                    Console.Write(" ")
                end if    
            Next   
            
            for j=1 to l step 1
               if j<l then 
                  Console.Write(j)
               end if
            Next
            
            for m=j-3 to 1 step -1
                if m<j then
                Console.Write(m)
                End if
            Next    
        
           Console.WriteLine()   
              
        Next
      
    End Sub
End Module
