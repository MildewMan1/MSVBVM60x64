# MSVBVM60x64
There is a 32-bit Windows dll named "MSVBVM60" that contains many functions for interacting with VBA6, but there doesn't seem to be a 64-bit version of this dll that I could find. 
One of the functions within that dll allows you to Set an Object in VBA using the address (i.e. ObjPtr(obj) ) of the Object.
This library is not meant to be a replacement for that dll, but instead to implement a function that gives the same capability as the function mentioned above.

Why would this function be useful? 
Because there are times when you may want to store a reference to a parent Object within a child Object.
If the parent also keeps a reference to the child, you now have a circular reference, and if the in-class references aren't cleared before the
last out-of-class references are set to nothing, then you have created a memory leak.
With this dll, you can store an address (LongPtr) to the parent within the child class and then use "vbaObjSetByAddress" to create a temporary reference to the parent as needed.

Compile the code as a 64-bit dll (or 32-bit dll if you still want to use it with 32-bit Office).
Once compiled, declare the function in VBA like so:

    #If VBA7 Then
        Private Declare PtrSafe Function vbaObjSetByAddress Lib "((Path to the dll here))\MSVBVM60x64.dll" (ByVal objectAddress As LongPtr) As Object
    #Else
        Private Declare Function vbaObjSetByAddress Lib "((Path to the dll here))\MSVBVM60x64.dll" (ByVal objectAddress As LongPtr) As Object
    #End If

Note that it is declared Private. I would recommend that you create a Public function within this module to call it, so you can
wrap the calls in error handling code.
Example:

    Public Function ObjPtrToObject(ByVal address As LongPtr) As Object
          If address <> 0 Then                                  'Check to make sure address <> 0 (nullptr)
              'Ignore any thrown errors.
              On Error Resume Next                              'Or you can implement an error handling section in this function if desired.
              Set ObjPtrToObject = vbaObjSetByAddress(address)  'Call external function.
              On Error GoTo 0                                   'Turn error throwing back on.
          End If
    End Function

Here is an example of how to use it.

    Private Sub Testing()
          Dim oOrg As MyObject
          Dim oRetval As MyObject
          Dim address As LongPtr
          
          Set oOrg = New MyObject  'Create a new Object

          'this variable can be omitted if desired. It's just here to make it easier to read.
          address = ObjPtr(oOrg)  

          'Try to get another reference to the new object.
          Set oRetval = ObjPtrToObject(address)       'Set oRetval = ObjPtrToObject(ObjPtr(oOrg)) -- if not using the address variable

          Debug.Assert (oRetval Is oOrg)    'validate that they reference the same MyObject.

          Set oOrg = Nothing
          Set oRetval = Nothing
    End Sub

