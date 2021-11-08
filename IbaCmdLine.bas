#Compile SLL "..\bin\baCmdLine.sll"
#Include Once "Win32Api.inc"
'------------------------------------------------------------------------------
'*** Constants ***
'------------------------------------------------------------------------------
$DELIM_PARAM = "/"
$DELIM_VALUE = "="

Class cBACmdLine Common


   Instance msParamDelimiter As String ' Parameter delimiter, typically a "/"
   Instance msValueDelimiter As String ' Key/value delimiter, typically a "="
   Instance mdwParamCount As Dword     ' Number of parameters
   Instance masParams() As String      ' All parameter keys

   Instance mdwParamID As Dword        ' Collection key of parameter
   Instance mdwValueID As Dword        ' Collection key of value

   ' Read only
   Instance mwsClassName As WString

   Instance mcolValues As IPowerCollection   ' All parameter values. Data type Variant, because
                                             ' a value can be anything
'------------------------------------------------------------------------------

   Class Method Create()
   ' Do initialization

      ' Set defaults
      msParamDelimiter = $DELIM_PARAM
      msValueDelimiter = $DELIM_VALUE
      mwsClassName = "cBACmdLine"

      ' Initialize collections
      Let mcolValues = Class "PowerCollection"

   End Method
'------------------------------------------------------------------------------

   Class Method Destroy()
   ' Do cleanup

      ' Release resources
      Let mcolValues = Nothing

   End Method
'------------------------------------------------------------------------------

   Class Method FullMethodName(ByVal wsMethodName As WString) As WString
      Trace Off
      Method = "! " & mwsClassName & ":" & wsMethodName
   End Method
'------------------------------------------------------------------------------

   Interface IBACmdLine

      Inherit Dual

      ' ParamDelimiter
      Property Get ParamDelimiter() As String
         Property = msParamDelimiter
      End Property
      Property Set ParamDelimiter(ByVal sValue As String)
         msParamDelimiter = sValue
      End Property
'------------------------------------------------------------------------------
      ' ValueDelimiter
      Property Get ValueDelimiter() As String
         Property = msValueDelimiter
      End Property
      Property Set ValueDelimiter(ByVal sValue As String)
         msValueDelimiter = sValue
      End Property
'------------------------------------------------------------------------------

      Method ErrString(ByVal lErr As Long, Optional ByVal vntPrefix As Variant) As String
      '------------------------------------------------------------------------------
      'Purpose  : Returns an formatted error string from an (PB) error number
      '
      'Prereq.  : -
      'Parameter: lErr        - (PB) runtime error number
      '           vntPrefix   - Optional prefix for the returned string
      'Returns  : -
      'Note     : -
      '
      '   Author: Knuth Konrad 12.02.2016
      '   Source: -
      '  Changed: -
      '------------------------------------------------------------------------------
         Local sPrefix As String

         Trace Off

         If Not IsMissing(vntPrefix) Then
            sPrefix = Variant$(vntPrefix)
         End If

         Method = sPrefix & Format$(lErr) & " - " & Error$(lErr)

      End Method
      '------------------------------------------------------------------------------

      Method GetParamByIndex(ByVal lIndex As Long) As String
      '------------------------------------------------------------------------------
      'Purpose  : Retrieve a parameter's key by index
      '
      'Prereq.  : -
      'Parameter: lIndex   - 1 based collection index
      'Returns  : -
      'Note     : -
      '
      '   Author: Knuth Konrad 12.02.2016
      '   Source: -
      '  Changed: -
      '------------------------------------------------------------------------------
         Local vntValue As Variant

         Trace Off
         Trace Print Me.FullMethodName(FuncName$)

         ' *** Safe guard
         If lIndex < 1 Or lIndex > Me.ValuesCount Then
            Method = ""
            Exit Method
         End If

         Local wsKey As WString

         Try
            mcolValues.Entry lIndex, wsKey, vntValue

            If ObjResult = %S_False Then
               Method = ""
               Trace Print " -- GetParamByIndex: ObjResult = %S_FALSE"
            Else
               Method = wsKey
            End If
         Catch
            Trace Print " -- GetParamByIndex (Error): " & Format$(Err) & ", " & Me.ErrString(Err)
            ErrClear
         End Try

      End Method
      '------------------------------------------------------------------------------

      Method HasParam(ByVal wsKey As WString, Optional ByVal vntKeyAlias As Variant) As Long
      '------------------------------------------------------------------------------
      'Purpose  : Check wether a specific command line parameter is present
      '
      'Prereq.  : -
      'Parameter: wsKey       - parameter to search for
      '           vntKeyAlias - parameter alias
      'Returns  : -
      'Note     : -
      '
      '   Author: Knuth Konrad 12.02.2016
      '   Source: -
      '  Changed: 10.11.2016
      '           - Check for an additional (alias) parameter having the same
      '           meaning at the same time.
      '           I.e. /f and /file
      '           30.06.2017
      '           - FIX: Do NOT exit, if primray key isn't passed, but
      '           alias instead
      '------------------------------------------------------------------------------
         Local vntValue As Variant

         Trace Off
         Trace Print Me.FullMethodName(FuncName$)

         ' *** Safe guard
         If (Len(wsKey) < 1) And (IsMissing(vntKeyAlias)) Then
            Method = %TRUE
            Exit Method
         End If

         Try
            vntValue = mcolValues.Item(wsKey)

            If ObjResult = %S_Ok Then
               Method = %TRUE
               Exit Method
            Else
               Trace Print " -- HasParam wsKey: ObjResult = %S_FALSE"
            End If
         Catch
            Trace Print " -- HasParam (Error): " & Format$(Err) & ", " & Me.ErrString(Err)
            ErrClear
         End Try

         ' msKey not found, try alias instead
         If Not IsMissing(vntKeyAlias) Then
            Local wsKeyAlias As WString
            wsKeyAlias = Variant$$(vntKeyAlias)

            Try
               vntValue = mcolValues.Item(wsKeyAlias)

               If ObjResult = %S_False Then
                  Method = %FALSE
                  Trace Print " -- HasParam vntKeyAlias: ObjResult = %S_FALSE"
                  Exit Method
               Else
                  Method = %TRUE
                  Exit Method
               End If
            Catch
               Trace Print " -- HasParam (Error): " & Format$(Err) & ", " & Me.ErrString(Err)
               ErrClear
            End Try

         Else
            Trace Print " -- HasParam vntKeyAlias: parameter not passed."
         End If

      End Method
      '------------------------------------------------------------------------------

      Method ValuesAdd(ByVal wsKey As WString, ByVal vntValue As Variant) As String
      '------------------------------------------------------------------------------
      'Purpose  : Add a key/value pair to the parameters collection
      '
      'Prereq.  : -
      'Parameter: wsKey    - Key to add. Make sure it's unique!
      '           vntValue - Value to add
      'Returns  : -
      'Note     : -
      '
      '   Author: Knuth Konrad 12.02.2016
      '   Source: -
      '  Changed: -
      '------------------------------------------------------------------------------
         Local hResult As Long

         Trace Off
         Trace Print Me.FullMethodName(FuncName$)

         Trace Print "  - wsKey: " & wsKey
         Trace Print "  - vntValue: " & Variant$(vntValue)

         Try
            mcolValues.Add(wsKey, vntValue)
            hResult = ObjResult
            Trace Print "  - ObjResult: " & Hex$(hResult, 8)
            Method = wsKey
         Catch
            Trace Print " -- ValuesAdd (Error): " & Format$(Err) & ", " & Me.ErrString(Err)
            ErrClear
         End Try

      End Method
      '------------------------------------------------------------------------------

      Method ValuesClear()
      '------------------------------------------------------------------------------
      'Purpose  : Clear all values
      '
      'Prereq.  : -
      'Parameter: -
      'Returns  : -
      'Note     : -
      '
      '   Author: Knuth Konrad 12.02.2016
      '   Source: -
      '  Changed: -
      '------------------------------------------------------------------------------

         Trace Off
         Trace Print Me.FullMethodName(FuncName$)

         Try
            mcolValues.Clear()
         Catch
            Trace Print " -- ValuesClear (Error): " & Format$(Err) & ", " & Me.ErrString(Err)
            ErrClear
         End Try

      End Method
      '------------------------------------------------------------------------------

      Method ValuesCount() As Dword
      '------------------------------------------------------------------------------
      'Purpose  : Return the number of values
      '
      'Prereq.  : -
      'Parameter: -
      'Returns  : -
      'Note     : -
      '
      '   Author: Knuth Konrad 12.02.2016
      '   Source: -
      '  Changed: -
      '------------------------------------------------------------------------------

         Trace Off
         Trace Print Me.FullMethodName(FuncName$)

         Try
            Method = mcolValues.Count
         Catch
            Trace Print " -- ValuesCount Error: " & Format$(Err) & ", " & Me.ErrString(Err)
            ErrClear
            Method = 0
         End Try

      End Method
      '------------------------------------------------------------------------------

      Method ValuesGet(ByVal lIndex As Long) As Variant
      '------------------------------------------------------------------------------
      'Purpose  : Retrieve a value by index
      '
      'Prereq.  : -
      'Parameter: lIndex   - Index of value collection
      'Returns  : -
      'Note     : -
      '
      '   Author: Knuth Konrad 12.02.2016
      '   Source: -
      '  Changed: -
      '------------------------------------------------------------------------------
         Local vntValue As Variant, wsKey As WString

         Trace Off
         Trace Print Me.FullMethodName(FuncName$)

         Try
            If lIndex >= 1 And lIndex <= mcolValues.Count Then
               mcolValues.Entry lIndex, wsKey, vntValue
               If ObjResult = %S_False Then
                  Trace Print " -- ValuesGet: ObjResult = %S_FALSE"
               Else
                  Method = vntValue
               End If
            End If
         Catch
            Trace Print " -- ValuesGet(lIndex) Error: (" & Format$(lIndex) & ") " & _
               Format$(Err) & ", " & Me.ErrString(Err)
            ErrClear
         End Try

      End Method
      '------------------------------------------------------------------------------

      Method GetValueByIndex(ByVal lIndex As Long) As Variant
      '------------------------------------------------------------------------------
      'Purpose  : Retrieve a value by index (wrapper/alias for ValuesGet)
      '
      'Prereq.  : -
      'Parameter: lIndex -  Index of value collection (1 to number of parameters(=ParamCount))
      'Returns  : Value of this parameter
      'Note     : -
      '
      '   Author: Knuth Konrad 27.09.2000
      '   Source: -
      '  Changed: -
      '------------------------------------------------------------------------------

         Trace Off
         Trace Print Me.FullMethodName(FuncName$)

         ' Index within range?
         If lIndex >= 1 And lIndex <= mcolValues.Count Then
            Method = me.ValuesGet(lIndex)
         End If

      End Method
      '------------------------------------------------------------------------------

      Method GetValueByName(ByVal sParam As String, Optional ByVal vntParamAlias As Variant,  _
         Optional vntCaseSensitive As Variant) As Variant
      '------------------------------------------------------------------------------
      'Purpose  : Retrieve a parameter's value by name (key)
      '
      'Prereq.  : -
      'Parameter: sParam            -  Name (key) to retrieve value for
      '           bolCaseSensitive  -  sParam is case sensitive
      'Returns  : -
      'Note     : -
      '
      '   Author: Knuth Konrad 27.09.2000
      '   Source: -
      '  Changed: -
      '------------------------------------------------------------------------------
         Local bolCaseSensitive As Long
         Local i As Dword
         Local vntValue As Variant
         Local wsParam, wsParamAlias, wsKey As WString

         Trace Off
         Trace Print Me.FullMethodName(FuncName$)

         wsParam = sParam

         If IsMissing(vntCaseSensitive) Then
            bolCaseSensitive = %FALSE
         Else
            If IsTrue(Variant#(vntCaseSensitive)) Then
               bolCaseSensitive = %TRUE
            Else
               bolCaseSensitive = %FALSE
            End If
         End If

         If IsTrue(bolCaseSensitive) Then

            For i = 1 To mcolValues.Count
               mcolValues.Entry i, wsKey, vntValue

               Trace Print " -- GetValueByName(), wsParam, wsKey: " & wsParam & ", " & wsKey

               If LCase$(wsParam) = LCase$(wsKey) Then
                  Method = vntValue
                  Exit Method
               End If
            Next i

            If Not IsMissing(vntParamAlias) Then
               wsParamAlias = Variant$$(vntParamAlias)

               Trace Print " -- GetValueByName(), wsParamAlias, wsKey: " & wsParamAlias & ", " & wsKey

               For i = 1 To mcolValues.Count
                  mcolValues.Entry i, wsKey, vntValue
                  If LCase$(wsParamAlias) = LCase$(wsKey) Then
                     Method = vntValue
                     Exit Method
                  End If
               Next i
            End If

         Else     '// If IsTrue(bolCaseSensitive)


            For i = 1 To mcolValues.Count
               mcolValues.Entry i, wsKey, vntValue

               Trace Print " -- GetValueByName(), wsParam, wsKey: " & wsParam & ", " & wsKey

               If wsParam = wsKey Then
                  Method = vntValue
                  Exit Method
               End If
            Next i

            If Not IsMissing(vntParamAlias) Then
               wsParamAlias = Variant$$(vntParamAlias)

               Trace Print " -- GetValueByName(), wsParamAlias, wsKey: " & wsParamAlias & ", " & wsKey

               For i = 1 To mcolValues.Count
                  mcolValues.Entry i, wsKey, vntValue
                  If wsParamAlias = wsKey Then
                     Method = vntValue
                     Exit Method
                  End If
               Next i
            End If


         End If   '// If IsTrue(bolCaseSensitive)

      End Method
      '------------------------------------------------------------------------------

      Method Init(ByVal sCmd As String) As Long
      '------------------------------------------------------------------------------
      'Purpose  : Initializes this object. Must be called *prior* to using any other
      '           methods.
      '
      'Prereq.  : -
      'Parameter: sCmd        -  (PB's) COMMAND$
      'Returns  : True/False  - Init succeeded?
      'Note     : -
      '
      '   Author: Knuth Konrad 27.09.2000
      '   Source: -
      '  Changed: -
      '------------------------------------------------------------------------------
         Local i, lParamCount As Long
         Local vntValue As Variant
         Dim awsParams() As WString

         Trace Off
         Trace Print Me.FullMethodName(FuncName$)

         Trace Print "  - sCmd: " & sCmd

         On Error GoTo InitError

         ' ** Safe guards
         ' Any parameters at all?
         sCmd = LTrim$(sCmd)
         lParamCount = ParseCount(sCmd, msParamDelimiter)
         If lParamCount < 2 Then
            Method = %TRUE
            Exit Method
         End If

         Trace Print "  - ParseCount(sCmd): " & Format$(lParamCount)

         ' ParseCount returns 3 for strings like "/123 /abc" where delimiter is '/', as there's an
         ' 'empty' entry in front of the first '/'
         Dim awsParams(lParamCount - 1) As WString
         Parse sCmd, awsParams(), msParamDelimiter
         Me.ValuesClear

         Trace Print "  - LBound(awsParams): " & Format$(LBound(awsParams))
         Trace Print "  - UBound(awsParams): " & Format$(UBound(awsParams))

         For i = LBound(awsParams) To UBound(awsParams)
            Local wsKey As WString

            Trace Print "  - awsParams(i): (" & Format$(i) & "), " & awsParams(i)

            ' Only if there's at least one parameter...
            If Len(Trim$(awsParams(i))) > 0 Then

               wsKey = Trim$(Remove$(Extract$(awsParams(i), msValueDelimiter), msParamDelimiter))

               Trace Print "  - wsKey: " & wsKey

               ' Parameter is of type /User=Knuth.
               ' "User" is the key, "Knuth" the value
               If InStr(awsParams(i), msValueDelimiter) > 0 Then
                  vntValue = Trim$(Remain$(awsParams(i), msValueDelimiter))

                  Trace Print "  - vntValue: " & Variant$(vntValue)

                  Me.ValuesAdd(wsKey, vntValue)

               Else
               'Parameter is of type /Quit.
               '"Quit" is the key, if present, the value is %True
                  Me.ValuesAdd(wsKey, %TRUE)

                  Trace Print "  - vntValue: %TRUE"

               End If

            End If

         Next i

         Method = %TRUE

      InitExit:
         On Error GoTo 0
         Exit Method

      InitError:
         Method = %FALSE
         Trace Print " -- Init(sCmd) Error: " & sCmd & Format$(Err) & ", " & Me.ErrString(Err)
         ' Cleanup collection
         Me.ValuesClear
         ErrClear
         Resume InitExit

      End Method
      '==============================================================================

   End Interface
'------------------------------------------------------------------------------

End Class
