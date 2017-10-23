Imports System.ComponentModel
Imports System.Reflection
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.IO
Imports System.Xml.Serialization

Namespace AppSharedLayer

    Public Class MemberInit

        Public Shared Function initMembers(ByVal [object] As Object) As Object
            Dim locObj As Object = [object]
            Try
                Dim locType As Type = locObj.GetType
                Dim locMembers() As MemberInfo = locType.GetMembers()
                For Each locMember As MemberInfo In locMembers
                    If locMember.MemberType = MemberTypes.Property Then
                        Dim locPropertyInfo As PropertyInfo = CType(locMember, PropertyInfo)
                        Dim locAttributeCollection As AttributeCollection = TypeDescriptor.GetProperties(locObj)(locMember.Name).Attributes
                        Dim locDefaultAttribute As DefaultValueAttribute = CType(locAttributeCollection(GetType(DefaultValueAttribute)), DefaultValueAttribute)
                        If locDefaultAttribute IsNot Nothing And locPropertyInfo.CanWrite = True Then
                            locPropertyInfo.SetValue(locObj, locDefaultAttribute.Value, Nothing)
                        End If
                    End If
                Next
                Return locObj
            Catch ex As Exception
                GlobalErrManagement.GlobalErrHandler(ex)
                Return [object]
            End Try
        End Function

    End Class

    Public Class ADObjectCloner

        Public Shared Function DeepCopy(ByVal [Object] As Object) As Object
            Try
                Return DeserializeFromByteArray(SerializeToByteArray([Object]))
            Catch ex As Exception
                GlobalErrManagement.GlobalErrHandler(ex)
                Return Nothing
            End Try
        End Function

        Shared Function SerializeToByteArray(ByVal [object] As Object) As Byte()
            Try
                Dim retByte() As Byte
                Dim locMs As MemoryStream = New MemoryStream
                Dim locBinaryFormatter As New BinaryFormatter(Nothing, New StreamingContext(StreamingContextStates.Clone))
                locBinaryFormatter.Serialize(locMs, [object])
                locMs.Flush()
                locMs.Close()
                retByte = locMs.ToArray()
                Return retByte

            Catch ex As Exception
                GlobalErrManagement.GlobalErrHandler(ex)
                Return [object]
            End Try
        End Function

        Shared Function DeserializeFromByteArray(ByVal by As Byte()) As Object
            Try
                Dim locObject As Object
                Dim locFs As MemoryStream = New MemoryStream(by)
                Dim locBinaryFormatter As New BinaryFormatter(Nothing, New StreamingContext(StreamingContextStates.File))
                locObject = locBinaryFormatter.Deserialize(locFs)
                locFs.Close()
                Return locObject
            Catch ex As Exception
                GlobalErrManagement.GlobalErrHandler(ex)
                Return by
            End Try
        End Function

    End Class

    Public Class GlobalErrManagement

        Public Shared Function GlobalErrHandler(ByVal ex As Exception) As ErrResult

            Dim errString As String

            errString = "Ausnahme ausgelöst von: " & My.Computer.Name & vbCrLf & _
                        "Ausnahme ausgelöst um : " & DateTime.Now.ToString & vbCrLf & _
                        "Ausnahme ausgelöst in : " & ex.GetType.FullName & vbCrLf & _
                        "Ausnahmentext         : " & ex.Message & vbCrLf & _
                        "Ausnahmenquelle       : " & ex.Source & vbCrLf & _
                        "Ausnahmenmethode      : " & ex.TargetSite.ToString & vbCrLf & _
                        "Ausnahmen_StackTrace  : " & ex.StackTrace.ToString & vbCrLf

            If Debugger.IsAttached Then
                Debug.WriteLine(errString)
                Stop
            Else
                File.WriteAllText(Microsoft.VisualBasic.FileIO.FileSystem.CombinePath(Application.StartupPath, _
                                                                                      "ErrLog.txt"), errString)

                Dim myForm As New Form

                Dim myTB As New TextBox With {.Dock = DockStyle.Fill, .Multiline = True, .Text = errString}

                myForm.Controls.Add(myTB)



                myForm.Show()



                End
            End If

            Return ErrResult.unhandledException

        End Function


        Public Enum ErrResult As Int16
            unhandledException
            unhandledWarning
        End Enum

    End Class

    Public Class XMLSerialization

        Public Shared Function DeSerialization(ByRef xmlFile As String, ByVal [Object] As Object) As Object
            Try
                Using locFileStream As New FileStream(xmlFile, FileMode.Open, FileAccess.Read)
                    Dim locFormatter As New XmlSerializer([Object].GetType)
                    Return DirectCast(locFormatter.Deserialize(locFileStream), [Object])
                End Using
            Catch ex As Exception
                Return Nothing
            End Try
        End Function


        Public Shared Function Serialization(ByRef xmlFile As String, ByVal [Object] As Object) As Boolean
            Try
                Using locFileStream As New FileStream(xmlFile, FileMode.Create, FileAccess.Write)
                    Dim locFormatter As New XmlSerializer([Object].GetType)
                    locFormatter.Serialize(locFileStream, [Object])
                    locFileStream.Close()
                    Return True
                End Using
            Catch ex As Exception
                Return False
            End Try
        End Function

    End Class


End Namespace
