Imports System.IO
Imports System.Reflection
Imports Documents.Extensions
Imports Documents.Utilities
Imports Operations.Extensions

Public Class ExtensionManager

  Public Shared Sub RegisterAllLocalExtensions()
    Try
      ' The currently executing instance of this assembly
      Dim lstrCurrentAssemblyPath As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
      Dim lobjAssembly As Assembly
      Dim lstrLocalDlls As String() = Directory.GetFiles(lstrCurrentAssemblyPath, "*.dll")

      For Each lstrDllPath As String In lstrLocalDlls
        lobjAssembly = System.Reflection.Assembly.LoadFrom(lstrDllPath)
        RegisterAssemblyExtensions(lobjAssembly)
      Next

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub



  Public Shared Sub RegisterAssemblyExtensions(ByRef lpAssembly As Assembly)
    Try

      Dim lobjCatalog = ExtensionCatalog.Instance

      ' The complete list of current extensions
      Dim lobjAvailableExtensions As ExtensionEntries = lobjCatalog.Extensions
      ' The currently executing instance of this assembly
      'Dim lobjCurrentAssembly As Assembly = Assembly.GetExecutingAssembly
      Dim lstrCurrentAssemblyPath As String = lpAssembly.Location

      ' The complete set of currently defined subclasses of OperationExtension in the Ecmg.Cts.Projects assembly
      Dim list As Object = From lobjOperationExtension In lpAssembly.GetTypes Where
          (lobjOperationExtension.IsSubclassOf(GetType(OperationExtension)) _
           AndAlso lobjOperationExtension.IsAbstract = False) Select lobjOperationExtension

      Dim lobjExtensionEntry As IExtensionInformation = Nothing
      Dim lobjExtensionInstance As IOperationInformation = Nothing

      Dim lstrExtensionName As String = String.Empty

      ' Loop through each currently defined operation extension in this assembly and 
      ' check if it is named in the extension catalog.
      For Each lobjExtension As Type In list
        lstrExtensionName = lobjExtension.Name
        'If lobjAvailableExtensions.ContainsKey(lstrExtensionName) = False Then
        lobjExtensionEntry = lobjAvailableExtensions.Item(lstrExtensionName.Replace("Operation", String.Empty))

        If lobjExtensionEntry Is Nothing Then
          ' The operation extension is not listed, we will add it.
          lobjExtensionInstance = lpAssembly.CreateInstance(lobjExtension.FullName)
          lobjCatalog.Add(lobjExtensionInstance.Name,
                                        lobjExtensionInstance.DisplayName,
                                        lobjExtensionInstance.Description,
                                        lobjExtensionInstance.CompanyName,
                                        lobjExtensionInstance.ProductName,
                                        lpAssembly.Location)
        Else
          If String.Compare(lobjExtensionEntry.Path, lstrCurrentAssemblyPath, True) <> 0 Then
            ' The entry in the catalog is referencing a different path, we want to be able to load the extension from the current assembly.
            lobjCatalog.Remove(lobjExtensionEntry.Name)
            lobjExtensionInstance = lpAssembly.CreateInstance(lobjExtension.FullName)
            lobjCatalog.Add(lobjExtensionInstance.Name,
                                          lobjExtensionInstance.DisplayName,
                                          lobjExtensionInstance.Description,
                                          lobjExtensionInstance.CompanyName,
                                          lobjExtensionInstance.ProductName,
                                          lpAssembly.Location)
          End If
          ' The operation extension is currently listed, we can skip it.
          Continue For
        End If
      Next

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

End Class
