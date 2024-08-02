'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.ComponentModel
Imports System.Convert
Imports System.Diagnostics
Imports System.Environment
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.Marshal

'This module contains this program's core procedures.
Module GetMemoryModule
   'The Microsoft Windows API constants, functions and structures used by this program.
   Private Const ERROR_SUCCESS As Integer = 0%
   Private Const MEM_COMMIT As Integer = &H1000%
   Private Const MEM_PRIVATE As Integer = &H20000%
   Private Const PAGE_GUARD As Integer = &H100%
   Private Const PROCESS_QUERY_INFORMATION As Integer = &H400%
   Private Const PROCESS_VM_READ As Integer = &H10%

   <StructLayout(LayoutKind.Sequential)>
   Public Structure MEMORY_BASIC_INFORMATION
      Public BaseAddress As IntPtr
      Public AllocationBase As UIntPtr
      Public AllocationProtect As UInteger
      Public RegionSize As IntPtr
      Public State As UInteger
      Public Protect As UInteger
      Public Type As UInteger
   End Structure

   <StructLayout(LayoutKind.Sequential)>
   Public Structure SYSTEM_INFO
      Public dwOemId As UInteger
      Public dwPageSize As UInteger
      Public lpMinimumApplicationAddress As IntPtr
      Public lpMaximumApplicationAddress As IntPtr
      Public dwActiveProcessorMask As UIntPtr
      Public dwNumberOfProcessors As UInteger
      Public dwProcessorType As UInteger
      Public dwAllocationGranularity As UInteger
      Public dwProcessorLevel As UShort
      Public dwProcessorRevision As UShort
   End Structure

   <DllImport("Kernel32.dll", SetLastError:=True)> Private Function CloseHandle(ByVal hObject As IntPtr) As Integer
   End Function
   <DllImport("Kernel32.dll", SetLastError:=True)> Private Function OpenProcess(ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer
   End Function
   <DllImport("Kernel32.dll", SetLastError:=True)> Private Function ReadProcessMemory(ByVal hProcess As IntPtr, ByVal lpBaseAddress As IntPtr, ByVal lpBuffer As IntPtr, ByVal nSize As IntPtr, ByRef lpNumberOfBytesRead As UInteger) As Integer
   End Function
   <DllImport("Kernel32.dll", SetLastError:=True)> Private Function VirtualQueryEx(ByVal hProcess As IntPtr, ByVal lpAddress As IntPtr, ByVal lpBuffer As IntPtr, ByVal dwLength As UInteger) As Integer
   End Function
   <DllImport("Kernel32.dll", SetLastError:=True)> Private Sub GetSystemInfo(ByVal lpSystemInfo As IntPtr)
   End Sub

   'This procedure checks whether an error has occurred during the most recent Windows API call.
   Private Function CheckForError(Optional ReturnValue As Object = Nothing) As Object
      Try
         Dim ErrorCode As Integer = GetLastWin32Error()

         If Not ErrorCode = ERROR_SUCCESS Then
            Console.WriteLine($"API Error: {ErrorCode}")
            Console.WriteLine(New Win32Exception(ErrorCode).Message)
            Console.WriteLine($"Return value: {ReturnValue}")
            Console.WriteLine()
         End If
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return ReturnValue
   End Function

   'This procedure displays any errors that occur.
   Private Sub DisplayError(ExceptionO As Exception)
      Try
         With Console.Error
            .WriteLine($"Error: {ExceptionO.Message}")
            .WriteLine("Press Enter to continue...")
            .WriteLine()
         End With

         Console.ReadLine()
      Catch
         [Exit](0)
      End Try
   End Sub

   'This procedure writes the memory contents of the selected process to a file.
   Private Sub GetMemory(ProcessId As Integer, MemoryFile As String)
      Try
         Dim Buffer() As Byte = {}
         Dim BufferH As New IntPtr
         Dim BytesRead As New UInteger
         Dim MemoryBasicInformation As New MEMORY_BASIC_INFORMATION
         Dim MemoryBasicInformationH As New IntPtr
         Dim Offset As New IntPtr
         Dim ProcessH As New IntPtr
         Dim ReturnValue As New Integer
         Dim SystemInformation As New SYSTEM_INFO
         Dim SystemInformationH As New IntPtr

         If Not ProcessId = Nothing Then
            Process.EnterDebugMode()

            SystemInformationH = AllocHGlobal(SizeOf(SystemInformation))
            GetSystemInfo(SystemInformationH)
            SystemInformation = DirectCast(PtrToStructure(SystemInformationH, SystemInformation.GetType), SYSTEM_INFO)
            FreeHGlobal(SystemInformationH)

            ProcessH = New IntPtr(CInt(CheckForError(OpenProcess(PROCESS_VM_READ Or PROCESS_QUERY_INFORMATION, CInt(False), ProcessId))))
            If Not ProcessH = Nothing Then
               Using FileO As New FileStream(MemoryFile, FileMode.CreateNew)
                  Offset = SystemInformation.lpMinimumApplicationAddress
                  Do While Offset.ToInt64() <= SystemInformation.lpMaximumApplicationAddress.ToInt64()
                     MemoryBasicInformationH = AllocHGlobal(SizeOf(MemoryBasicInformation))
                     ReturnValue = CInt(CheckForError(VirtualQueryEx(ProcessH, Offset, MemoryBasicInformationH, CUInt(SizeOf(MemoryBasicInformation)))))
                     MemoryBasicInformation = DirectCast(PtrToStructure(MemoryBasicInformationH, MemoryBasicInformation.GetType), MEMORY_BASIC_INFORMATION)
                     FreeHGlobal(MemoryBasicInformationH)
                     If ReturnValue = 0 Then Exit Do

                     If Not (MemoryBasicInformation.Protect And PAGE_GUARD) = PAGE_GUARD Then
                        If MemoryBasicInformation.Type = MEM_PRIVATE Then
                           If MemoryBasicInformation.State = MEM_COMMIT Then
                              BufferH = AllocHGlobal(MemoryBasicInformation.RegionSize)
                              ReturnValue = CInt(CheckForError(ReadProcessMemory(ProcessH, Offset, BufferH, MemoryBasicInformation.RegionSize, BytesRead)))
                              ReDim Buffer(0 To CInt(BytesRead))
                              Copy(BufferH, Buffer, Buffer.GetLowerBound(0), Buffer.Length)
                              FreeHGlobal(BufferH)
                              If Not ReturnValue = 0 Then FileO.Write(Buffer, Buffer.GetLowerBound(0), Buffer.Length)
                           End If
                        End If
                     End If

                     Offset = New IntPtr(MemoryBasicInformation.BaseAddress.ToInt64 + MemoryBasicInformation.RegionSize.ToInt64)
                  Loop
               End Using
               CheckForError(CloseHandle(ProcessH))
            End If

            Process.LeaveDebugMode()
         End If
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try
   End Sub

   'This procedure is executed when this program is started.
   Public Sub Main()
      Try
         Dim MemoryFile As String = Nothing
         Dim ProcessId As New Integer
         Dim SelectedProcess As String = Nothing

         With My.Application.Info
            My.Computer.FileSystem.CurrentDirectory = .DirectoryPath
         End With

         Console.WriteLine($"{ProgramInformation()}{NewLine}")
         Console.Write("Path or process id (prefixed with ""*""): ")
         SelectedProcess = Console.ReadLine()

         If Not SelectedProcess = Nothing Then
            Console.Write("Write memory to the following file: ")
            MemoryFile = Console.ReadLine()

            If Not MemoryFile = Nothing Then

               If SelectedProcess.StartsWith("*") Then
                  ProcessId = ToInt32(SelectedProcess.Substring(1))
               Else
                  With New Process
                     .StartInfo.FileName = SelectedProcess
                     .Start()
                     ProcessId = .Id
                  End With
               End If

               If Not ProcessId = Nothing Then GetMemory(ProcessId, MemoryFile)
            End If
         End If
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try
   End Sub

   'This procedure returns information about this program.
   Private Function ProgramInformation() As String
      Try
#If PLATFORM = "x86" Then
         Dim BitModus As String = "x86"
#ElseIf PLATFORM = "x64" Then
         Dim BitModus As String = "x64"
#End If
         Dim Information As String = Nothing

         With My.Application.Info
            Information = $"{ .Title} ({BitModus}) v{ .Version} - by: { .CompanyName}, ***{ .Copyright}***"
         End With

         Return Information
      Catch ExceptionO As Exception
         DisplayError(ExceptionO)
      End Try

      Return Nothing
   End Function
End Module
