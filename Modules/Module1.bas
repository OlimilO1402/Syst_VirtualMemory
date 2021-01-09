Attribute VB_Name = "Module1"
'Option Explicit
'VirtualAlloc
'Reserves or commits a region of pages in the virtual address space of the calling process. Memory allocated by this function is automatically
'initialized to zero, unless MEM_RESET is specified.
'To allocate memory in the address space of another process, use the VirtualAllocEx function.
'
'Syntax
'
'LPVOID WINAPI VirtualAlloc(
'  __in_opt  LPVOID lpAddress,
'  __in      SIZE_T dwSize,
'  __in      DWORD flAllocationType,
'  __in      DWORD flProtect
');
'
'Parameters
' lpAddress [in, optional]
'   The starting address of the region to allocate. If the memory is being reserved, the specified address is rounded down to the nearest
'   multiple of the allocation granularity. If the memory is already reserved and is being committed, the address is rounded down to the
'   next page boundary. To determine the size of a page and the allocation granularity on the host computer, use the GetSystemInfo function.
'   If this parameter is NULL, the system determines where to allocate the region.
' dwSize [In]
'   The size of the region, in bytes. If the lpAddress parameter is NULL, this value is rounded up to the next page boundary.
'   Otherwise, the allocated pages include all pages containing one or more bytes in the range from lpAddress to lpAddress+dwSize.
'   This means that a 2-byte range straddling a page boundary causes both pages to be included in the allocated region.
'
' flAllocationType [In]
'   The type of memory allocation. This parameter must contain one of the following values.
'
'
'
'Value                'Meaning
'
'MEM_COMMIT 0x1000
'Allocates physical storage in memory or in the paging file on disk for the specified reserved memory pages. The function initializes the
'memory to zero.
'To reserve and commit pages in one step, call VirtualAlloc with MEM_COMMIT | MEM_RESERVE.
'The function fails if you attempt to commit a page that has not been reserved. The resulting error code is ERROR_INVALID_ADDRESS.
'An attempt to commit a page that is already committed does not cause the function to fail. This means that you can commit pages without
'first determining the current commitment state of each page.
'
'MEM_RESERVE 0x2000
'Reserves a range of the process's virtual address space without allocating any actual physical storage in memory or in the paging file on disk.
'You can commit reserved pages in subsequent calls to the VirtualAlloc function. To reserve and commit pages in one step, call VirtualAlloc with
'MEM_COMMIT | MEM_RESERVE.
'Other memory allocation functions, such as malloc and LocalAlloc, cannot use a reserved range of memory until it is released.
'
'MEM_RESET 0x80000
'Indicates that data in the memory range specified by lpAddress and dwSize is no longer of interest. The pages should not be read from or
'written to the paging file. However, the memory block will be used again later, so it should not be decommitted. This value cannot be used
'with any other value.
'Using this value does not guarantee that the range operated on with MEM_RESET will contain zeros. If you want the range to contain zeros,
'decommit the memory and then recommit it.
'When you specify MEM_RESET, the VirtualAlloc function ignores the value of flProtect. However, you must still set flProtect to a valid
'protection value, such as PAGE_NOACCESS.
'VirtualAlloc returns an error if you use MEM_RESET and the range of memory is mapped to a file. A shared view is only acceptable if it is
'mapped to a paging file.
'
'
'
'
'This parameter can also specify the following values as indicated.
'
'Value Meaning

'MEM_LARGE_PAGES 0x20000000
'Allocates memory using large page support.
'The size and alignment must be a multiple of the large-page minimum. To obtain this value, use the GetLargePageMinimum function.
'
'MEM_PHYSICAL 0x400000
'Reserves an address range that can be used to map Address Windowing Extensions (AWE) pages.
'This value must be used with MEM_RESERVE and no other values.
'
'MEM_TOP_DOWN 0x100000
'Allocates memory at the highest possible address. This can be slower than regular allocations, especially when there are many allocations.
'
'MEM_WRITE_WATCH 0x200000
'Causes the system to track pages that are written to in the allocated region. If you specify this value, you must also specify MEM_RESERVE.
'
'To retrieve the addresses of the pages that have been written to since the region was allocated or the write-tracking state was reset, call
'the GetWriteWatch function. To reset the write-tracking state, call GetWriteWatch or ResetWriteWatch. The write-tracking feature remains
'enabled for the memory region until the region is freed.
'
'
'
' flProtect [In]
'The memory protection for the region of pages to be allocated. If the pages are being committed, you can specify any one of the memory
'protection constants.
'Return value
'If the function succeeds, the return value is the base address of the allocated region of pages.
'If the function fails, the return value is NULL. To get extended error information, call GetLastError.
'
'Remarks
'Each page has an associated page state. The VirtualAlloc function can perform the following operations:
' •Commit a region of reserved pages
' •Reserve a region of free pages
' •Simultaneously reserve and commit a region of free pages
'
'VirtualAlloc cannot reserve a reserved page. It can commit a page that is already committed. This means you can commit a range of pages,
'regardless of whether they have already been committed, and the function will not fail.
'You can use VirtualAlloc to reserve a block of pages and then make additional calls to VirtualAlloc to commit individual pages from the
'reserved block. This enables a process to reserve a range of its virtual address space without consuming physical storage until it is needed.
'If the lpAddress parameter is not NULL, the function uses the lpAddress and dwSize parameters to compute the region of pages to be allocated.
'The current state of the entire range of pages must be compatible with the type of allocation specified by the flAllocationType parameter.
'Otherwise, the function fails and none of the pages are allocated. This compatibility requirement does not preclude committing an already
'committed page, as mentioned previously.
'To execute dynamically generated code, use VirtualAlloc to allocate memory and the VirtualProtect function to grant PAGE_EXECUTE access.
'
'The VirtualAlloc function can be used to reserve an Address Windowing Extensions (AWE) region of memory within the virtual address space of
'a specified process. This region of memory can then be used to map physical pages into and out of virtual memory as required by the application.
'The MEM_PHYSICAL and MEM_RESERVE values must be set in the AllocationType parameter. The MEM_COMMIT value must not be set.
'The page protection must be set to PAGE_READWRITE.
'The VirtualFree function can decommit a committed page, releasing the page's storage, or it can simultaneously decommit and release a
'committed page. It can also release a reserved page, making it a free page.
'When creating a region that will be executable, the calling program bears responsibility for ensuring cache coherency via an appropriate
'call to FlushInstructionCache once the code has been set in place. Otherwise attempts to execute code out of the newly executable region
'may produce unpredictable results.
'
'Examples
'
'For an example, see Reserving and Committing Memory.
'The following example illustrates the use of the VirtualAlloc and VirtualFree functions in reserving and committing memory as needed for
'a dynamic array. First, VirtualAlloc is called to reserve a block of pages with NULL specified as the base address parameter, forcing the
'system to determine the location of the block. Later, VirtualAlloc is called whenever it is necessary to commit a page from this reserved
'region, and the base address of the next page to be committed is specified.
'The example uses structured exception-handling syntax to commit pages from the reserved region. Whenever a page fault exception occurs
'during the execution of the __try block, the filter function in the expression preceding the __except block is executed. If the filter
'function can allocate another page, execution continues in the __try block at the point where the exception occurred. Otherwise, the
'exception handler in the __except block is executed. For more information, see Structured Exception Handling.
'As an alternative to dynamic allocation, the process can simply commit the entire region instead of only reserving it. Both methods result
'in the same physical memory usage because committed pages do not consume any physical storage until they are first accessed. The advantage
'of dynamic allocation is that it minimizes the total number of committed pages on the system. For very large allocations, pre-committing
'an entire allocation can cause the system to run out of committable pages, resulting in virtual memory allocation failures.
'The ExitProcess function in the __except block automatically releases virtual memory allocations, so it is not necessary to explicitly
'free the pages when the program terminates through this execution path. The VirtualFree function frees the reserved and committed pages
'if the program is built with exception handling disabled. This function uses MEM_RELEASE to decommit and release the entire region of
'reserved and committed pages.
'The following C++ example demonstrates dynamic memory allocation using a structured exception handler.
'
'// A short program to demonstrate dynamic memory allocation
'// using a structured exception handler.
'
'#include <windows.h>
'#include <tchar.h>
'#include <stdio.h>
'#include <stdlib.h>             // For exit
'
'#define PAGELIMIT 80            // Number of pages to ask for
'
'LPTSTR lpNxtPage;               // Address of the next page to ask for
'DWORD dwPages = 0;              // Count of pages gotten so far
'DWORD dwPageSize;               // Page size on this computer
'
'INT PageFaultExceptionFilter(DWORD dwCode)
'{
'    LPVOID lpvResult;
'
'    // If the exception is not a page fault, exit.
'
'    if (dwCode != EXCEPTION_ACCESS_VIOLATION)
'    {
'        _tprintf(TEXT("Exception code = %d.\n"), dwCode);
'        return EXCEPTION_EXECUTE_HANDLER;
'    }
'
'    _tprintf(TEXT("Exception is a page fault.\n"));
'
'    // If the reserved pages are used up, exit.
'
'    if (dwPages >= PAGELIMIT)
'    {
'        _tprintf(TEXT("Exception: out of pages.\n"));
'        return EXCEPTION_EXECUTE_HANDLER;
'    }
'
'    // Otherwise, commit another page.
'
'    lpvResult = VirtualAlloc(
'                     (LPVOID) lpNxtPage, // Next page to commit
'                     dwPageSize,         // Page size, in bytes
'                     MEM_COMMIT,         // Allocate a committed page
'                     PAGE_READWRITE);    // Read/write access
'    if (lpvResult == NULL )
'    {
'        _tprintf(TEXT("VirtualAlloc failed.\n"));
'        return EXCEPTION_EXECUTE_HANDLER;
'    }
'    Else
'    {
'        _tprintf(TEXT("Allocating another page.\n"));
'    }
'
'    // Increment the page count, and advance lpNxtPage to the next page.
'
'    dwPages++;
'    lpNxtPage = (LPTSTR) ((PCHAR) lpNxtPage + dwPageSize);
'
'    // Continue execution where the page fault occurred.
'
'    return EXCEPTION_CONTINUE_EXECUTION;
'}
'
'VOID ErrorExit(LPTSTR lpMsg)
'{
'    _tprintf(TEXT("Error! %s with error code of %ld.\n"),
'             lpMsg, GetLastError ());
'    exit (0);
'}
'
'VOID _tmain(VOID)
'{
'    LPVOID lpvBase;               // Base address of the test memory
'    LPTSTR lpPtr;                 // Generic character pointer
'    BOOL bSuccess;                // Flag
'    DWORD i;                      // Generic counter
'    SYSTEM_INFO sSysInfo;         // Useful information about the system
'
'    GetSystemInfo(&sSysInfo);     // Initialize the structure.
'
'    _tprintf (TEXT("This computer has page size %d.\n"), sSysInfo.dwPageSize);
'
'    dwPageSize = sSysInfo.dwPageSize;
'
'    // Reserve pages in the virtual address space of the process.
'
'    lpvBase = VirtualAlloc(
'                     NULL,                 // System selects address
'                     PAGELIMIT*dwPageSize, // Size of allocation
'                     MEM_RESERVE,          // Allocate reserved pages
'                     PAGE_NOACCESS);       // Protection = no access
'    if (lpvBase == NULL )
'        ErrorExit(TEXT("VirtualAlloc reserve failed."));
'
'    lpPtr = lpNxtPage = (LPTSTR) lpvBase;
'
'    // Use structured exception handling when accessing the pages.
'    // If a page fault occurs, the exception filter is executed to
'    // commit another page from the reserved block of pages.
'
'    for (i=0; i < PAGELIMIT*dwPageSize; i++)
'    {
'        __try
'        {
'            // Write to memory.
'
'            lpPtr[i] = 'a';
'        }
'
'        // If there's a page fault, commit another page and try again.
'
'        __except ( PageFaultExceptionFilter( GetExceptionCode() ) )
'        {
'
'            // This code is executed only if the filter function
'            // is unsuccessful in committing the next page.
'
'            _tprintf (TEXT("Exiting process.\n"));
'
'            ExitProcess( GetLastError() );
'
'        }
'
'    }
'
'    // Release the block of pages when you are finished using them.
'
'    bSuccess = VirtualFree(
'                       lpvBase,       // Base address of block
'                       0,             // Bytes of committed pages
'                       MEM_RELEASE);  // Decommit the pages
'
'    _tprintf (TEXT("Release %s.\n"), bSuccess ? TEXT("succeeded") : TEXT("failed") );
'
'}
'##############################################################################################################
'VirtualProtect function
'Changes the protection on a region of committed pages in the virtual address space of the calling process.
'To change the access protection of any process, use the VirtualProtectEx function.
'
'Syntax
'
'BOOL WINAPI VirtualProtect(
'  __in   LPVOID lpAddress,
'  __in   SIZE_T dwSize,
'  __in   DWORD flNewProtect,
'  __out  PDWORD lpflOldProtect
');
'
'Parameters
' lpAddress [In]
'   A pointer an address that describes the starting page of the region of pages whose access protection attributes are to be changed.
'   All pages in the specified region must be within the same reserved region allocated when calling the VirtualAlloc or VirtualAllocEx
'   function using MEM_RESERVE. The pages cannot span adjacent reserved regions that were allocated by separate calls to VirtualAlloc
'   or VirtualAllocEx using MEM_RESERVE.
' dwSize [In]
'   The size of the region whose access protection attributes are to be changed, in bytes. The region of affected pages includes all
'   pages containing one or more bytes in the range from the lpAddress parameter to (lpAddress+dwSize). This means that a 2-byte range
'   straddling a page boundary causes the protection attributes of both pages to be changed.
' flNewProtect [In]
'   The memory protection option. This parameter can be one of the memory protection constants.
'   This value must be compatible with the access protection specified for the pages using VirtualAlloc or VirtualAllocEx.
' lpflOldProtect [out]
'   A pointer to a variable that receives the previous access protection value of the first page in the specified region of pages.
'   If this parameter is NULL or does not point to a valid variable, the function fails.
'
'Return value
'If the function succeeds, the return value is nonzero.
'If the function fails, the return value is zero. To get extended error information, call GetLastError.
'
'Remarks
'You can set the access protection value on committed pages only. If the state of any page in the specified region is not committed, the
'function fails and returns without modifying the access protection of any pages in the specified region.
'The PAGE_GUARD protection modifier establishes guard pages. Guard pages act as one-shot access alarms. For more information, see Creating
'Guard Pages.
'It is best to avoid using VirtualProtect to change page protections on memory blocks allocated by GlobalAlloc, HeapAlloc, or LocalAlloc,
'because multiple memory blocks can exist on a single page. The heap manager assumes that all pages in the heap grant at least read and write access.
'When protecting a region that will be executable, the calling program bears responsibility for ensuring cache coherency via an appropriate
'call to FlushInstructionCache once the code has been set in place. Otherwise attempts to execute code out of the newly executable region may
'produce unpredictable results.
'
'Requirements
'   Minimum supported client
'       Windows XP
'   Minimum supported server
'       Windows Server 2003
'
'Header
'Winbase.h (include Windows.h)
'
'Library
'Kernel32.lib
'
'dll
'Kernel32.dll
'
'###########################################################################################################
'
'VirtualLock function
'Locks the specified region of the process's virtual address space into physical memory, ensuring that subsequent access to the region
'will not incur a page fault.
'
'Syntax
'
'BOOL WINAPI VirtualLock(
'  __in  LPVOID lpAddress,
'  __in  SIZE_T dwSize
');
'
'Parameters
' lpAddress [In]
'   A pointer to the base address of the region of pages to be locked.
' dwSize [In]
'   The size of the region to be locked, in bytes. The region of affected pages includes all pages that contain one or more bytes in the range
'   from the lpAddress parameter to (lpAddress+dwSize). This means that a 2-byte range straddling a page boundary causes both pages to be locked.
'
'Return value
'   If the function succeeds, the return value is nonzero.
'   If the function fails, the return value is zero. To get extended error information, call GetLastError.
'
'Remarks
'   All pages in the specified region must be committed. Memory protected with PAGE_NOACCESS cannot be locked.
'   Locking pages into memory may degrade the performance of the system by reducing the available RAM and forcing the system to swap out other
'   critical pages to the paging file. Each version of Windows has a limit on the maximum number of pages a process can lock. This limit is
'   intentionally small to avoid severe performance degradation. Applications that need to lock larger numbers of pages must first call the
'   SetProcessWorkingSetSize function to increase their minimum and maximum working set sizes. The maximum number of pages that a process can
'   lock is equal to the number of pages in its minimum working set minus a small overhead.
'   Pages that a process has locked remain in physical memory until the process unlocks them or terminates. These pages are guaranteed not to
'   be written to the pagefile while they are locked.
'   To unlock a region of locked pages, use the VirtualUnlock function. Locked pages are automatically unlocked when the process terminates.
'   This function is not like the GlobalLock or LocalLock function in that it does not increment a lock count and translate a handle into a
'   pointer. There is no lock count for virtual pages, so multiple calls to the VirtualUnlock function are never required to unlock a region
'   of pages.
'
'Examples
'
'For an example, see Creating Guard Pages.
'
'Requirements
'   Minimum supported client
'       Windows XP
'   Minimum supported server
'       Windows Server 2003
'Header
'   Winbase.h (include Windows.h)
'Library
'   Kernel32.lib
'dll
'   Kernel32.dll
'
