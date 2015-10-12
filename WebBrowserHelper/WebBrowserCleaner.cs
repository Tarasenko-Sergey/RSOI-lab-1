using System;
using System.Runtime.InteropServices;
using FileTime = System.Runtime.InteropServices.ComTypes.FILETIME;

namespace WebBrowserHelper
{
    /// <summary>
    ///     http://stackoverflow.com/a/24401521
    /// </summary>
    public static class WebBrowserCleaner
    {
        /// <summary>
        ///     Clears the cache of the web browser
        /// </summary>
        public static void ClearCache()
        {
            const int errorInsufficientBuffer = 0x7A;

            DeleteGroups();

            // Start to delete URLs that do not belong to any group.
            var cacheEntryInfoBufferSizeInitial = 0;
            ClearCacheNativeMethods.FindFirstUrlCacheEntry(null, IntPtr.Zero, ref cacheEntryInfoBufferSizeInitial);
            // should always fail because buffer is too small
            if (Marshal.GetLastWin32Error() != errorInsufficientBuffer)
                return;

            var cacheEntryInfoBufferSize = cacheEntryInfoBufferSizeInitial;
            var cacheEntryInfoBuffer = Marshal.AllocHGlobal(cacheEntryInfoBufferSize);
            try
            {
                var enumHandle = ClearCacheNativeMethods.FindFirstUrlCacheEntry(null, cacheEntryInfoBuffer,
                    ref cacheEntryInfoBufferSizeInitial);
                if (enumHandle != IntPtr.Zero)
                {
                    bool more;
                    do
                    {
                        var internetCacheEntry =
                            Marshal.PtrToStructure<ClearCacheNativeMethods.InternetCacheEntryInfoA>(cacheEntryInfoBuffer);
                        cacheEntryInfoBufferSizeInitial = cacheEntryInfoBufferSize;
                        ClearCacheNativeMethods.DeleteUrlCacheEntry(internetCacheEntry.SourceUrlName);
                        more = ClearCacheNativeMethods.FindNextUrlCacheEntry(enumHandle, cacheEntryInfoBuffer,
                            ref cacheEntryInfoBufferSizeInitial);
                        if (!more && Marshal.GetLastWin32Error() == errorInsufficientBuffer)
                        {
                            cacheEntryInfoBufferSize = cacheEntryInfoBufferSizeInitial;
                            cacheEntryInfoBuffer = Marshal.ReAllocHGlobal(cacheEntryInfoBuffer,
                                (IntPtr)cacheEntryInfoBufferSize);
                            more = ClearCacheNativeMethods.FindNextUrlCacheEntry(enumHandle, cacheEntryInfoBuffer,
                                ref cacheEntryInfoBufferSizeInitial);
                        }
                    } while (more);
                }
            }
            finally
            {
                Marshal.FreeHGlobal(cacheEntryInfoBuffer);
            }
        }

        /// <summary>
        ///     Delete the groups first.
        ///     Groups may not always exist on the system.
        ///     For more information, visit the following Microsoft Web site:
        ///     http://msdn.microsoft.com/library/?url=/workshop/networking/wininet/overview/cache.asp
        ///     By default, a URL does not belong to any group. Therefore, that cache may become
        ///     empty even when the CacheGroup APIs are not used because the existing URL does not belong to any group.
        /// </summary>
        private static void DeleteGroups()
        {
            // Indicates that all of the cache groups in the user's system should be enumerated
            const int cacheGroupSearchAll = 0x0;
            // Indicates that all the cache entries that are associated with the cache group
            // should be deleted, unless the entry belongs to another cache group.
            const int cacheGroupFlagFlushUrlOnDelete = 0x2;

            long groupId = 0;
            var enumHandle = ClearCacheNativeMethods.FindFirstUrlCacheGroup(0, cacheGroupSearchAll, IntPtr.Zero, 0,
                ref groupId,
                IntPtr.Zero);
            if (enumHandle == IntPtr.Zero) return;

            do
            {
                ClearCacheNativeMethods.DeleteUrlCacheGroup(groupId, cacheGroupFlagFlushUrlOnDelete, IntPtr.Zero);
            } while (ClearCacheNativeMethods.FindNextUrlCacheGroup(enumHandle, ref groupId, IntPtr.Zero));
        }

        private static class ClearCacheNativeMethods
        {
            /// <summary>
            ///     Initiates the enumeration of the cache groups in the Internet cache
            /// </summary>
            [DllImport(@"wininet",
                SetLastError = true,
                CharSet = CharSet.Auto,
                EntryPoint = "FindFirstUrlCacheGroup",
                CallingConvention = CallingConvention.StdCall)]
            public static extern IntPtr FindFirstUrlCacheGroup(
                int flags,
                int filter,
                IntPtr searchConditionPtr,
                int searchCondition,
                ref long groupId,
                IntPtr reserved);

            /// <summary>
            ///     Retrieves the next cache group in a cache group enumeration
            /// </summary>
            [DllImport(@"wininet",
                SetLastError = true,
                CharSet = CharSet.Auto,
                EntryPoint = "FindNextUrlCacheGroup",
                CallingConvention = CallingConvention.StdCall)]
            public static extern bool FindNextUrlCacheGroup(
                IntPtr find,
                ref long groupId,
                IntPtr reserved);

            /// <summary>
            ///     Releases the specified GROUPID and any associated state in the cache index file
            /// </summary>
            [DllImport(@"wininet",
                SetLastError = true,
                CharSet = CharSet.Auto,
                EntryPoint = "DeleteUrlCacheGroup",
                CallingConvention = CallingConvention.StdCall)]
            public static extern bool DeleteUrlCacheGroup(
                long groupId,
                int flags,
                IntPtr reserved);

            /// <summary>
            ///     Begins the enumeration of the Internet cache
            /// </summary>
            [DllImport(@"wininet",
                SetLastError = true,
                CharSet = CharSet.Auto,
                EntryPoint = "FindFirstUrlCacheEntryA",
                CallingConvention = CallingConvention.StdCall)]
            public static extern IntPtr FindFirstUrlCacheEntry(
                [MarshalAs(UnmanagedType.LPTStr)] string urlSearchPattern,
                IntPtr firstCacheEntryInfo,
                ref int firstCacheEntryInfoBufferSize);


            /// <summary>
            ///     Retrieves the next entry in the Internet cache
            /// </summary>
            [DllImport(@"wininet",
                SetLastError = true,
                CharSet = CharSet.Auto,
                EntryPoint = "FindNextUrlCacheEntryA",
                CallingConvention = CallingConvention.StdCall)]
            public static extern bool FindNextUrlCacheEntry(
                IntPtr find,
                IntPtr nextCacheEntryInfo,
                ref int nextCacheEntryInfoBufferSize);

            /// <summary>
            ///     Removes the file that is associated with the source name from the cache, if the file exists
            /// </summary>
            [DllImport(@"wininet",
                SetLastError = true,
                CharSet = CharSet.Auto,
                EntryPoint = "DeleteUrlCacheEntryA",
                CallingConvention = CallingConvention.StdCall)]
            public static extern bool DeleteUrlCacheEntry(IntPtr urlName);

            /// <summary>
            ///     Contains information about an entry in the Internet cache
            /// </summary>
            [StructLayout(LayoutKind.Explicit)]
            public struct ExemptDeltaOrReserverD
            {
                [FieldOffset(0)]
                public readonly uint Reserved;
                [FieldOffset(0)]
                public readonly uint ExemptDelta;
            }

            [StructLayout(LayoutKind.Sequential)]
            public struct InternetCacheEntryInfoA
            {
                public readonly uint StructSize;
                public readonly IntPtr SourceUrlName;
                public readonly IntPtr LocalFileName;
                public readonly uint CacheEntryType;
                public readonly uint UseCount;
                public readonly uint HitRate;
                public readonly uint SizeLow;
                public readonly uint SizeHigh;
                public readonly FileTime LastModifiedTime;
                public readonly FileTime ExpireTime;
                public readonly FileTime LastAccessTime;
                public readonly FileTime LastSyncTime;
                public readonly IntPtr HeaderInfo;
                public readonly uint HeaderInfoSize;
                public readonly IntPtr FileExtension;
                public readonly ExemptDeltaOrReserverD ExemptDeltaOrReserved;
            }
        }
    }
}