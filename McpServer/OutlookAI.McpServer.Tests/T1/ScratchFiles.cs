namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Best-effort removal of a test's own scratch files and folders, written once (Q154): a file an
/// antivirus scan still holds, or a folder Windows will not let go of yet, stays behind rather than
/// failing a test whose result is already decided. Used by <see cref="AttachmentFileWriterDiskTests"/>
/// and <c>T2.LiveSaveAttachmentFolderTests</c>. A folder tree is removed by .NET's recursive delete,
/// which removes a junction or symbolic link in it as a link and never follows it.
/// </summary>
internal static class ScratchFiles
{
    /// <summary>Deletes the file <paramref name="path"/> when it is there.</summary>
    internal static void TryDeleteFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    /// <summary>Deletes the folder <paramref name="path"/> and everything in it when it is there.</summary>
    internal static void TryDeleteTree(string path)
    {
        try
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, recursive: true);
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}
