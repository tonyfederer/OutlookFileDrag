using System;
using System.IO;
using log4net;

namespace OutlookFileDrag
{
    static class FileUtility
    {
        private static ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static string GetTempPath()
        {
            log.Debug("Getting temp path");
            string path = Path.Combine(Path.GetTempPath(), "OutlookFileDrag", Guid.NewGuid().ToString());
            if (!System.IO.Directory.Exists(path))
                System.IO.Directory.CreateDirectory(path);
            log.DebugFormat("Temp path: {0}", path);
            return path;
        }

        public static void CleanupTempPath(int tempFileExpiration)
        {
            log.Debug("Cleaning up temp path");
            string path = Path.Combine(Path.GetTempPath(), "OutlookFileDrag");
            log.InfoFormat("Temp path: {0}", path);
            if (!System.IO.Directory.Exists(path))
            {
                log.Info("Temp path does not exist");
                return;
            }

            var dirInfo = new DirectoryInfo(path);
            foreach(DirectoryInfo subfolder in dirInfo.GetDirectories())
            {
                //If folder was created before expiration window, delete it
                if (subfolder.CreationTime < DateTime.Now.AddMinutes(tempFileExpiration))
                    try
                    {
                        log.InfoFormat("Deleting temp folder: {0}", subfolder.FullName);
                        subfolder.Delete(true);
                    }
                    catch
                    {
                        log.WarnFormat("Could not delete temp folder: {0}", subfolder.FullName);
                    }
            }
        }

        public static string GetUniqueFilename(string filename)
        {
            string filenameNoExt;
            string ext;

            //If filename is too long, truncate filename
            if (filename.Length >= NativeMethods.MAX_PATH)
            {
                ext = Path.GetExtension(filename);
                filename = filename.Substring(0, NativeMethods.MAX_PATH - ext.Length - 1) + ext;
            }
            
            //If file does not exist, use original filename
            if (!File.Exists(filename))
                return filename;

            //Try appending number to filename until unique filename is found
            filenameNoExt = Path.Combine(Path.GetDirectoryName(filename), Path.GetFileNameWithoutExtension(filename));
            ext = Path.GetExtension(filename);

            for (int index = 1; index < 1024; index++)
            {
                string newFilename = string.Format("{0} ({1}){2}", filenameNoExt, index, ext);

                //If new filename is too long, truncate new filename
                if (newFilename.Length > NativeMethods.MAX_PATH)
                {
                    newFilename = string.Format("{0} ({1}){2}", filenameNoExt.Substring(0, NativeMethods.MAX_PATH - ext.Length - 8), index, ext);
                }

                if (!File.Exists(newFilename))
                    return newFilename;
            }

            //If no unique filename could be found, throw exception
            throw new Exception(string.Format("Could not generate unique filename for file {0}", filename));
        }


    }
}
