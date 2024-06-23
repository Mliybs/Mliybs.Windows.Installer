using System;
using System.Collections.Generic;
using System.IO;
using System.Security;
using System.Security.Principal;
using IWshRuntimeLibrary;
using Microsoft.Win32;

namespace Mliybs.Windows.Installer
{
    public static class Installer
    {
        public const string DisplayName = nameof(DisplayName);

        public const string DisplayIcon = nameof(DisplayIcon);

        public const string DisplayVersion = nameof(DisplayVersion);

        public const string VersionMajor = nameof(VersionMajor);

        public const string VersionMinor = nameof(VersionMinor);

        public const string InstallLocation = nameof(InstallLocation);

        public const string InstallSource = nameof(InstallSource);

        public const string UninstallString = nameof(UninstallString);

        public const string Publisher = nameof(Publisher);

        public const string ProductID = nameof(ProductID);

        public const string Version = nameof(Version);

        public const string HelpLink = nameof(HelpLink);

        public const string HelpTelephone = nameof(HelpTelephone);

        public const string InstallDate = nameof(InstallDate);

        public const string URLInfoAbout = nameof(URLInfoAbout);

        public const string URLUpdateInfo = nameof(URLUpdateInfo);

        public const string AuthorizedCDFPrefix = nameof(AuthorizedCDFPrefix);

        public const string Comments = nameof(Comments);

        public const string Contact = nameof(Contact);

        public const string EstimatedSize = nameof(EstimatedSize);

        public const string Language = nameof(Language);

        public const string ModifyPath = nameof(ModifyPath);

        public const string Readme = nameof(Readme);

        public const string SettingsIdentifier = nameof(SettingsIdentifier);

        public const string RegOwner = nameof(RegOwner);

        public const string RegCompany = nameof(RegCompany);

        public const string NoModify = nameof(NoModify);

        public const string NoRepair = nameof(NoRepair);

        public static void Install(string id, IDictionary<string, object> pairs, bool permissionCheck = true, bool installToCurrentUser = false)
        {
            if (permissionCheck)
            {
                var identity = WindowsIdentity.GetCurrent();
                var principle = new WindowsPrincipal(identity);
                if (!principle.IsInRole(WindowsBuiltInRole.Administrator)) throw new SecurityException("请以管理员权限运行！");
            }

            var uninstall = (installToCurrentUser ? Registry.CurrentUser : Registry.LocalMachine).OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall", true);

            var app = uninstall.CreateSubKey(id);

            foreach (var (key, value) in pairs) app.SetValue(key, value);

            app.Close();
        }

        public static string Install(string id, string displayName, string displayIcon, string displayVersion, string uninstallPath, string installLocation = null, bool installToCurrentUser = false)
        {
            var path = installLocation ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), displayName);

            Install(id, new Dictionary<string, object>()
            {
                { DisplayName, displayName },
                { DisplayIcon, displayIcon },
                { DisplayVersion, displayVersion },
                { InstallLocation, path },
                { UninstallString, uninstallPath}
            }, installToCurrentUser: installToCurrentUser);

            return path;
        }

        public static void Uninstall(string id, bool permissionCheck = true, bool installToCurrentUser = false)
        {
            if (permissionCheck)
            {
                var identity = WindowsIdentity.GetCurrent();
                var principle = new WindowsPrincipal(identity);
                if (!principle.IsInRole(WindowsBuiltInRole.Administrator)) throw new SecurityException("请以管理员权限运行！");
            }

            var uninstall = (installToCurrentUser ? Registry.CurrentUser : Registry.LocalMachine).OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall", true);

            uninstall.DeleteSubKey(id);
        }

        public static void CreateShortcut(string name, string originPath, string iconLocation = null)
        {
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), name + ".lnk");
            var shell = new WshShell();
            var shortcut = (IWshShortcut)shell.CreateShortcut(path);
            shortcut.TargetPath = originPath;
            if (iconLocation is not null) shortcut.IconLocation = iconLocation;
            shortcut.Save();
        }
    }
}
