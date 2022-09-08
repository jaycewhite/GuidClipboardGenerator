using GuidClipboardGenerator.Properties;
using IWshRuntimeLibrary;
using NonInvasiveKeyboardHookLibrary;

namespace GuidClipboardGenerator
{
    public class CustomApplicationContext : ApplicationContext
    {
        private NotifyIcon notifyIcon = new();
        private KeyboardHookManager keyboardHookManager = new KeyboardHookManager();

        Guid hookGuid = Guid.Empty;
        public CustomApplicationContext()
        {
            CreateShortcut("GuidGenerator", Environment.GetFolderPath(Environment.SpecialFolder.Startup), Directory.GetCurrentDirectory() + "\\GuidClipboardGenerator.exe");
            notifyIcon.Icon = Resources.icon;
            notifyIcon.ContextMenuStrip = new()
            {
                Items = {new ToolStripMenuItem("Exit",null,exit)}
            };
            notifyIcon.Visible = true;

            var identifier = notifyIcon.GetHashCode();

            hookGuid = keyboardHookManager.RegisterHotkey(ModifierKeys.Control | ModifierKeys.Alt | ModifierKeys.Shift, 0x47, GenerateGuid);
            keyboardHookManager.Start();

        }

        [STAThread]
        void GenerateGuid()
        {
            notifyIcon.ShowBalloonTip(10000, "Guid Generated", "You can now paste", ToolTipIcon.Info);
            var thread = new Thread(() => Clipboard.SetText(Guid.NewGuid().ToString()));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

        }

        void exit(object? sendr, EventArgs e)
        {
            keyboardHookManager.Stop();
            keyboardHookManager.UnregisterHotkey(hookGuid);
            notifyIcon.Visible=false;
            Application.Exit();
        }

        public void CreateShortcut(string shortcutName, string shortcutPath, string targetFileLocation)
        {
            string shortcutLocation = System.IO.Path.Combine(shortcutPath, shortcutName + ".lnk");
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutLocation);

            shortcut.Description = "GuidGenerator";   // The description of the shortcut
            shortcut.TargetPath = targetFileLocation;                 // The path of the file that will launch when the shortcut is run
            shortcut.Save();                                    // Save the shortcut
        }
    }
}
