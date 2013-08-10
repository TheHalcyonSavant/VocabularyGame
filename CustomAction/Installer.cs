using System.Collections;
using System.ComponentModel;
using System.Security.Permissions;
using System.Security.AccessControl;
using System.Security.Principal;
using System.IO;

namespace CustomAction
{
    [RunInstaller(true)]
    public partial class Installer : System.Configuration.Install.Installer
    {
        public Installer()
        {
            InitializeComponent();
        }

        [SecurityPermission(SecurityAction.Demand)]
        public override void Install(IDictionary stateSaver)
        {
            base.Install(stateSaver);
        }

        [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Demand)]
        public override void Commit(IDictionary savedState)
        {
            base.Commit(savedState);
            string dir = Context.Parameters["TDIR"].ToString();
            DirectorySecurity sec = Directory.GetAccessControl(dir);
            sec.AddAccessRule(new FileSystemAccessRule(new SecurityIdentifier(WellKnownSidType.WorldSid, null),
                FileSystemRights.Modify | FileSystemRights.Synchronize, InheritanceFlags.ContainerInherit
                | InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow
            ));
            Directory.SetAccessControl(dir, sec);
            base.Dispose();
        }

        [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Demand)]
        public override void Rollback(IDictionary savedState)
        {
            base.Rollback(savedState);
        }

        [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Demand)]
        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);
        }
    }
}
