using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;
using System.Runtime.InteropServices;
using System;

namespace MsaSQLEditor
{
    [ComVisible(false)]
    [RunInstaller(true)]
    public partial class InstallerFunctions : System.Configuration.Install.Installer
    {

        [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Demand)]
        public override void Install(IDictionary stateSaver)
        {
           base.Install(stateSaver);

            RegistrationServices regsrv = new RegistrationServices();
            if (!regsrv.RegisterAssembly(GetType().Assembly, AssemblyRegistrationFlags.SetCodeBase))
            {
                throw new InstallException("Failed to register for COM Interop.");
            }

        }

        [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Demand)]
        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);

            RegistrationServices regsrv = new RegistrationServices();
            if (!regsrv.UnregisterAssembly(GetType().Assembly))
            {
                throw new InstallException("Failed to unregister for COM Interop.");
            }
        }
    }
}
