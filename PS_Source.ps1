$scriptblock = {
    $id = get-random;
    $code =@"
public class program$id {
        static string resultstr = string.Empty; public static ROG71BoxControlCOMLib.ROG71BoxControlCOM ROG71BoxCtrl = null;
        public static void Main() {
            if (ROG71BoxCtrl == null) { try {
                    ROG71BoxCtrl = new ROG71BoxControlCOMClass();
                    int num = 0; ROG71BoxCtrl.Initialize(out num);
                    if (num == 1) {
                        ROG71BoxControlCOMLib.EDeviceType DeviceType = ROG71BoxControlCOMLib.EDeviceType.EType_Unknown;
                        ROG71BoxCtrl.GetDeviceType(out DeviceType);
                        if (DeviceType == ROG71BoxControlCOMLib.EDeviceType.EType_Headphone) { ROG71BoxCtrl.SetDeviceType(ROG71BoxControlCOMLib.EDeviceType.EType_Speaker); }
                        if (DeviceType == ROG71BoxControlCOMLib.EDeviceType.EType_Speaker) { ROG71BoxCtrl.SetDeviceType(ROG71BoxControlCOMLib.EDeviceType.EType_Headphone); }
                        ROG71BoxCtrl.GetDeviceType(out DeviceType);
                        resultstr = "New Device:\t" + DeviceType.ToString().Replace("EType_", "");
                    } else { ROG71BoxCtrl = null; resultstr = "Initialization failed"; }
                } catch (System.Runtime.InteropServices.COMException Ex) {resultstr = "Error Message:\r\n"+Ex.Message; ROG71BoxCtrl = null; }}}
        public static string result() { return resultstr; }}

[System.Runtime.InteropServices.ComImport, System.Runtime.InteropServices.Guid("2EECE3E8-84C6-4AE1-893C-851398DBA4F4")]
public class ROG71BoxControlCOMClass : ROG71BoxControlCOMLib.IROG71BoxControl, ROG71BoxControlCOMLib.ROG71BoxControlCOM {
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.InternalCall)]
    public virtual extern void Initialize(out int pbStatus);
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.InternalCall)]
    public virtual extern void GetVersionNumber(out uint pMajor, out uint pMinor);
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.InternalCall)]
    public virtual extern void GetDeviceType(out ROG71BoxControlCOMLib.EDeviceType pDeviceType);
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.InternalCall)]
    public virtual extern void SetDeviceType([System.Runtime.InteropServices.In] ROG71BoxControlCOMLib.EDeviceType DeviceType);
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.InternalCall)]
    public virtual extern void GetMute([System.Runtime.InteropServices.In] ROG71BoxControlCOMLib.EDeviceType DeviceType, [System.Runtime.InteropServices.In] ROG71BoxControlCOMLib.EChannelIndex ChannelIndex, out int pMuteState);
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.InternalCall)]
    public virtual extern void SetMute([System.Runtime.InteropServices.In] ROG71BoxControlCOMLib.EDeviceType DeviceType, [System.Runtime.InteropServices.In] ROG71BoxControlCOMLib.EChannelIndex ChannelIndex, [System.Runtime.InteropServices.In] int MuteState);
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.InternalCall)]
    public virtual extern void GetVolume([System.Runtime.InteropServices.In] ROG71BoxControlCOMLib.EDeviceType DeviceType, [System.Runtime.InteropServices.In] ROG71BoxControlCOMLib.EChannelIndex ChannelIndex, out uint pVolume);
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.InternalCall)]
    public virtual extern void SetVolume([System.Runtime.InteropServices.In] ROG71BoxControlCOMLib.EDeviceType DeviceType, [System.Runtime.InteropServices.In] ROG71BoxControlCOMLib.EChannelIndex ChannelIndex, [System.Runtime.InteropServices.In] uint Volume);
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.InternalCall)]
    public virtual extern void Get71ModeHP(out int pMode);
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.InternalCall)]
    public virtual extern void RegisterNotificationClient([System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Interface)][System.Runtime.InteropServices.In] ROG71BoxControlCOMLib.IROG71BoxControlNotificationClient pNotificationClient);
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.InternalCall)]
    public virtual extern void UnregisterNotificationClient([System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Interface)][System.Runtime.InteropServices.In] ROG71BoxControlCOMLib.IROG71BoxControlNotificationClient pNotificationClient);
}
"@;

    $dll = [System.Reflection.Assembly]::LoadFile("C:\Program Files\ASUSTeKcomputer.Inc\nhAsusROG71\UserInterface\Interop.ROG71BoxControlCOMLib.dll");
    Add-Type -Language CSharp -TypeDefinition $code -ReferencedAssemblies $dll;
    $classname = "program$id"
    $($classname -as [Type])::Main();
    $data = $($classname -as [Type])::result();
    return $data;
};

New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT -ErrorAction SilentlyContinue |Out-Null
$REGPATH = "HKCR:\WOW6432Node\CLSID\{2EECE3E8-84C6-4AE1-893C-851398DBA4F4}";
if (!(Test-Path -Path $REGPATH)){
     Start-Process powershell -Verb runas -ArgumentList "-command & { New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT; New-Item -Path 'HKCR:\WOW6432Node\CLSID\{2EECE3E8-84C6-4AE1-893C-851398DBA4F4}';New-Item -Path 'HKCR:\WOW6432Node\CLSID\{2EECE3E8-84C6-4AE1-893C-851398DBA4F4}\InprocServer32';Set-ItemProperty -Path 'HKCR:\WOW6432Node\CLSID\{2EECE3E8-84C6-4AE1-893C-851398DBA4F4}\InprocServer32' -Value 'C:\Progra~1\ASUSTeKcomputer.Inc\nhAsusROG71\UserInterface\ROG71BoxControlCOM.dll' -Name '(Default)';New-ItemProperty -Path 'HKCR:\WOW6432Node\CLSID\{2EECE3E8-84C6-4AE1-893C-851398DBA4F4}\InprocServer32' -Name 'ThreadingModel' -PropertyType String -Value 'Both';}" -ErrorAction SilentlyContinue;
}
$jobiD = Start-Job -ScriptBlock $scriptblock -RunAs32; 
while ($jobiD.State -ne "Completed"){};
Receive-Job -Job $jobiD;