dim Vms

Set Vms = CreateObject("Scripting.Dictionary")

' change the path to your vmware_vsphere_client's path

Vms.Add "0", """D:\Program Files (x86)\VMware\Infrastructure\Virtual Infrastructure Client\Launcher\VpxClient.exe"" -s 192.168.1.99 -u username -p password"
Vms.Add "1", """D:\Program Files (x86)\VMware\Infrastructure\Virtual Infrastructure Client\Launcher\VpxClient.exe"" -s 192.168.1.98 -u username -p password"

For Each Vm In Vms
    set wshshell=CreateObject("wscript.shell")
    set oExec=wshshell.Exec(Vms.Item(Vm))

Next
