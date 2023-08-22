using Sharprompt;
using System.IO.Ports;
using System.Management;


var devices = new List<SerialDevice>();

// ベンダーIDが 2E8A(Raspberry Pi) のデバイスを列挙
foreach (ManagementObject drive in new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_SerialPort").Get())
    if (drive.GetPropertyValue("Caption") is string displayName && drive.GetPropertyValue("DeviceID") is string devId && ((drive.GetPropertyValue("PNPDeviceID") as string)?.Contains("VID_2E8A") ?? false))
        devices.Add(new SerialDevice(displayName, devId));

if (devices.Count == 0)
{
    Console.WriteLine("デバイスが見つかりませんでした。");
    Console.ReadKey();
    return;
}

// 引数で指定されたデバイスがある場合はそれを選択
var argSelectedDevice = devices.FirstOrDefault(d => args.FirstOrDefault() == d.ComName);
if (args.Length > 0 && argSelectedDevice == null)
{
    Console.WriteLine($"引数で指定されたデバイス ({args[0]}) が見つかりませんでした。");
    Console.ReadKey();
    return;
}

// 複数ある場合は選択させる
var device = argSelectedDevice ?? devices[0];
if (argSelectedDevice == null && devices.Count > 1)
    device = Prompt.Select("デバイスが複数あります。選択してください", devices, defaultValue: devices[0], textSelector: d => d.DisplayName);

// ボーレート 1200 で開いて閉じるとUSBストレージ書き込みモードで再起動する
using var port = new SerialPort(device.ComName, 1200);
try
{
    port.Open();
    port.Close();
}
catch (UnauthorizedAccessException)// ほかのアプリで使用中の場合
{
    Console.WriteLine(device.ComName + " にアクセスできません。ほかのアプリで使用していないか確認してください。");
    Console.ReadKey();
    return;
}
catch (IOException) // 本来であれば Open した時点で切断されるので無視する
{
}
Console.WriteLine(device.ComName + " を再起動させました");

record SerialDevice(string DisplayName, string ComName);
