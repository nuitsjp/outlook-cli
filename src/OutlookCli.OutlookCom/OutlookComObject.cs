using System.Runtime.InteropServices;

namespace OutlookCli.OutlookCom;

internal static class OutlookComObject
{
    public static object CreateApplication()
    {
        var applicationType = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException("クラシック Outlook の COM 登録が見つかりません。");
        return Activator.CreateInstance(applicationType)!;
    }

    public static void Release(object? instance)
    {
        if (instance is not null)
        {
            Marshal.ReleaseComObject(instance);
        }
    }
}
