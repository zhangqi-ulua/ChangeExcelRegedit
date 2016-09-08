using Microsoft.Win32;
using System;

public class Program
{
    private static string _KEY_NAME = "TypeGuessRows";

    private static string[] _PATHS_32 =
    {
        @"SOFTWARE\Microsoft\Jet\4.0\Engines\Excel",
        @"SOFTWARE\Microsoft\Jet\4.0\Engines\Lotus",
        @"SOFTWARE\Microsoft\Office\12.0\Access Connectivity Engine\Engines\Excel",
        @"SOFTWARE\Microsoft\Office\12.0\Access Connectivity Engine\Engines\Lotus",
        @"SOFTWARE\Microsoft\Office\14.0\Access Connectivity Engine\Engines\Excel",
        @"SOFTWARE\Microsoft\Office\15.0\Access Connectivity Engine\Engines\Excel",
        @"SOFTWARE\Microsoft\Office\16.0\Access Connectivity Engine\Engines\Excel",
    };

    private static string[] _PATHS_64 =
    {
        @"SOFTWARE\Wow6432Node\Microsoft\Jet\4.0\Engines\Excel",
        @"SOFTWARE\Wow6432Node\Microsoft\Jet\4.0\Engines\Lotus",
        @"SOFTWARE\Wow6432Node\Microsoft\Office\12.0\Access Connectivity Engine\Engines\Excel",
        @"SOFTWARE\Wow6432Node\Microsoft\Office\12.0\Access Connectivity Engine\Engines\Lotus",
        @"SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Access Connectivity Engine\Engines\Excel",
        @"SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Access Connectivity Engine\Engines\Excel",
        @"SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Access Connectivity Engine\Engines\Excel",
    };

    static void Main(string[] args)
    {
        Console.WriteLine("本程序用于自动修改Windows7及以上版本系统注册表指定项，解决OleDb方式读取Excel文件时超过256字符的单元格中的内容无法读取完整的问题\n\n");

        OperatingSystem osInfo = Environment.OSVersion;
        // 操作系统主版本号
        int majorVersion = osInfo.Version.Major;
        // 操作系统副版本号
        int minorVersion = osInfo.Version.Minor;
        if (majorVersion < 6 || (majorVersion == 6 && minorVersion < 1))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("本程序仅支持Windows7及以上版本系统\n");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("按任意键退出");
            Console.ReadKey();
        }

        if (Environment.Is64BitOperatingSystem == true)
        {
            foreach (string path in _PATHS_64)
                ChangeOneRegeditValue(path);
        }
        else
        {
            foreach (string path in _PATHS_32)
                ChangeOneRegeditValue(path);
        }

        Console.ForegroundColor = ConsoleColor.White;
        Console.WriteLine("\n修改完毕，按任意键退出");
        Console.ReadKey();
    }

    public static void ChangeOneRegeditValue(string path)
    {
        try
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(string.Format("修改\"{0}\\{1}\"：", path, _KEY_NAME));
            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(path, true);
            if (registryKey == null)
            {
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine("不存在，无需修改");
            }
            else
            {
                object keyInfo = registryKey.GetValue(_KEY_NAME);
                if (keyInfo == null)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine(string.Format("不存在，无需修改。警告：此路径下本应存在名为\"{0}\"的项", _KEY_NAME));
                }
                else
                {
                    string value = keyInfo.ToString();
                    if ("0".Equals(value))
                    {
                        Console.ForegroundColor = ConsoleColor.Blue;
                        Console.WriteLine("值已设为0，无需修改");
                    }
                    else
                    {
                        registryKey.SetValue(_KEY_NAME, 0);
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine(string.Format("值已由\"{0}\"修改为\"0\"", value));
                    }
                }
            }
        }
        catch (System.Security.SecurityException)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("失败，必须以管理员身份运行本程序");
        }
        catch (Exception exception)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(string.Format("失败，原因为{0}", exception));
        }
    }
}
