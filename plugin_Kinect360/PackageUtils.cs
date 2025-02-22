using System;
using System.Diagnostics;
using System.IO;
using Windows.ApplicationModel;
using Windows.Storage;
using Amethyst.Plugins.Contract;
using System.Reflection;
using System.Threading.Tasks;

namespace plugin_Kinect360;

public static class PathsHandler
{
    public static async Task Setup()
    {
        if (IsAmethystPackaged) return;

        var root = await StorageFolder.GetFolderFromPathAsync(
            Path.Join(ProgramLocation.DirectoryName!));

        TemporaryFolderUnpackaged = await (await root
                .CreateFolderAsync("AppData", CreationCollisionOption.OpenIfExists))
            .CreateFolderAsync("TempState", CreationCollisionOption.OpenIfExists);
    }

    public static bool IsAmethystPackaged
    {
        get
        {
            try
            {
                return Package.Current is not null;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }

    public static FileInfo ProgramLocation => new(Assembly.GetExecutingAssembly().Location);

    public static StorageFolder TemporaryFolder => IsAmethystPackaged ? ApplicationData.Current.TemporaryFolder : TemporaryFolderUnpackaged;

    public static StorageFolder TemporaryFolderUnpackaged { get; set; } // Assigned on Setup()
}

public static class StorageExtensions
{
    public static void CopyToFolder(this DirectoryInfo source, string destination, bool log = false)
    {
        // Now Create all of the directories
        foreach (var dirPath in source.GetDirectories("*", SearchOption.AllDirectories))
            Directory.CreateDirectory(dirPath.FullName.Replace(source.FullName, destination));

        // Copy all the files & Replaces any files with the same name
        foreach (var newPath in source.GetFiles("*.*", SearchOption.AllDirectories))
            newPath.CopyTo(newPath.FullName.Replace(source.FullName, destination), true);
    }
}

public class GestureDetector
{
    private bool Value { get; set; }
    private bool ValueBlock { get; set; }
    private Stopwatch Timer { get; set; } = new();

    public bool Update(bool value)
    {
        // ReSharper disable once ConvertIfStatementToSwitchStatement
        if (!Value && value)
        {
            //Console.WriteLine("Restarting gesture timer...");
            ValueBlock = false;
            Timer.Restart();
            Value = true;
            return false;
        }

        if (!Value && !value)
        {
            //Console.WriteLine("Resetting gesture timer...");
            ValueBlock = false;
            Timer.Reset();
            return false;
        }

        Value = value;

        switch (Timer.ElapsedMilliseconds)
        {
            case >= 1000 when !ValueBlock:
                //Console.Write("Gesture detected! ");
                Kinect360.HostStatic?.PlayAppSound(SoundType.Focus);
                ValueBlock = true;
                return true;
            case >= 3000 when ValueBlock:
                //Console.Write("Restarting timer...");
                Kinect360.HostStatic?.PlayAppSound(SoundType.Focus);
                ValueBlock = false;
                Timer.Restart();
                return true;
            default:
                //Console.WriteLine("Gesture detected! Waiting for the timer...");
                return false;
        }
    }
}