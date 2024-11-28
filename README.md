# Windows Wallpaper Rotator with a Hook

A small windows wallpaper rotating app with the option to execute another app on every wallpaper change

The purpose of this program is to change windows wallpaper on user-defined intervals, with, optional, exe hook to execute another program on each wallpaper change.
This is useful if you have wallpaper overlays that need to be redrawn after a wallpaper change.

## How to use
Download the ZIP file from the Releases section.
Unzip the archive wherever you like.
The program is portable.

Edit the INI file supplied.


>[Wallpaper Directory]

You can choose the wallpaper directory, either a relative or an absolute path.

>[Hook Executable]

You can choose the hook exe path, either a relative or an absolute one.
To disable hook functionality just leave the INI key Hook Path empty.
Optionally, you can define hook arguments, and run mode (one of Hide, Min or Max).

>[Timer]

Set the timer for how often do you want the wallpaper change to occur. You can type, for example, 20s (20 seconds), or 5m (5 minutes), or 2h (2 hours), or 3d (3 days). Maximum is 49d. 

>[Startup]

Set if you would like the wallpaper to be changed on startup or not.


Enjoy!



Developed by AtmanActive, 2024.
