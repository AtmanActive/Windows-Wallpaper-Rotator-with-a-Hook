#SingleInstance Force

A_AllowMainWindow := 0

; https://www.autohotkey.com/docs/v2/

; GET APP'S RUNTIME PATH
my_exe_path := A_ScriptDir


; GET APP'S INI FILE PATH
my_ini_path := SubStr( my_exe_path "\" A_ScriptName, 1, -4 ) ".ini"
my_ico_path := SubStr( my_exe_path "\" A_ScriptName, 1, -4 ) ".ico"


; SET DEFAULT VALUES FOR PARAMETERS
default_value_wallpaper_path := "wallpapers"
default_value_hook_path := ""
default_value_hook_args := ""
default_value_hook_run := "Hide"
default_value_change_timer := "2h"
default_value_change_start := 1

; CHECK IF INI FILE EXISTS
if not FileExist( my_ini_path )
{
	; CREATE INI FILE IF IT DOESN'T EXIST
	IniWrite default_value_wallpaper_path, my_ini_path, "Wallpaper Directory", "Images"
	IniWrite default_value_hook_path, my_ini_path, "Hook Executable", "Hook Path"
	IniWrite default_value_hook_args, my_ini_path, "Hook Executable", "Hook Arguments"
	IniWrite default_value_hook_run, my_ini_path, "Hook Executable", "Hook Launch Mode"
	IniWrite default_value_change_timer, my_ini_path, "Timer", "Change"
	IniWrite default_value_change_start, my_ini_path, "Startup", "Change"
}


; READ VALUES FROM INI FILE
ini_wallpaper_path := IniRead( my_ini_path, "Wallpaper Directory", "Images", default_value_wallpaper_path )
ini_hook_path := IniRead( my_ini_path, "Hook Executable", "Hook Path", default_value_hook_path )
ini_hook_args := IniRead( my_ini_path, "Hook Executable", "Hook Arguments", default_value_hook_args )
ini_hook_run := IniRead( my_ini_path, "Hook Executable", "Hook Launch Mode", default_value_hook_run )
ini_change_timer := IniRead( my_ini_path, "Timer", "Change", default_value_change_timer )
ini_change_start := IniRead( my_ini_path, "Startup", "Change", default_value_change_start )

use_wallpaper_path := ""
use_hook_path := ""
use_timer_integer := 0
use_timer_unit := "m"
use_timer_seconds := 0

use_wallpaper_menu_name := ""
use_wallpaper_menu_prefix := "Current wallpaper: "
use_wallpaper_filename := ""
use_wallpaper_fullpath := ""

; CONVERT HOOK PATH TO ABSOLUTE, IF IT IS A RELATIVE ONE
If StrLen( ini_hook_path )
{
	If InStr( ini_hook_path, ":" )
	{
		use_hook_path := ini_hook_path
	}
	else
	{
		use_hook_path := my_exe_path "\" ini_hook_path
	}
}



; CONVERT WALLPAPER PATH TO ABSOLUTE, IF IT IS A RELATIVE ONE
If InStr( ini_wallpaper_path, ":" )
{
	use_wallpaper_path := ini_wallpaper_path
}
else
{
	use_wallpaper_path := my_exe_path "\" ini_wallpaper_path
}	



; CHECK IF WALLPAPER DIRECTORY EXISTS
if not DirExist( use_wallpaper_path )
{
	MsgBox "ERROR`n`nWallpaper directory doesn't exist: `n`n" use_wallpaper_path
  ExitApp 1
}



; CHECK IF WALLPAPER DIRECTORY CONTAINS FILES
filewCount := 0
Loop Files, use_wallpaper_path "\*.jpg"
{
	filewCount++
}
if ( filewCount = 0 )
{
	MsgBox "ERROR`n`nNo jpg files found in the specified directory: `n`n" use_wallpaper_path
	ExitApp 1
}



; CHECK HOOK EXECUTABLE
If StrLen( ini_hook_path )
{
	if not FileExist( use_hook_path )
	{
		MsgBox "ERROR`n`nHook executable doesn't exist: `n`n" use_hook_path
		ExitApp 1
	}
}



; PARSE INPUT TIMER STRING AND CONVERT TO TOTAL SECONDS
If StrLen( ini_change_timer )
{
	
	timer_integer_split := SubStr( ini_change_timer, 1, ( StrLen( ini_change_timer ) - 1 ) )
	timer_unit_split := SubStr( ini_change_timer, -1, 1 )
	
	if ( timer_unit_split ~= "\A(s|m|h|d)\z" )
	{
		use_timer_unit := timer_unit_split
	}
	else
	{
		AbortOnInvalidTimerString()
	}
	
	if ( IsInteger( timer_integer_split ) )
	{
		use_timer_integer := Number( timer_integer_split )
	}
	else
  {
    AbortOnInvalidTimerString()
  }
	
	if ( use_timer_unit == "s" )
  {
    use_timer_seconds := use_timer_integer
  }
	else if ( use_timer_unit == "m" )
  {
    use_timer_seconds := use_timer_integer * 60
  }
	else if ( use_timer_unit == "h" )
  {
    use_timer_seconds := use_timer_integer * 60 * 60
	}
  else if ( use_timer_unit == "d" )
  {
		
		if ( use_timer_integer > 49 )
		{
      MsgBox "ERROR`n`nTimer value is too large. `n`nPlease enter a value less than or equal to 49 days."
			ExitApp 1
    }
		
		use_timer_seconds := use_timer_integer * 60 * 60 * 24
  }
  
}
else
{
  AbortOnInvalidTimerString()
}





;
;
; FUNCTIONS
;
;



; FUNCTION TO STOP IF TIMER STRING CAN'T BE PARSED CORRECTLY
AbortOnInvalidTimerString()
{
  MsgBox "ERROR`n`nInvalid timer string. `n`nPlease format it as a digit followed by a letter, one of s, m, h, d`n`nWhere s is for seconds, m is for minutes, h is for hours and d is for days.`n`nFor example: 2h (two hours), or 15m (fifteen minutes), or 5d (five days)"
  ExitApp 1
}





; FUNCTION TO EXECUTE AN ACTUAL WALLPAPER CHANGE
ChangeWallpaper( imagePath ) 
{
	DllCall( "SystemParametersInfo"
		, "UInt", 0x14     ; SPI_SETDESKWALLPAPER
		, "UInt", 0
		, "Str", imagePath
		, "UInt", 3)       ; SPIF_UPDATEINIFILE | SPIF_SENDCHANGE
}





; FUNCTION TO CHOOSE A RANDOM WALLPAPER FILE AND APPLY IT
ChangeWPaperFromDir( folderPath )
{
	
	fileCount := 0
	
	; First loop to count files
	Loop Files, folderPath "\*.jpg"
	{
		fileCount++
	}
	
	if ( fileCount = 0 )
	{
		MsgBox "ERROR`n`nNo jpg files found in the specified directory: `n`n" folderPath
    ExitApp 1
	}
	
	; Generate random number between 1 and fileCount
	randNum := Random( 1, fileCount )
	
	; Second loop to get the random file
	currentIndex := 0
	Loop Files, folderPath "\*.jpg" 
	{
		currentIndex++
		if ( currentIndex = randNum )
		{
			global use_wallpaper_filename := A_LoopFileName
			global use_wallpaper_fullpath := folderPath "\" A_LoopFileName
			changeWallpaper( use_wallpaper_fullpath )
			;;;MsgBox use_wallpaper_menu_name
			A_TrayMenu.Rename( use_wallpaper_menu_name, use_wallpaper_menu_prefix use_wallpaper_filename )
			TrayMenuNameDynamicSet()
			Break
		}
	}
	
}




; FUNCTION TO EXECUTE THE HOOK NOW
ExecuteHookNow()
{
  if StrLen( use_hook_path )
	{
		if StrLen( ini_hook_args )
		{
			Run( use_hook_path " " ini_hook_args, , ini_hook_run )
		}
		else
		{
			Run( use_hook_path, , ini_hook_run )
		}
	}
}



; FUNCTION TO CHANGE WALLPAPER AND EXECUTE THE HOOK, IF ANY
ChangeWallpaperAndExecuteHook()
{
  
  ChangeWPaperFromDir( use_wallpaper_path )
	
	ExecuteHookNow()
	
}



; FUNCTION TO SET THE TRAY MENU
SetTrayMenu()
{
	
	TraySetIcon my_ico_path, 1, 1
	
	
	; Remove all standard menu items
	A_TrayMenu.Delete()
	
	
	
	A_TrayMenu.Add( A_ScriptName, MenuHandlerTitle )
	A_TrayMenu.Disable( A_ScriptName )
	
	A_TrayMenu.Add()  ; Creates a separator line.
	A_TrayMenu.Add( "Change wallpaper now", MenuHandlerChangeNow )
	if StrLen( use_hook_path )
	{
		A_TrayMenu.Add()  ; Creates a separator line.
		A_TrayMenu.Add( "Execute the hook now", MenuHandlerExecuteNow )
	}
	else
	{
		A_TrayMenu.Add()  ; Creates a separator line.
		A_TrayMenu.Add( "No hook defined to execute", MenuHandlerNoHookDefined )
	}
	A_TrayMenu.Add()  ; Creates a separator line.
	A_TrayMenu.Add( "Changing wallpapers every " use_timer_seconds " seconds", MenuHandlerTimer )  ; Creates a new menu item.
	A_TrayMenu.Add( "Reading wallpapers from dir: " ini_wallpaper_path, MenuHandlerPath )  ; Creates a new menu item.
	
	
	TrayMenuNameDynamicSet()
	
	A_TrayMenu.Add( use_wallpaper_menu_name, MenuHandlerCurrentFile )  ; Creates a new menu item.
	
	A_TrayMenu.Add()  ; Creates a separator line.
	A_TrayMenu.Add( "Open the INI file for editing", MenuHandlerOpenINIFileForEditing )
	A_TrayMenu.Add( "Restart to load INI changes", MenuHandlerRestartAfterINIFileEditing )
	A_TrayMenu.Add( "About", MenuHandlerAbout )
	A_TrayMenu.Add()  ; Creates a separator line.
	A_TrayMenu.Add( "Exit", MenuHandlerExit )
	Persistent
	
	
	MenuHandlerTitle( ItemName, ItemPos, MyMenu )
	{
		
	}
	
	MenuHandlerExecuteNow( ItemName, ItemPos, MyMenu )
	{
		ExecuteHookNow()
	}
	
	MenuHandlerNoHookDefined( ItemName, ItemPos, MyMenu )
	{
		MsgBox "Optionally set by the INI file: `n`n" SubStr( A_ScriptName, 1, -4 ) ".ini`n`nin section: [Hook Executable]`n`nIf populated, the exe + args will be executed after each wallpaper change"
	}
	
	MenuHandlerChangeNow( ItemName, ItemPos, MyMenu )
	{
		ChangeWallpaperAndExecuteHook()
	}
	
	MenuHandlerTimer( ItemName, ItemPos, MyMenu )
	{
		MsgBox "Set by the INI file: `n`n" SubStr( A_ScriptName, 1, -4 ) ".ini`n`nin section: [Timer]`n`nwith key: Change=" ini_change_timer
	}
	
	MenuHandlerPath( ItemName, ItemPos, MyMenu ) 
	{
		MsgBox "Set by the INI file: `n`n" SubStr( A_ScriptName, 1, -4 ) ".ini`n`nin section: [Wallpaper Directory]`n`nwith key: Images=" ini_wallpaper_path "`n`nabsolute path:`n`n" use_wallpaper_path
	}
	
	MenuHandlerCurrentFile( ItemName, ItemPos, MyMenu ) 
	{
		MsgBox use_wallpaper_fullpath
	}
	
	MenuHandlerOpenINIFileForEditing( ItemName, ItemPos, MyMenu ) 
	{
		Run( my_ini_path )
	}
	
	MenuHandlerRestartAfterINIFileEditing( ItemName, ItemPos, MyMenu ) 
	{
		Reload
	}
	
	MenuHandlerAbout( ItemName, ItemPos, MyMenu ) 
	{
		Result := MsgBox( "Windows Wallpaper Rotator with a Hook`n`n`n`nThe purpose of this program is to change windows wallpaper on user-defined intervals, with, optional, exe hook to execute another program on each wallpaper change. `n`nThis is useful if you have wallpaper overlays that need to be redrawn after a wallpaper change. `n`nDeveloped by AtmanActive, 2024. `n`n`n`nWould you like to open the home page?",, "YesNo")
		if ( Result = "Yes" )
		{
			Run( "https://github.com/AtmanActive/Windows-Wallpaper-Rotator-with-a-Hook" )
		}
	}
	
	MenuHandlerExit( ItemName, ItemPos, MyMenu ) 
	{
		ExitApp 0
	}
}



TrayMenuNameDynamicSet()
{
	global use_wallpaper_menu_name := use_wallpaper_menu_prefix use_wallpaper_filename
}







;
;
; GO
;
;


; SET TRAY MENU
SetTrayMenu()



if ( ini_change_start )
{
	ChangeWallpaperAndExecuteHook()
}
else
{
	use_wallpaper_filename := "waiting ..."
	use_wallpaper_fullpath := "Waiting for the next change at every " use_timer_seconds " seconds"
	A_TrayMenu.Rename( use_wallpaper_menu_name, use_wallpaper_menu_prefix use_wallpaper_filename )
	TrayMenuNameDynamicSet()
}


SetTimer( ChangeWallpaperAndExecuteHook, ( use_timer_seconds * 1000 ) )





