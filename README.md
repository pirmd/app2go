# INTRODUCTION
App2go is a small framework to build and run so-called "portable windows
applications", usually from a USB drive you carry out with you here and there.
Well-known portable apps suites can be found on the web (like
[WinPenPack](www.winpenpack.com), [FramaKey](www.framakey.org) or
[PortableApps](portableapps.com)).

This toolkit does not pretend to replace them (even if it most probably can
with further debugging and polishing), it was mainly developed to:
- solve small annoyances and shortcomings of existing suites I experienced for
  my use case (like cannot easily share programs or service like web server
  between apps);
- allow to quickly set-up new portable apps that are not available in the main
  suites or at least not available in a flavor I like (notably
  gVim/ReST/docutils combo, unison/ssh, cygwin, python...);
- escape from the NSIS/Autoit scripts with unreadable non modularized code and
  over-verbose INI files (at least according to my feeling) which I have
  difficulty to maintain when I urgently need to update or modify the behavior
  of one PortableApp;
- have basic logging/debugging information;
- have fun.

# DESCRIPTION
App2Go is a simple set of vbs scripts therefore lightweigth and easily hackable
and extensible.

The tool is organised in three parts:
- _a core script_ that run a kind of poor-man state machine that set-up the
  portable environment, launch shared resources if not already done, clean-up
  afterwards. It provides:
  * the main loop allows to implement in a vbs way a kind of try...finally loop
    (or a bash trap) to make sure that cleaning actions are run whatever
    happen;
  * two task queues, one to run the actions and the second to run the
    cleaning-actions once everything is finished (either normally or in case of
    errors);
  * a way to manage shared resources between several app2go apps that allows to
    keep shared resources (like a http server, or a 'subst"-ituted drive name)
    alive if there are apps still claiming them.

- _a library_ providing useful functions for setting-up a portable environment.
  The library adopts a (hopefully) simple API to perform some kind of
  high-level tasks while hiding the details for easy configuration writing. The
  set of functions also take care of:
  * automatically adding cleaning/falback actions so that they don't need to be
    identified within the configuration stage.
  * automatically registering/releasing shared resources among several app2go
    apps
  * logging messages to inform what is going on It is not mandatory to use
    functions from the library when creating a new app2go apps, you can use any
    vbs code, it is just supposed to be easier. Short API list is found [here
    after] (URL "API").

- _a set of "configuration" files_ that are run by the app2go core script to
  launch a portable app. They are in fact directly vbs scripts and give
  therefore a good flexibility to advance user and remain easy to configure for
  someone without any scripting knowledge (at least as easy as any portable
  apps suite INI file out there) Configuration files can be chained using a
  "include" mecanisms that help defining standard and common customization
  (like $HOME, $LANG, app2go standards file organisation). It allows also some
  app2go apps to offer or record services to other app2go appas thta just need
  to include such configuration if it exists. Some fun things can also be
  achieved like running a configuration files depending on the Host-computer
  name or some even fancier criteria.

- _a relative-shortcut tool_ that simplifies the launch of an apps2go
  configuration script. It also allows to have a nice (and small) exe with nice
  icons and whatever very important metadata you want people to know off (like
  a fancy war-name, a copyright, a trademark, a highly complex and meaningful
  version coding scheme). At the time being it is achieved throug a dead-simple
  full of bugis C-based code that looks for a configuration script with the same
  base name in a predertermined set of folder relatove to the shortcut.


# EXAMPLES
Example of use of this framework is availbale in the package 'config.example' folder
with:
- portable versions of cygvim (setup, shell, vim), firefox, qemu
- all of these programmes have access to common resources (useful tools
  available by cygwin, python, installed python packages)
- file structure is completly tweakable using the configuration files app2go.2go so that
  you can organize your stuff as you like.
  The example struture is:
    * simply drop firefox / cygwin runtime folder in the same directory than the configuration files
    * configuration files are looked in the $HOME folder (like on Linux)
- fun but maybe not useful things I can do:
    * Identify the USB key by its name and not only by extracting the drive
      name containing the portable apps. It is useful for example if you have
      several partitions on your drive to separate documents from tools
    * run Mozilla addon-sdk with firefox portable (access to python, tweak
      PATH, portable version of Firefox which allow to path a user defined
      profile name)
    * debug my apps2go with a verbose logging to the console

# API
Available functions from the framework are:
* _Methods_:
        AddCmdLine
            : concatanate two parts of a command line

        AllCmdLineOption
            : add a command line option only if not already present (usefull to
            define overridable command line default in config file e.g. firefox
            profile)

        SanitizeCmdLine
            : Add additional " to take care of arguments with space that can
            confuse WShell.Run

        SetEnv
            : define an environemnt variable

        AddToPATH
            : add to the PATH environment variable

        DefaultTo
            : set a default value for a variable

        Link
            : create a junction, shortcut or a subst

        Shortcut
            : create a shortcut on host Desktop

        AddToDesktop
            : create a link on the host desktop

        AddToSendTo
            : create a link in the 'Send To' right-click menu

        LoadScript
            : load an app2go script from a file

        LoadScriptIfExists:
            : load an app2go script from a file if it exists

        MergeHiveFile
            : merge a hive reg file

        SetRegKey
            : modify a registry key

        GetDriveByVolumeName
            : Get the drive name (H:, P:,...) of a disk given its volum name

        GetPhysicalDrive
            : Get the physical drive Id of a disk given its volume name


        FixDriveLetter
            : Replace %LauncherDrive% with the drive letter from which the prog
            is launched

        FixUserFolders
            : Replace in a text file any reference to user folders by their
            actual value. In that order:
                - %DocFolder% by the folder name that contains the user
                  documents
                - %RootFolder% by the folder name that contains the App2go
                  platform
                - %LauncherDrive% by the drive letter that contains the
                  launcher

        ReplaceInFile
            : replace a string in a text file

        Run
            : execute a commande (WshShell.Run wrapper) programs launhed are
            automatically killed when the launcher is stopped.

        RunOnExit
            : execute a commande only when the launcher exit

        RunService
            : execute a command as a shared 'service' for all lauchers

* _Constants/variables_:
    LauncherDrive
        : drive name of the launcher

    LauncherFolder
        : folder name of the launcher

    LauncherParentFolder
        : parent folder of the launcher

    MyDocsOnHost
        : location of "My Documnets" on host

    DesktopOnHost
        : location of "Desktop" on host

    SendToOnHost
        : location of "SendTo" on host

    AppDirOnHost
        : location of "Application data" on host

    TempOnHost
        : location of 'Temp'dir on host

    HostName
        : Host Computername


# EXTENDING
Details will be documented later but basic idea is to write some vbs functions following one of the existing one in lib2go.vbi

See modules in 'libgo' folder which should be straight forward

# TODO
- Improve a lot the documentation
