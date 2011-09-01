Remove Non-present Devices using DevCon
=======================================

removedevices.js is a script to automatically remove all non-present (i.e.
disconnected) devices from a Windows computer system.  This can often be
useful to prevent misbehaving and/or unnecessary drivers from being loaded.

WARNING
-------
This script will remove/uninstall devices from your computer system.
Although several important types of devices have been excluded from the
default removal process, the possibility of removing devices critical to
the operation of your computer system still exists.  PLEASE double-check
the list of devices which will be removed before pressing "Y" and PLEASE
create a system backup if you are remotely concerned about making your
system unbootable.

Examples
--------
 - To remove all non-present devices (except legacy and software devices)
   either double-click the script or invoke it from the command prompt
   without options.

 - To delete devices without confirmation, run

        removedevices.js /noconfirm

 - To see the output of DevCon as the script executes, run

        removedevices.js /verbose

 - To create a list of devices which would be removed, run

        removedevices.js /outfile:devicelist.txt

 - To delete all devices (device IDs) listed in a file, run

        removedevices.js /infile:devicelist.txt


More Information
----------------
Installation instructions are available in INSTALL.txt.
Major changes are listed in ChangeLog.txt.
Complete license text is available in COPYING.txt.

The latest version of this software and the bug tracking system are
available on GitHub <https://github.com/kevinoid/remove-nonpresent-devices>.
