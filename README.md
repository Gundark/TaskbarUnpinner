# Taskbar Unpinner

Experimental code using VBScript to unpin unwanted default shortcuts.

Based on Stefan Hofbauer's code from https://pinto10blog.wordpress.com/2016/09/10/pinto10/
originally posted May 22, 2017, retrieved Dec 1, 2017

Notes:

1.  Shortcuts to pinned apps are stored in 
    C:\Users\name\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar

2.  If shortcuts come back at login due to a provisioning package, add a shortcut to this script to run
    at login time:  https://www.computerhope.com/issues/ch000322.htm


3.  On my work laptop, if I pin Internet Explorer to the taskbar, this program works fine to remove it.
    But for the default pinned IE icon created by the company provisioning package at login time, it doesn't work
    because in the context used by this program (Shell:Applications), the 'unpin from taskbar' verb is not available,
    either through the GUI or the shellFolderItem.Verbs() method, even though IE is visibly pinned.

4.  Also on my work laptop, this program doesn't File Explorer.  Unlike IE, File Explorer doesn't have the
    "Unpin from taskbar" verb, or the "Pin to Taskbar" item in the GUI context menu when in the Shell:Application
    context.  Oddly enough, it does have it when in the Start menu context; perhaps using a different context
    for finding pinned apps would solve this problem.

5.  It's possible that the provisioning package used by my company isn't set up correctly for File Explorer
    and Internet Explorer.
	
6.  It may be that there is a way to add the missing verb to the File Explorer application in the
    shell:Application context, since it has the verb in the GUI when using the Start menu.  Need to
	determine why the two contexts are different.

7.  The behavior of IE and file Explorer need more research.  This page may contain clues:
    https://community.spiceworks.com/topic/549865-vbs-script-for-pin-and-unpinning-taskbar-icons
	
8.  When using the debug option, note that the 'unpin' verb will not be present when an app is not currently pinned.

