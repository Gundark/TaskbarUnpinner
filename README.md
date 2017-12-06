# Taskbar Unpinner

Experimental code using VBScript to unpin unwanted default shortcuts.

Based on Stefan Hofbauer's code from https://pinto10blog.wordpress.com/2016/09/10/pinto10/
originally posted May 22, 2017, retrieved Dec 1, 2017

Notes:

1.  Shortcuts to pinned apps are stored in 
    C:\Users\name\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar

2.  If shortcuts come back at login due to a provisioning package, add a shortcut to this script to run
    at login time:  https://www.computerhope.com/issues/ch000322.htm

3.  The sample code that inspiring this uses the namespace "shell:AppsFolder".  It turned out this didn't work 
	for either IE or File Explorer.  Neither do "C:\Program Files (x86)", or "C:\Program Files", because the verb
	list doesn't contain the 'Unpin from taskbar' verb.
	
4.  On my laptop, if I pin Internet Explorer to the taskbar, the 'shell:AppsFolder' namespace works.
    But for the default pinned IE icon created by my company's provisioning package at login time, it doesn't work
    because the 'unpin from taskbar' verb is not available, either through the GUI or the shellFolderItem.Verbs() 
	method, even though IE is visibly pinned.

5.  In the end, I had to use two different namespaces to pick up all of the programs; there was no
	single namespace that worked for all.

