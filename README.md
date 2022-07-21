# OutlookNoteToOneNote
writes your Outlook notes to OneNote

DISCLAIMER
this script send inputs around, so a lot of things can go wrong
perform at own risk

Guide

A CSV file with the notes is needed.
File->open and export->import/export->as file->csv(comma seprated values)->select your notes->where it is saved->confirm
can take some time

before you run the script you should change some values in it (you can't run a powershell script by default anyway)

$csv is the path to your csv file

$text is name for the column that contains the notes
open your csv file with the editor or something similar

your first row will probably look like this
"column1","column2","column3","column4","column5"

set $text to the name of the column that contains the notes

$text = "column1"

$windowname is the window name of the programm
leave it as "onenote for windows 10" when you use that or change it to "onenote" when using the normal version
(you can also use other programm but do it at your own risk)

$badChars contains special charaters that are/might get interpreted wrong.
e.g. ~ results in the enter button being pressed

open your onenote(for windows 10) 

copy the script and paste it into the powershell
(or run it if you disabled the policy)




