# Slow_iCloud_move
Windows powershell script to move large numbers of files to iCloud one file at a time. 
The script copies a file to the target icloud directory and polls the status as it moves from sync pending to syncing to 'always available on this device.'
Once the file reaches the Always Available level - the source file is removed.  
All of this work is written to a detailed, time stamped log file as well as displayed on the console.  

# Why does this exist
As I moved my notes from Evernote to Obsidian on my PC, I noted that icould drive would simply stop working and freeze up for days with no apparent end in site and no utilities to really observe where it was stuck.  
After exhausting my search for ways to find stuck files, I decided perhaps I would just copy one file at a time because that seemed to work ok.  However, one file at a time when there are multiple directories and hundreds of thousands of individual files didn't seem like the best use of my time so I developed this concept using some assistance from ChatGPT, tested it, iterated, validated and came up with this solution that seems to be working.

Caution - this worked for my specific use case.  I have not tried this on other cloud solutions only icloud drive.  It may not work for you and given that the final step is a deletion, you may lose data. You can run 'what if' scenarios or just try one file at a time as a test to determine if it works for you. Other parameters are embeded in the file.

