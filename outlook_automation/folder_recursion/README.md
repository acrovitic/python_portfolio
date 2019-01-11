I read that there was no easy way to recursively travel through Outlook folders. For example, I was unable to find a function or
library to allow me to find all subfolder names of a folder, the subfolders of THOSE subfolders, and return to execute the same task 
for the next main folder target, without writing a long and ugly if-else statement.

To address this issue, I created the two functions found in outlook_folder_loop_functions.py.

The outlook_depth_detector will find the deepest subfolder of an outlook folder. 

The outlook_depth_pathfinder will retrieve the formatted folder paths of all subfolders at all depths.

These functions can likely be modified to suit multiple types of recursive activities in Outlook.
