# eso-student-rep

you can create a randomized schedule of students from the file names.txt (one name per row).

I updated the script to assign three people per 2 week block, I guess this makes it easier. To get a schedule from a given monday on, run this:
 
 
``` ./schedule_YBX.py -names students.txt -start 2025-06-30```

 
If the date is not a Monday, the next Monday will be taken as a starting date.
