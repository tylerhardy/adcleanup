# adcleanup

AD cleanup script created to cleanup inactive stale computers off of a OU and sub OU's after a time of inactivity.

## Installation

Requires Powershell v3, RSAT installed on the computer running the script, and the ActiveDirectory module imported.  Edit the script domain variables to match your own.

## Usage

1. [Get-StaleADComputer] - Finds all the potential stale computers on the network inactive for 90 days and exports the computers to a .CSV file.
2. [Disable-staleADComputer] - Disables all stale AD computers that have been inactive within 90 days and export the log of the computers disabled.
3. [Delete-disableADComputer] - Deletes all disabled stale AD computers that have been inactive within 120 days and export the log of the computers deleted.

## Notes

I have based much of my script off of this blog:
https://blogs.technet.microsoft.com/chadcox/2016/06/08/my-guidance-on-identifying-stale-computers-objects-in-active-directory-using-powershell/.
As well as elements taken from these other sites (mainly for the switch confirmation in the script): 
https://technet.microsoft.com/en-us/library/ff730939.aspx.
There are various other blogs and articles that have helped me create this script, unfortunately I cannot remember them all.  Thank you to all those whom I have yet to name!