# adcleanup

AD cleanup script created to cleanup inactive stale computers off of a OU and sub OU's after a time of inactivity.

## Installation

Requires Powershell v3, RSAT installed on the computer running the script, and the ActiveDirectory module imported.  Edit the script domain variables to match your own.

## Usage

1. [Get-StaleADComputer] - Finds all the potential stale computers on the network inactive for 90 days and exports the computers to a .CSV file.
2. [Disable-staleADComputer] - Disables all stale AD computers that have been inactive within 90 days and export the log of the computers disabled.
3. [Delete-disableADComputer] - Deletes all disabled stale AD computers that have been inactive within 120 days and export the log of the computers deleted.