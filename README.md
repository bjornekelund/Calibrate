# Calibrate
A PowerShell script kit to run nightly to calibrate [RBN](https://www.reversebeacon.net/main.php) skimmers.

Calibrates up to two instances of CW Skimmer Server and RTTY Skimmer Server each plus one instance of 
CWSL_DIGI based on skew data published nightly on [sm7iun.se](https://sm7iun.se/rbn/analytics/)

Since the calibration adjustment is proportional, the script should be run no 
more than once per day and preferrably around 00:20 UTC, shortly after the content 
of the web page with skew information has been updated.

`dailycalibration.bat` is intended to run nightly using Windows Task Scheduler and also creates a log file. 

When creating the task in Windows Task Scheduler, check the box "Synchronize across time zones" to make the time UTC.

Depending on your PC's policy settings, you may need to run the commands in `setexecutionpolicy.txt` in a 
PowerShell window as administrator to allow execution of the script files, 
including `dailycalibration.bat` by Windows Task Scheduler.

`testrun.bat` and `testrunv.bat` (verbose) are test scripts that runs without updating any files. 
