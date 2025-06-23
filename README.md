# Calibrate
A PowerShell script kit to run nightly to calibrate [RBN](https://www.reversebeacon.net/main.php?rows=10&max_age=10,hours&hide=distance_km) skimmers.

Calibrates CW (CW Skimmer Server), RTTY (RTTY Skimmer Server), and digital mode (CWSL_DIGI) 
instances based on skew data published on [sm7iun.se](https://sm7iun.se/rbn/analytics/)

Since the calibration adjustment is proportional, the script should be run no 
more than once per day and preferrably around 00:20 UTC, shortly after the content 
of the web page with skew information has been updated.

To allow execution of the `dailycalibration.bat` batch file from Windows scheduler, 
you may need to run `setexecutionpolicy.ps1` in a PowerShell window as administrator. 

`testrun.bat` and `testrunv.bat` (verbose) are test scripts that runs without updating any files. 
