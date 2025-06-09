# Calibrate
A PowerShell script kit to run nightly to calibrate RBN skimmers

Calibrates CW (CW Skimmer Server) and digital mode (CWSL_DIGI) skimmers based on skew
data published on [sm7iun.se](https://sm7iun.se/rbn/analytics/)

Since the calibration adjustment is proportional, the script should be run no 
more than once per day and preferrably around 00:30 UTC, directly after the content 
of the web page with skew information has been updated.

To allow execution of the dailycalibration batch file from Windows scheduler, 
run setexecutionpolicy.ps1 in a PowerShell window as administrator. 
