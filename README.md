# staff3anno
Maps the STAFF III Database of ECG data annotations file from one line per patient to one line per file

## Details

The STAFF III database [1] provides useful ECG data for patients whose ECGs were monitored in various states, including in rooms and cath labs prior to intervention. They then had balloon inflations in specific arteries at given times and subsequent deflations. ECGs were then recorded during recovery.

The precise details of what recordings were taken and their timings vary substantially from patient to patient.

This is a very small Python script that takes the XLSX annotations file that accompanies the Staff III ECG Database and remaps it. The file supplied has a single line per patient, but there is a very variable number of files per patient with different circumstances, so there is no clear pattern in the labelling of files.

This script simply produces a new XLSX with a single line per data file which provides data per file on MI location (if any) and balloon inflation times etc.

The script is not useful without the Staff III database itself [1].

You need Python 3.6+ and the openpyxl library to run this script, both of which are freely available.

[1] https://physionet.org/content/staffiii/1.0.0/
