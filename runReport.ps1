Remove-Item -Path .\myoutput.txt
#.\Get-HyperVReport.ps1 -Cluster P1-CL01
.\Get-HyperVReport.ps1 -VMHost S2D-SERVER01,S2D-SERVER06
python sum.py
