
cd c:\automation\ApplicationSpecific\Tools
MSIClean32.exe {00C2D443-43D9-4550-ABEA-318288E23E57}
TIMEOUT 2
MSIClean32.exe {519B4ED1-AF5F-4812-B2A8-B18D783AEFE8}
TIMEOUT 2
MSIClean32.exe {E5AE4F66-CDA1-432A-A69E-C685D454ABDA}
TIMEOUT 2

cd %ProgramFiles(x86)%
rd Quicken /s /q 
TIMEOUT 2

rd %1 /s /q 
TIMEOUT 2

rd %2 /s /q 
TIMEOUT 2

rd %3 /s /q 
TIMEOUT 2



