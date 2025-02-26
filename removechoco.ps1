# Remove o Chocolatey
Remove-Item -Path "$env:ProgramData\chocolatey" -Recurse -Force
Remove-Item -Path "$env:ALLUSERSPROFILE\chocolatey" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$env:SystemDrive\ProgramData\chocolatey" -Recurse -Force -ErrorAction SilentlyContinue
exit 0