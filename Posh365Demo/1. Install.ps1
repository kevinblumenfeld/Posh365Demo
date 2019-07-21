
# Install Posh365. Requires Win10 or Win2016
Install-Module Posh365 -Force -Scope CurrentUser -Confirm:$false

# Remotely install Posh365 on older OS. Requires access to remote computer's C$
Install-Module Posh365 -Force -Scope CurrentUser -Confirm:$false
Install-ModuleOnServer -Server DC01

# Problems? Run 'powerShell as administrator' and run this:
Set-ExecutionPolicy RemoteSigned -Force

