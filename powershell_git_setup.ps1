<#
    .Description Setting PowerShell ISE up for GitHub

    Resources: 
    https://github.com/dahlbyk/posh-git
    http://haacked.com/archive/2011/12/19/get-git-for-windows.aspx
    http://haacked.com/archive/2011/12/13/better-git-with-powershell.aspx
#>

# 1. Install GitHub for Windows and msysgit
# http://msysgit.github.com/

# 2. Add git alias and the msysgit/bin folder to your PATH variable to PowerShell profile
New-Item -path alias:git -Value 'C:\msysgit\cmd\git.exe' #This is where mine installed.
$env:path += ";C:\msysgit\bin" #This is where mine installed.

# 3. Clone Posh-Git to your Repository from the web interface, button at the top
# https://github.com/dahlbyk/posh-git

# 4. Run the install.ps1 script
cd .\Documents\GitHub\posh-git
.\install.ps1

# 5. Move the code from Profile.Example.ps1 to your Profile

# Restart PowerShell

# Future Usage
# cd to GitHub directory
git help # for available commands

git status

# For a new Repo, or to reinitialize an existing one
git init
git fetch origin

# To upload files
git add <name>

# To commit files to stage
git commit -m "description"

# To update GitHub
git push origin master
