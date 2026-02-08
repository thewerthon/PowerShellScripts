$DistroName = "Debian"
$WindowsUser = "Ewerton"
$LinuxUser = Read-Host "Enter username"
$LinuxPass = Read-Host "Enter password" -MaskInput
$RepoUrl = "https://github.com/olympus-app/Olympus"
$CertFile = "C:\Users\$WindowsUser\.aspnet\wsl.pfx"
$SSHID = "id_ed25519"

wsl --unregister $DistroName
wsl --install -d $DistroName --web-download --no-launch

if ($LASTEXITCODE -ne 0) { exit }

Start-Sleep -Seconds 3

if (-not (Test-Path $CertFile)) {

	New-Item -ItemType Directory -Force -Path "C:\Users\$WindowsUser\.aspnet" | Out-Null
	dotnet dev-certs https -ep $CertFile -p $LinuxPass
	dotnet dev-certs https --trust

}

$BashScript = @"
#!/bin/bash

# Setup user
useradd -m -s /bin/bash $LinuxUser
echo "$($LinuxUser):$($LinuxPass)" | chpasswd
usermod -aG sudo $LinuxUser
echo "$LinuxUser ALL=(ALL) NOPASSWD:ALL" > /etc/sudoers.d/$LinuxUser

# Setup WSL
echo "[boot]" > /etc/wsl.conf
echo "systemd=true" >> /etc/wsl.conf
echo "[user]" >> /etc/wsl.conf
echo "default=$LinuxUser" >> /etc/wsl.conf

# Setup file watcher
echo "fs.inotify.max_user_instances=1024" >> /etc/sysctl.conf
echo "fs.inotify.max_user_watches=524288" >> /etc/sysctl.conf

# Update and install dependencies
apt-get update && apt-get upgrade -y
apt-get install -y wget git python3 python-is-python3 build-essential fastfetch psmisc

# Prepare .NET SDK
wget -q https://packages.microsoft.com/config/debian/13/packages-microsoft-prod.deb -O packages-microsoft-prod.deb
dpkg -i packages-microsoft-prod.deb
rm packages-microsoft-prod.deb

# Prepare Docker
apt-get install -y ca-certificates curl
install -m 0755 -d /etc/apt/keyrings
curl -fsSL https://download.docker.com/linux/debian/gpg -o /etc/apt/keyrings/docker.asc
chmod a+r /etc/apt/keyrings/docker.asc
cat <<EOF | tee /etc/apt/sources.list.d/docker.sources
Types: deb
URIs: https://download.docker.com/linux/debian
Suites: `$(. /etc/os-release && echo "`$VERSION_CODENAME")
Components: stable
Signed-By: /etc/apt/keyrings/docker.asc
EOF

# Install Packages
apt-get update
apt-get install -y dotnet-sdk-10.0 libicu-dev docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin

# Setup Docker
usermod -aG docker $LinuxUser
systemctl enable docker.service
systemctl enable containerd.service
mkdir -p /etc/docker
cat <<EOF > /etc/docker/daemon.json
{
  "log-driver": "json-file",
  "log-opts": {
    "max-size": "10m",
    "max-file": "3"
  },
  "features": {
    "buildkit": true
  }
}
EOF

runuser -l $LinuxUser -c '

    # Setup bash
    echo "alias cls=\"clear\"" >> ~/.bashrc
    echo "export DOTNET_NOLOGO=true" >> ~/.bashrc
    echo "export DOTNET_CLI_TELEMETRY_OPTOUT=1" >> ~/.bashrc
    echo "export PS1=\"\[\033[01;34m\]\w\[\033[00m\]\\`$ \"" >> ~/.bashrc

	# Define vars
	export DOTNET_NOLOGO=true
    export DOTNET_CLI_TELEMETRY_OPTOUT=1

    # Setup git
    mkdir -p ~/.ssh
    cp /mnt/c/Users/$WindowsUser/.ssh/$SSHID* ~/.ssh/
    chmod 700 ~/.ssh
    chmod 600 ~/.ssh/$SSHID
    chmod 644 ~/.ssh/$SSHID.pub
    git config --global user.name "Éwerton Ferreira"
    git config --global user.email "eeeeewerton@gmail.com"
    git config --global gpg.format ssh
    git config --global user.signingkey ~/.ssh/$SSHID.pub
    git config --global commit.gpgsign true
    git config --global credential.helper "/mnt/c/Program\ Files/Git/mingw64/bin/git-credential-manager.exe"
    git clone $RepoUrl Olympus && cd Olympus
    git checkout dev

    # Setup AspNet
	dotnet dev-certs https --import /mnt/c/Users/$WindowsUser/.aspnet/wsl.pfx -p $LinuxPass

    # Setup NuGet
    mkdir -p ~/.nuget/local
    dotnet nuget add source ~/.nuget/local -n local
    dotnet pack Packages/Olympus.Analyzers -c Release -o ~/.nuget/local
    dotnet pack Packages/Olympus.Generators -c Release -o ~/.nuget/local
    dotnet pack Packages/Olympus.Sdk -c Release -o ~/.nuget/local

    # Restore solution
    dotnet workload restore Olympus.slnx
    dotnet tool restore
    dotnet restore

    # Finish
    find . -type d \( -name bin -o -name obj \) -prune -exec rm -rf {} +
    echo "cd ~/Olympus" >> ~/.bashrc
'
"@

$ScriptPath = "\\wsl.localhost\$DistroName\root\setup.sh"
[System.IO.File]::WriteAllText($ScriptPath, ($BashScript -replace "`r", ""))

wsl -d $DistroName -u root -- bash /root/setup.sh
wsl --terminate $DistroName
wsl ~
