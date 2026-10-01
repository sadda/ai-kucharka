# Run convert.py with specified venv and output files
# Additional arguments are passed to convert.py, e.g. .\convert.ps1 --config configs\company123.toml ZAK001
$arguments = @("-u", "convert.py") + ($args | ForEach-Object { "`"$_`"" })
Start-Process -WindowStyle Hidden `
  -WorkingDirectory $PSScriptRoot `
  -FilePath "venv/Scripts/python.exe" `
  -ArgumentList $arguments `
  -RedirectStandardOutput "output.log" `
  -RedirectStandardError "error.log"

# Get-Process python | Stop-Process -Force
