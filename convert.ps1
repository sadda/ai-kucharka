# Run convert.py with specified venv and output files
Start-Process -WindowStyle Hidden `
  -WorkingDirectory $PSScriptRoot `
  -FilePath "venv/Scripts/python.exe" `
  -ArgumentList "-u convert.py" `
  -RedirectStandardOutput "output.log" `
  -RedirectStandardError "error.log"

# Get-Process python | Stop-Process -Force