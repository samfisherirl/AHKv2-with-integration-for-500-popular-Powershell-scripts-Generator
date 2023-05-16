# AHKv2 with integration for 500 popular Powershell scripts + Generator
AHKv2 with integration for 500 popular Powershell scripts + Generator
# PowerShell Script Runner with AHK GUI

This repository contains an AutoHotKey (AHK) application that allows you to import and run PowerShell scripts directly from a script list. The application provides a GUI interface to write AHK scripts that execute PowerShell commands from the imported scripts, without the need to store the scripts in separate PS1 files. This approach helps to avoid Windows Defender warnings associated with executing PS1 files directly.

## Features

- Import and manage a collection of PowerShell scripts.
- Write AHK scripts that utilize the imported PowerShell scripts.
- Execute PowerShell commands directly from the AHK GUI.
- Seamless integration of AHK and PowerShell for automation purposes.

![image](https://github.com/samfisherirl/AHKv2-with-integration-for-500-popular-Powershell-scripts-Generator/assets/98753696/0c173eaf-7c76-49cf-8c17-f939aa3a624d)

## Requirements

- AutoHotKey (AHK): The AHK interpreter is required to run the application. You can download it from the official website: [AutoHotKey Downloads](https://www.autohotkey.com/download/)

## Usage

1. Clone or download this repository to your local machine.
2. Install AutoHotKey if you haven't already done so.
3. Open the AHK script file (`PowerShellScriptRunner.ahk`) using the AutoHotKey interpreter.
4. The application will launch, displaying the GUI interface.
5. Use the "Import Scripts" option to add your desired PowerShell scripts to the application.
6. Write AHK scripts in the provided text area, utilizing the imported PowerShell scripts.
7. Run the AHK script by pressing the assigned hotkey or using the "Run" button in the GUI.
8. The AHK script will execute the selected PowerShell commands from the script list.

## Example Output

```autohotkey
ps_path := A_ScriptDir "\ps_script.txt"
FileAppend(PS_Script(ps_path), ps_path)
RunWait(command())
FileDelete(ps_path)

command() {
return "powershell -noexit -ExecutionPolicy Bypass -Command `"& { $scriptContent = Get-Content -Path 'ps_script.txt' -Raw; Invoke-Expression -Command $scriptContent }`""
}

PS_Script(ps_path) {
	if FileExist(ps_path) {
		FileDelete(ps_path)
	}

	return "
	(
	param([string]$Message = "sca")
	try {
		if ($Message -eq "" ) { $URL = read-host "Enter alert message" }
		echo "ALERT: $Message"
		curl --header "Access-Token: o.PZl5XCp6SBl4F5PpaNXGDfFpUJZKAlEb" --header "Content-Type: application/json" --data-binary '{"type": "note", "title": "ALERT", "body": "$Message"}' --request POST https://api.pushbullet.com/v2/pushes
	} catch {
		"?? Error in line $($_.InvocationInfo.ScriptLineNumber): $($Error[0])"
	}
	
	)"
}
```

## Notes

- It's recommended to review the imported PowerShell scripts for security and compatibility before running them.
- The AHK script runner provides a convenient way to execute PowerShell commands without storing them in separate PS1 files, reducing the chances of triggering Windows Defender warnings.
- Make sure to adjust your system's execution policy if required to allow the execution of PowerShell commands.

## Contributing

Contributions are welcome! If you have any ideas, bug fixes, or feature enhancements, feel free to open an issue or submit a pull request.

## License

This project is licensed under the [MIT License](LICENSE).
The 500 powershell scripts were sourced here: https://github.com/fleschutz/PowerShell/blob/master/LICENSE
