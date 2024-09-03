# Driver Cabinet Creation Script
This PowerShell script automates the driver package creation process. Essentially all you have to do is go to the vendor's site and download the model's driver package. Run this script and all of the drivers will get copied over from the vendor's folder hierarchy into the UPS standardized folder hierarchy handling all edge cases. The script takes on avg 30-50 seconds to complete. Upon completion, the runtime will be displayed and a .log file can be found in the user's documents folder. The previous manual process of correctly copying the files took 2-3 hours and more than half of the times files were missed or copied to wrong location.

# FilesAndFoldersCounter Script
This PowerShell script simply returns the average of the number of files, folders, and folders directly containing files in a designated set of driver packages.

Final Note:
I did my best to redact any names/data that could be considered sensitive information to UPS. I have the unredacted versions in the gitignore. Contact me if needed.

