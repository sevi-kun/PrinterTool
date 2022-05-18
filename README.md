# PrinterTool
A PowerShell tool to manage printers

![PrinterTool Preview](https://github.com/sevi-kun/PrinterTool/blob/main/preview.png?raw=true)

I wrote this tool for my company so it can be used as end-user tool.

## Configuring
Edit the config.xml file to fit your needs.

#### config.xml
You will find two lists in the config.xml file, you may have to change.
1. Printserver: 

    Add all your printservers you wish need. From there the remote-printer list will be generated.

2. Blacklist:

    If you have some Printers you don't want to show in the remote-printer list, you can add them here.

#### main.ps1
If you store your config.xml not in the same folder as the main.ps1 you may want to change the variable on line 18.

## Troubleshooting
* If you want to start the tool from a network-share, 
you have to enter the absolute path on line 18 in the main.ps1.
