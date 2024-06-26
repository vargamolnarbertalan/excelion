# Excelion by B3RC1

## What is this?
A web controller for Microsoft Excel. I would not call this an API, since you have to write the functions you want to use on the Excel side too. 
The purpose of this app is to modify data inside Excel through network requests. This enables you to control Excel from third-party applications such as Bitfocus Companion.

## Install
- Download the latest release and run the installer.
- Upon completing the installation proccess run the Excelion app that is now added to your program list in Windows.
> Please note that by doing this, you won't be able to change the default port.

## Or compile
- Install Python for Windows
- Install dependencies:
```
    pip install Flask python-dotenv xlwings
```
- Compile and run python script:
```
    python app.py
```
> Please note that by doing this, you can change the default port anytime.

### Alternative run
Start the **run.bat** file found in the application's directory after installing python and dependencies.
> Please note that by doing this, you can change the default port anytime.

### Port number
You can set your preferred port number for the app to run on, just change it in the **.env** file. *Default is 15000*

## Usage
The only way to interact with the application is to send HTTP GET requests to it.

### Request template:
**http(s)://server_ip:server_port/modify_excel?filePath=ARG1&scriptName=ARG2**
- Where **ARG1** is the **path with filename and extension** in Windows format (\\\ between directories and filename). Should be something like this: c:\\\path\\\to\\\document.xlsm
- And **ARG2** is the **VBA script name** that is declared in the document - without the () - something like this: MyFavoriteScript

*Example request:*
*http://192.168.1.64:15000/modify_excel?filePath=c:\\path\\to\\document.xlsm&scriptName=MyFavoriteScript*

## Good to know
- Works both with .xlsx and .xlsm files.
- Keep in mind that right now this app only handles VBA methods that do not take parameters.
- Windows only.
