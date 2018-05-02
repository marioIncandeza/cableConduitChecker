# Conduit and Cable Schedule Checker

This script will backcheck the contents of the "CONDUIT NO." (x2), "CABLES INCLUDED", and "CABLE NO." columns and make sure their contents match on both the Cable and Conduit Schedule. 

## Install
To run the script, first you'll need to install [Node.js](https://nodejs.org/en/) on your machine. Node.js is a runtime environment that allows you to run JavaScript outside of your browser.
Follow the link and download the recommended version. The link should automatically point you to the right version for your operating system. Once the download finishes, run the installer and click through all 
of the default options.

Once you've installed Node, you'll need to download the repository. The most direct way to do this is to download the .zip file. Look for the big green button near the upper-right. Click it. Select "Download ZIP".
Now unzip the folder and put it wherever. Now you're going to have to navigate to that folder using the Command Prompt. So go ahead and open up the command prompt (press that windows button or just type `cmd` into that box in the lower-left corner of the screen). Once you're there, you can check the present working directory by typing in `dir`. To change directories, use the `cd` command. 

For example, if I unzipped the folder on my Desktop, I'd do the following:
```
cd C:\Users\lae\Desktop
````

Once you're in the right directory, all you have to do is run 'npm install', and you'll be all ready to go.

## Running the Script
The first thing you'll have to do is point the script to the local copy of the Conduit and Cable Schedule. Open up ccsCheck.js and change the filepath on line 1 to the path of the schedule you're trying to check.
It'll look like this:
```
var projectSchedule = 'C:\Path\to\your\file.xlsx';
``````````````````````````````````````````````````````
Remember to save your changes.

Now to check your schedule, just type `node ccsChecker.js` or `npm start` in the command prompt.

The output will just be an array containing problematic entries in the Schedule. The program starts by tracing the path of each cable through its respective series of conduit(s). Then the 
program will move onto the conduit schedule and confirm the included cables.

By default, the script is pointed toward "trial.xlsx" which has some intentional mistakes in it to help you interpret the output array. Ideally, if there are no mistakes in the columns under consideration, the array returned by 
the script will be empty.
```
[] \\ you're good to go!
````````````````````````````````````````````````````````