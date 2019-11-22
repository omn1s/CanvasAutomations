# CanvasAutomations
A set of Powershell utilities designed for the administration of the Canvas LMS.

# Overview
Primarily, I wrote these functions to assist with SIS Data import, so this code assumes you have your SIS exporting users/terms/courses/sections/enrollments already. I've included a few extra functions just for monitoring SIS importing that may not need any access to the data CSVs to be useful.
Users of this program will need to write a script or function that actually DOES what you'd want to happen, because it's designed as just a bunch of useful functions. At the bottom of the script, you'll find an example of how I actually use these functions.

# Usage
* Download CanvasUtilities.ps1 and Config.xml
* Make key files for your test and production Canvas instances
* Create a csv file and enter the terms which you'd like to batch import enrollments against (you only need a column called Id, then the Canvas Ids of the applicable terms)
* Edit Config.xml
  * Enter the locations of your key files in their respective tags
  * Enter the location of the terms file for batch importing
  * Put "prod" or "test" (no quotes) in the application level tag
  * Enter the location of your SIS export in <SISDataDirectory>
  * Enter the base URLs (including the account number and /) of your test/prod instances (https://canvas.x.x/api/v1/accounts/1/)
* Launch Powershell
* Navigate to the location of the downloaded script and type ./CanvasUtitilities.ps1 and press enter.
* You should now be able to run any of the functions present. I'd start with Get-CanvasTerms for a quick confirmation that all is working. 
* Check out the last function, Start-MyCanvasImports, which uses all of the functions present. We use it to do all of our general and batch imports (enrollments / term). 

# Notes
* We run this on a server with Windows Powershell 5, so I had to write my own Invoke-WebRequest with retries (version 6 has a retries option built in)
* I don't believe Canvas allows you to run more than one import at a time, so an asynchronous approach to solving this problem appears useless. If you know how multi_term_batch_mode works in the Canvas API, helping me understand it would be wonderful.

# Contributing
* Feel free to contribute, particularly if you have a solution to multi_term_batch_mode. 
* Feel free to edit, share, or use however. 
