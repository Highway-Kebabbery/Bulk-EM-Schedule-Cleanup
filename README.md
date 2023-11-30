# <div align="center">Bulk Environmental Monitoring Schedules Cleanup</div>


## Description
### Overview
This Python program loads the specified Excel file containing schedules for an environmental monitoring program and cleans it up by finding and deleting all entries for outdated versions, as well as making it visually easier to parse by adding blank rows between different environmental monitoring schedules.

I built this project because, in the process of assigning me a task, my manager gave me an Excel .csv containing over 20,000 rows of data about our existing environmental monitoring programs, of which less than 2,000 were relevant. The export contained data for all previous versions of each schedule. I needed to reduce the data set to only the current versions, and I didn't feel like processing all 20,000 rows manually. This program completed in a few hours a task that would've taken me days to complete manually.

### What I Learned
The main challenge with this project was that I couldn't visually see what my program was doing in Excel, so I had to implicitly figure out several of Excel's behaviors through trial and error in order to debug my program as I built it. One issue (which is a bit embarrassing) turned out to be that I didn't realize that the active cell remains the same when you insert or delete rows in Excel.

I also had to implicitly figure out that deleting lines individually is a far more time-intensive process than performing a bulk deletion. After that discovery, I reworked the program to store the row number of all rows to be deleted in a temporary list and to perform deletions in bulk after processing a chunk of outdated schedule versions.

I would never take home the proprietary company data on which this program operated (see [Why I'm Showcasing this Work](#why-im-showcasing-this-work) for more on this), so I've left it as it was when it last functioned (note that one comment makes reference to code I removed at the time of project completion). However, if I were to go back and update it, I'd want to re-write it from an object oriented standpoint for practice. I was just learning about object oriented programming when I completed this project in May of 2023.

I would: avoid using global variables; instantiate an object containing the open workbook, the column letters, the current row number, etc. as attributes; and define the function and the main while-loop as methods that interact with the aforementioned attributes through setter methods.

### Why I'm Showcasing this Work
I wrote this program on company time. However, it is in no way related to any product that my former employer sells. It is essentially a single-use tool written by me to solve a specific problem they will likely never encounter again, and a problem which they anticipated paying me to solve manually. I don't realistically believe that it is of value to the company or to anybody else any longer. Nothing in this script reveals anything related to company intellectual property (aside from the script itself). The only remaining value of the script is to document that I applied programming knowledge to solve a real problem, that I saved my employer thousands of dollars, and that I saved myself days to weeks of time by executing this project.


## Features
The sole purpose of this script is to take an Excel .csv file exported from a company's instance of LabWare LIMS containing all versions of all environmental monitoring sampling schedules ever implemented at their site and to clean it up so that only active schedules (the highest version of each schedule) remain. It also adds a blank row between schedules for human readability.

It was not necessary, but I had the program take some input through the command line because I find programs that are interactive to be more fun and satisfying.

The program renames the file and saves it as a new file to preserve the parent copy.


## How to Use
### Software Requirements
* Python 3
* Microsoft Excel

### Instructions
1. Run the program
2. Follow the prompts in the terminal.
3. Press enter to run.
4. The terminal will print the list of rows awaiting deletion as the program runs.
5. The program will inform you in the terminal when it has completed.

## Technology
* **Python 3:** I chose Python 3 because it's what I was learning at the time. However, it is also a powerful and user-friendly language that made it easy to approach this project.


## Collaborators
I completed this project on my own.