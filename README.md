# STOCK ANALYSIS CHALLENGE IN VBA

## OVERVIEW AND PURPOSE OF PROJECT

### OVERVIEW

We worked on a stock VBA project to help, Steve, determine the most desirable stocks for his parents to choose.  The information was provided for the years 2017 and 2018, and included buttons to run the analysis, select the year, highlight positvit or negative outcomes, and a clear button to start over.

### Purpose
The purpose of the project is to take what we had previously done to for Steve and improve upon it.  The improvement process is called refactoring.  The goal of refractoring is to add/delete lines of code to increase efficiency and readibility of the previous version.  The efficiency will be the easiest to measure as it is time based.

## RESULTS

### ANALYSIS
Here are the results for 2017 and 2018 for the twelve stocks selected:

<img width="239" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/91889241/147021968-33c76d56-f7fc-4b8c-b4a9-d2172f2ced8c.png">
<img width="234" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/91889241/147021980-e5a97b06-1ac6-41fc-9781-bde7dc102f9a.png">

These images show only one stock, RUN, increased its return in 2018 from 2017.  All the other stocks took a massive hit into negative returns, other than ENPH.  ENPH still remained as a positive return in 2018, however, less than in the previous year.  We can also see the effect of returns and slowdowns in daily volume amongst most of the stock downturns, with a few exceptions, TER and VSLR.

### CODE
We only used twelve stocks for this example and hard coded them into our VBA.

›<img width="182" alt="Screen Shot 2021-12-21 at 8 04 44 PM" src="https://user-images.githubusercontent.com/91889241/147022657-b2ca2eb3-1856-4151-88fb-b1a55a7355e2.png">

If we to make any changes to these stock choices or selections, we would have to return and change the values or add them and increase our string count.  This was a limitation of the data set and what we were trying to accomplish.

From our first project, we created six subroutines to formulate, format, and output our results.  This takes additional time for VBA to run.  See the results below for the first project run time;

<img width="267" alt="Original 2017" src="https://user-images.githubusercontent.com/91889241/147024511-3849a4e2-8399-488b-ab5f-011857403602.png">
<img width="257" alt="Original 2018" src="https://user-images.githubusercontent.com/91889241/147024517-e67f7f9c-1183-4123-9d87-04e60bd27909.png">

After eliminating all extraneous subroutines down to one subroutine, the time results are:

<img width="251" alt="Refactored 2017 results" src="https://user-images.githubusercontent.com/91889241/147023863-4305dd8c-c13a-4fe0-bb4d-ef31b395752d.png">
<img width="262" alt="Refactored 2018 results" src="https://user-images.githubusercontent.com/91889241/147023865-5094a835-a9f1-4192-9d06-bfbf2d6ed921.png">

The only time savings was marginal at best with the 2018 dataset.  However, as the data set gets larger, the fewer subroutines should see exponentially lower run times.

## SUMMARY

### ADVANTAGES/DISADVANTAGES OF REFACTORING CODE

Code refactoring is the process of restructuring exisitng computer code without changing its external behavior.  

Advantages is a cleaner and more readable code.  This can aid in faster and more accurate running of the program.  This readability allows users, other than the creator, to understand the intent and purpose of the script.  Furthermore, we can garner patterns in the design of the code and make judgments as to continuing the design or making improvements throughout the entire script.

Disadvantages can come in several forms of time.  It can simply be a time constraint to perform refactoring on existing code.  If another project awaits, "it's working well for now" is a common slogan to keep moving forward and not look back.  Also, cost benefit of this time.  If a set of code is small or performs a limited function within a set of acceptable time limits, refactoring would more than likely not be beneficial.

### ADVANTAGES/DISADVANTAGES OF STOCK ANALYSIS CODE


