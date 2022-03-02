# stock-Analysis
**Overview of Project**
**Purpose**

The purpose of this Green-stocks Analysis projects is to refactor a Microsoft Excel VBA code to collect certain stock information in the year 2017 and 2018 and determine whether or not the stocks are worth investing. This process was originally completed in a similar format, however, the goal for this is to increase the efficiency of the original code. the data includes two charts with stock information on 12 different stocks. The stock information contains a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The goal is to retrieve the ticker, the total daily volume, and the return on each stock.

**Results**

**Analysis**
In the refactoring process, I followed the steps that were given and tried to apply the below formatted codes in a proper syntax and ran the code more efiiciently.
Below is the refactored code:

[
[refactored code for stock analysis.txt](https://github.com/Aishwaryakarthik/stock-Analysis/files/8171517/refactored.code.for.stock.analysis.txt)
](url)

Snapshots after refactoring the code 
For the years 2017 and 2018
[<img width="660" alt="VBA_challenge_2017" src="https://user-images.githubusercontent.com/99555513/156412320-5238aca6-1ae1-47b0-bff1-627084e84e8e.png">
](url)
[](url)<img width="736" alt="VBA_challenge_2018" src="https://user-images.githubusercontent.com/99555513/156412413-80782573-80ed-4044-9585-a5ad319bfae5.png">

Snapshots before refactoring the code
For the years 2017 and 2018
<img width="814" alt="runforallanalysis_2017" src="https://user-images.githubusercontent.com/99555513/156420806-9deff70d-0da2-4b34-9eac-f96b2d8d20a8.png">
<img width="771" alt="runforallanalysis_2018" src="https://user-images.githubusercontent.com/99555513/156420836-f2f96645-0190-4b2e-b7d2-7b85ee499d69.png">




**Summary**

**Pros and Cons of Refactoring Code**

•	Refactoring helps make our code cleaner and more organized.

•	 perform code refactoring in small steps. Make tiny changes in your program, each of the small changes makes your code slightly better and leaves the application in a working state.

•	 A few advantages of a cleaner code include design and software improvement, debugging, and faster programming.

•	 It may also benefit other users who view our projects because it becomes easier to read, as it is more concise and straightforward. However, we do not always have the luxury to refactor our code due to disadvantages.

•	These disadvantages may range from having applications that are too large to not having the proper test cases for the existing codes, which may ultimately pose some risk if we try to refactor our code.

**What are the advantages or disadvantages of refactoring code?**

**Disadvantages:**

•	A long procedure may contain the same line of code in several locations, you can change the logic to eliminate the duplicate lines.

•	A logical structure may be duplicated in two or more procedures (possibly via copy & paste coding). When detected, this logic is best moved to a new function and called from the other functions.

•	A complex unstructured code is usually best to split in several functions.

•	Refactoring process can affect the testing outcomes.

**Advantages:**

•	Logical errors easily appear in well structure code that contains nested conditionals and loops.

•	In our case, using Excel flow displays program logic in a more comprehensible manner, not tied to the order that the underlying code is written.

•	Advantages in Refactoring can optimize the code efficiency like we have seen in the challenge, also it can help figure out and debug the VBA code. In refactoring, duplicated subroutines, unnecessary loops, redundant statements or simply a faulty code can be removed and debugged.

The biggest benefit that occurred as a result of the refactoring was an decrease in macro run time. The original analysis took approximately one second to run, whereas our new analysis only took about a four of the time (fewer seconds) to run. 

**How do these pros and cons apply to refactoring the original VBA script?**

The cons are that refactoring the original code may lead to a more complex code that is more difficult to scale to a larger data set or may increase the risk of error. Pros apply to refactoring the original VBA script in determining how long the script took to run between the original code and the refactored code.
