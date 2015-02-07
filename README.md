# MarkingAid
An Excel based project that makes marking easier with highlighting and pre-made functions as well as a VBA backend to help automate aspects of the marking process

The project is mainly designed to aid teachers mark more quickly and more efficiently and was created to aid my father, a drama teacher, when entering his students' test scores. I realised that this could be a useful program for others so have decided to release it, after some cleaning up, on GitHub.

This README will start with the basics of using the program. The conditional formatting, formulae and VBA should be fairly self-explanatory so will not be covered but if you have questions about the logic of it, feel free to ask and I can help with a detail.

FIRST TIME USE
Like any Excel project using Macros, on opening the file, the first thing that you need to do is to enable Macros. Simply click on the warning at the top of the spreadsheet and click enable content. You can ignore the warning about the project being malicious, this is just there to let you know that code not written by Microsoft will be run (which can be harmful) but this is a very open project where the code is clearly visible so you do not need to worry (although, feel free to look through the code by pressing F11 if you want to see for yourself or clicking on view macros).

The next step is to choose the grading system and grade boundaries that you want to use. On the settings tab, click on the button labelled 'Set Grades' you will have the option of choosing one of two presets (GCSE or KS3 where you choose the highest NC level) for grades or can enter your own grading system but you MUST use 9 grades (although there is a way to get around this). You then have to enter the grade boundaries for each grade or level. You should enter the lowest percentage that achieves that grade (e.g. if A is 80% - 90%, enter 80 for A). If you want to skip some of the grades as you do not wish to use them, I suggest entering the same boundary multiple times when you reach the lower end (e.g. if you only want A* - D, use A* - 90, A - 80, B - 70, C - 60, D - 50, E - 50, F - 50, G - 50, U - 0 - this will avoid unwanted highlighting which results from setting them to 0). The lowest grade will always be set to 0 automatically.
If you later want to change the grade boundaries, simply edit the boundary or grade manually (but make sure that it follows logic, e.g. if A* is 90, A cannot be 95).

ADDING STUDENTS
Once you have set the grade boundaries, you can now enter the students. There are two ways of doing this. Either way, first choose the appropriate year tab at the bottom. These can be renamed, re-coloured and copied if desired. Note that every year will use the same grades and boundaries.
The first method is to add the students one at a time by using the 'Add Student' button. You will then be prompted to add the student's name, surname, teaching or tutor group, any AEN and their target grade. Formulae will also be added automatically to calculate their average mark, ranking, etc.. This method is best if you are adding the students from scratch.
The other method is designed for adding students in large batches. Click the 'Add Student Spaces' button and enter the number of students that you want to add. This will add the formulae for these students but leave fields such as name, group, etc. blank. You can then copy and paste this information over from another spreadsheet or enter it manually. This method is best if a lot of the data is similar and can just be copied or if you already have a .xlsm or .csv file containing the information (e.g. from a register or other mark book).

UTILISING THE PROGRAM
The program has a large number of features which can be used or ignored as needed.

In order to use any feature, you first need to enter some scores. Scores are always a percentage and should be entered in columns N to ZZ, giving you space for up to 689 tests (a limit which will never be reached).

The simplest feature is averages, this averages out any numerical value entered past column M. Feel free to label the headers of these columns, this was only not done automatically due to the high variance in the number needed and style desired. The average is also placed next to the students' information to make it clear when quickly scrolling.

The  highest and lowest scores that a student has received are also displayed to the left of the average on the right.

Another feature is a ranking system that ranks the students on each sheet based on their average score with a rank of 1 being the best student.

The most obvious feature is perhaps the cell highlighting. There are four different types used on each sheet. The first is raw percentages. This compares the data in the cell against the grade boundaries entered on the settings tab, it then highlights the cell in the correct colour (this is slightly oversimplified but this is how it appears front-end). The order of colours from high to low is gold, silver, bronze, dark green, light green, yellow, orange, red and grey. The second type uses the same colour system but compares a grade entered against the grades on the settings tab and highlights. The third type is for the ranking, this highlights the top 10% of students dark blue, the top 25% light blue, the bottom 25% light purple and the bottom 10% dark blue allowing the highest and lowest performing students and making them stand out. The forth type is to do with achievement and will be explained below.

Whilst these features occur automatically, the remaining features require pushing the 'Calculate Grades' button to update. This will calculate and display the grade of each student based on their average score in the 'Grade' column. This grade is then compared against their target grade and a score entered in the 'Achievement' column. This represents how many grades above or below their target they are. For example, a student with a target grade of B that is achieving an A will get a score of 1. If they are achieving a B, they get a score of 0 and if they are getting a D, they get -2. These scores are then highlighted using a shortened version of the standard colouring (3+ - gold, 2 - silver, 1 - bronze, -1 - yellow, -2 - orange, -3 - red). 0 is not highlighted in any cell but here it will likely be most noticeable due to blank cells having a value of 0 so every blank cell would also be highlighted. The decision was therefore made to leave all cells of value 0 un-highlighted.

KNOWN ISSUES
[none currently]
