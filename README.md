# CCT_Member_Sort
This is a small Java tool I made to sort through the member manifest of my club, the Cornell Club of Taiwan. Essentially, it goes through each name
on the list, goes online and searches the person's ID on the Cornell People database. If the person's ID can't be found, the program will then search 
the database using the person's name. Looking through the html code of the page pulled up from the database, the program can than tell if the person
is a student or an alumni and sort the person into different excel spreadsheets accordingly. In the end, names that could not be sorted is printed to
a file for the user to look through.