# Calculating Total Grades From Raw Data

A spreadsheet with raw grading data was downloaded that contained multiple assignment types, and the target assignment types ("Quiz", "Exam", and "Activity") needed to be extracted from the strings within each row of that column. Extraneous columns were deleted, and the VBA [Like operator](https://learn.microsoft.com/en-us/dotnet/visual-basic/language-reference/operators/like-operator) was used in conjunction with the [* wildcard](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/wildcard-characters-used-in-string-comparisons) to find and label data in a new column. Knowing that all target assignments were scored out of 100 points, an additional column was created to list the Total Points Possible only in cells that were not blank using conditional statments and the [ISBLANK function](https://support.microsoft.com/en-us/office/is-functions-0f2d7971-6019-40a0-a171-f2d869135665). Rows with other assignment types were left blank. 

<img width="344" alt="spreadsheet_screenshot" src="https://user-images.githubusercontent.com/111674383/214432247-14f6c1c3-64ce-4925-bb5a-532e260d6925.png">

A pivot table was then used to break down the total number of points earned and possible for each activity type. Filters were used to display only the target assignment types and none of the blank rows.

<img width="238" alt="pivot_table_screenshot" src="https://user-images.githubusercontent.com/111674383/214434153-a664f3b6-3a51-40c3-b716-a51db7ef3d8d.png">

The raw points earned and possible were then ready to be inputted into school records.
