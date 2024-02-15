# second - Stock Analysis
  * Created a script that loops through all the stocks for one year for all 3 worksheets and outputs the following information:
  * The ticker symbol
  * Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
  * The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
  * The total stock volume of the stock.
  * Adedd functionality to script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". 
  1. Use the Dim statement to declare the data type
  2. To loop worksheets use code provided by Instructor"For Each WS In ThisWorkbook.Worksheets" and insert WS before cell value to loop through all spreadsheets
  3. Defined Last row using code on Wallstreetmojo.com/vba-last-row/ as lastrow = WS.UsedRange.Rows.Count
  4. Established For Next loop to help repeat actions from row 2 through last row
  5. Set initial values for output row as 2 and increment output row by 1
  6. Set initial values for Total volume as 0
  7. Added summary table headers for each spreadsheet
  8. Write starting volume formula so it doesnt aggregate as zero
  9. Write If then statement year starting opening value. Format the value
  10. Write If then statement year closing value. Format the value
  11. Write ticker value of cells
  12. Write yearly change = Close Value - Opem Value
  13. Format the interior of the changes. Color key was provided by Instructor
  14. Write yearly percent change. The number format was provided over Slack
  15. Write Total Volume. Set TotalVolume back to 0 so that we can get the next ticker's volume.
  16. Increment the output row by 1
  17. Assign cell holding the value from range in Column K with worksheet function of Max or Min, format as decimal
  18. Maximum value funtion was available on stackoverflow.com/questions/42633273/finding-max-of-a-column-in-vba/42633375#42633375
  19. Assign cell holding the value from range in Column L with worksheet function of Max
  20. Format Max_Increase as percent
  21. Find and Activate the value of Max_Increase from range in Column K.  Lookup learn.microsoft.com/en-us/office/vba/api/excel/excel.range.activate
  22. Offset the activecell to get the corresponding ticker
  23. Format Min_Increase as percent
  24. Find and Activate the value of Min_Increase from range in Column K
  25. Offset the activecell to get the corresponding ticker
  26. Find and Activate the value of Max_Volume from range in Column L
