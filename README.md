# VBA-challenge

Purpose of this project was to create a VBA script that looped through all stocks and printed the following information:
  - Ticker Symbol
  - Yearly Change over the year
  - Percent Change over the year
  - Total Stock Volume
  
Additionally, a summary table was created that contained the following information:
  - Greatest % increase per year
  - Greatest % decrease per year
  - Greatest total volume
  
  
This program loops through each excel worksheet and prints necessary stock information in tables on the right.


Assistance was gathered from AskBCS to help with formatting numbers to percentages.
  ws.Cells(i, 11).NumberFormat = "0.00%"
  
Assistance was gathered from Microsoft to help autofit columns.
  ws.Columns("A:Q").AutoFit

Assistance from AskBCS to help save the file as a .vbs file through VSCode.
