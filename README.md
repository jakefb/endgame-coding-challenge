
# EndGame coding challenge

Mini spreadsheet build with JavaScript, HTML and CSS that looks and works like Excel.

You can view the live application here https://jakefb-mini-spreadsheet.netlify.app

Task steps:

1. Create index.html in your favourite text editor. Use pure JavaScript in your code – no JavaScript libraries or frameworks. Your JavaScript can either be in a separate .js file, or it can be contained in the index.html page.
2. When loading index.html into Chrome or Firefox, it should draw a 100x100 grid of cells, with columns labelled A-Z, AA, AB, AC, etc. and rows numbered 1 to 100.
3. When you click in a cell and enter a number, it should store the number in a JavaScript object (note: this would be lost when you refresh the page).
4. Have a refresh button that redraws the grid (without refreshing the page) and inserts data into cells where you've saved it.
5. Add support for some basic formulas. For example if you enter "=A1+A2" into A3 it should calculate the sum of these two cells and display the result in A3. Updating A1 would update A3.
6. Add support for some basic functions. For example if you enter "=sum(A1:A10)" into A11, then it should calculate the sum of all cells in the range and display the result in A11. Updating any value in the range would recalculate A11.
7. Add support for formatting, for example bold, italics and underline
