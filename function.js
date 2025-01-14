const functions = document.getElementsByClassName("function-dropdown")[0];
console.log(functions);

function onClickFunctions(e) {
  const functionClass = document.getElementsByClassName("function-dropdown")[0];
  functionClass.classList.toggle("show");
}

function functionSum(e) {
  activeCell.innerText = "=SUM()";
  activeCell.addEventListener("keydown", calculateSum);
}

function calculateSum(e) {
  if (e.keyCode === 13) { // Check if Enter key is pressed
    let sum = 0;
    const formula = e.target.innerText.trim(); // Get the formula
    const rangeMatch = formula.match(/=SUM\(([a-zA-Z]+\d+):([a-zA-Z]+\d+)\)/i); // Match the range

    if (rangeMatch) {
      const startCell = rangeMatch[1].toUpperCase();
      const endCell = rangeMatch[2].toUpperCase();

      // Convert cell references to row/column indices
      const [startCol, startRow] = parseCellReference(startCell);
      const [endCol, endRow] = parseCellReference(endCell);

      // Loop through the range and sum values
      for (let i = startRow; i <= endRow; i++) {
        for (let j = startCol; j <= endCol; j++) {
          const cellData = data[currentSheetIndex - 1][i][j];
          const cellValue = parseInt(cellData.innerText || "0", 10);
          sum += isNaN(cellValue) ? 0 : cellValue;
        }
      }

      activeCell.innerText = sum; // Update the cell with the sum
    } else {
      activeCell.innerText = "Invalid Formula"; // Handle invalid formulas
    }

    activeCell.removeEventListener("keydown", calculateSum);
  }
}

// Helper function to convert cell reference (e.g., A1) to row/column indices
function parseCellReference(cell) {
  const match = cell.match(/([a-zA-Z]+)(\d+)/);
  const col = match[1].toUpperCase().charCodeAt(0) - 65; // Convert column letter to index
  const row = parseInt(match[2], 10) - 1; // Convert row number to 0-based index
  return [col, row];
}


function functionAverage(e) {
  activeCell.innerText = "=AVERAGE()";
  activeCell.addEventListener("keydown", calculateAverage);
}

function calculateAverage(e) {
  if (e.keyCode === 13) { // Trigger on Enter key
    const formula = e.target.innerText.trim(); // Get the formula
    const cellText = formula.match(/[a-zA-Z]+\d+/g); // Match all cell references

    if (!cellText || cellText.length === 0) {
      activeCell.textContent = "Invalid Formula"; // Handle invalid formula
      return;
    }

    let sum = 0;
    let validCellCount = 0; // Count valid cells with numerical data

    for (let i = 0; i < cellText.length; i++) {
      let found = false; // Track if a cell was matched
      const targetCellId = cellText[i].toUpperCase(); // Normalize cell ID

      // Search for the target cell in the grid
      for (let row of data[currentSheetIndex - 1]) {
        for (let cell of row) {
          if (cell.id === targetCellId) {
            const cellValue = parseInt(cell.innerText || "0", 10); // Parse cell value
            if (!isNaN(cellValue)) {
              sum += cellValue; // Add to sum if valid number
              validCellCount++;
            }
            found = true; // Cell matched
            break; // Exit inner loop
          }
        }
        if (found) break; // Exit outer loop if cell found
      }
    }

    if (validCellCount === 0) {
      activeCell.textContent = "No Valid Cells"; // Handle case with no valid cells
    } else {
      activeCell.textContent = (sum / validCellCount).toFixed(2); // Calculate average and display it
    }

    activeCell.removeEventListener("keydown", calculateAverage);
  }
}

function functionCount(e) {
  activeCell.innerText = "=COUNT()";
  activeCell.addEventListener("keydown", calculateCount);
}

function calculateCount(e) {
  if (e.keyCode === 13) { // Trigger on Enter key
    const formula = e.target.innerText.trim(); // Get the formula text
    const rangeMatch = formula.match(/=COUNT\(([a-zA-Z]+\d+):([a-zA-Z]+\d+)\)/i); // Match the range

    if (rangeMatch) {
      const startCell = rangeMatch[1]; // Extract the starting cell (e.g., A10)
      const endCell = rangeMatch[2];   // Extract the ending cell (e.g., A55)

      // Extract row numbers from the start and end cells
      const startRow = parseInt(startCell.match(/\d+/)[0], 10); // Extract 10 from A10
      const endRow = parseInt(endCell.match(/\d+/)[0], 10);     // Extract 55 from A55

      // Calculate the count based on the range
      const count = Math.abs(endRow - startRow) + 1; // Add 1 to include both start and end rows
      activeCell.textContent = count; // Display the result in the cell
    } else {
      activeCell.textContent = "Invalid Formula"; // Handle invalid input
    }

    activeCell.removeEventListener("keydown", calculateCount);
  }
}

function functionMax(e) {
  activeCell.innerText = "=MAX()";
  activeCell.addEventListener("keydown", calculateMax);
}

function calculateMax(e) {
  if (e.keyCode === 13) {
    const formula = e.target.innerText;
    const cellText = formula.match(/[a-zA-Z]+\d+/g);

    let sum = -Infinity;

    for (let i = 0; i < cellText.length; i++) {
      here: for (let j = 0; j < data[currentSheetIndex - 1].length; j++) {
        for (let k = 0; k < data[currentSheetIndex - 1][j].length; k++) {
          if (
            cellText[i].toUpperCase() === data[currentSheetIndex - 1][j][k].id
          ) {
            sum = Math.max(
              parseInt(
                data[currentSheetIndex - 1][j][k].innerText === ""
                  ? "0"
                  : data[currentSheetIndex - 1][j][k].innerText
              ),
              sum
            );
            break here;
          }
        }
      }
    }

    activeCell.innerHTML = sum === -Infinity ? 0 : sum;
    activeCell.removeEventListener("keydown", calculateMax);
  }
}

function functionMin(e) {
  activeCell.innerText = "=MIN()";
  activeCell.addEventListener("keydown", calculateMin);
}

function calculateMin(e) {
  if (e.keyCode === 13) {
    const formula = e.target.innerText;
    const cellText = formula.match(/[a-zA-Z]+\d+/g);

    let sum = Infinity;

    for (let i = 0; i < cellText.length; i++) {
      here: for (let j = 0; j < data[currentSheetIndex - 1].length; j++) {
        for (let k = 0; k < data[currentSheetIndex - 1][j].length; k++) {
          if (
            cellText[i].toUpperCase() === data[currentSheetIndex - 1][j][k].id
          ) {
            sum = Math.min(
              parseInt(
                data[currentSheetIndex - 1][j][k].innerText === ""
                  ? "0"
                  : data[currentSheetIndex - 1][j][k].innerText
              ),
              sum
            );
            break here;
          }
        }
      }
    }

    activeCell.textContent = sum === Infinity ? 0 : sum;
    activeCell.removeEventListener("keydown", calculateMin);
  }
}