const COLUMN_LENGTH = 100
const ROW_LENGTH = 100

const CELL_WIDTH = '6rem'

const INITIAL_STATE = {
  A1: '1',
  B1: '2',
  C1: '3',
  A2: '4',
  B2: '5',
  C2: '6',
  A3: '7',
  B3: '8',
  C3: '9',
  A4: '=sum(A1:C3)',
  B4: '=average(A1:C3)',
  D1: '10',
  D2: '20',
  D3: '30',
  D4: '=sum(D1:D3)',
  C4: '=A1+C3/C1+B3-B1*C1',
  A5: '=sum(A4:C4)',
  B5: '=A4/B1',
  C5: '=7^2'
}

const excelColumnToNumber = column => {
  const base = 'A'.charCodeAt(0) - 1

  return column
    .split('')
    .reduce(
      (acc, character) => (acc * 'Z'.charCodeAt(0) + 1) + character.charCodeAt(0) - base, 
      0
    )
}

const indexToColumnName = ({ index, prefix = '' }) => 'A'.charCodeAt(0) + index <= 'Z'.charCodeAt(0)
  ? prefix + String.fromCharCode('A'.charCodeAt(0) + index)
  : indexToColumnName({
      index: index - ('Z'.charCodeAt(0) - 'A'.charCodeAt(0) + 1),
      prefix: prefix === ''
        ? 'A'
        : String.fromCharCode(prefix.charCodeAt(0) + 1)
    })

const initializeSheetState = () => {
  const state = Object.fromEntries(
    Array.from(
      { length: COLUMN_LENGTH },
      (_, rowIndex) => Array.from(
        { length: ROW_LENGTH },
        // Initialize key value pairs
        (_, columnIndex) => [indexToColumnName({ index: columnIndex }) + (rowIndex + 1), '']
      )
    )
      .flat()
  )

  return state
}

const header = Array.from(
  { length: COLUMN_LENGTH },
  (_, columnIndex) => indexToColumnName({ index: columnIndex })
)

const state = {
  ...initializeSheetState(),
  ...INITIAL_STATE
}

const convertCellToNumber = value => value.substring(0, 1) === '='
  ? calculateCell(value)
  : Number(value)

const clearHighlightedCells = () => {
  document.querySelectorAll('input.cell')
    .forEach(
      element => {
        element.style.backgroundColor = '#fff'
      }
    )
}

const filterCellsFromRange = ({ startRange, endRange }) => {
  const startRangeColumn = startRange.match(/^[A-Z]+/i)[0].toUpperCase() 
  const endRangeColumn = endRange.match(/^[A-Z]+/i)[0].toUpperCase()
  const startRangeRow = Number(startRange.match(/[0-9]+$/)[0])
  const endRangeRow = Number(endRange.match(/[0-9]+$/)[0])

  return Object.entries(state)
    .filter(
      ([key]) => excelColumnToNumber(key.match(/^[A-Z]+/)[0]) >= excelColumnToNumber(startRangeColumn) && excelColumnToNumber(key.match(/^[A-Z]+/)[0]) <= excelColumnToNumber(endRangeColumn)
    )
    .filter(
      ([key]) => key.match(/[0-9]+$/)[0] >= startRangeRow && key.match(/[0-9]+$/)[0] <= endRangeRow
    )
}

const findMatchesForCellFunction = (cellValue) => cellValue.replace(/\s/g, '').match(/^\=(sum|average)\(([A-Z]+[0-9]+):([A-Z]+[0-9]+)\)$/i)

const CELL_FORMULA_REGEX = /^\=[\^A-Z0-9*/+-]+$/i
const CELL_FORMULA_OPERATORS_ARRAY = ['^', '*', '/', '+', '-', ]
const CELL_FORMULA_OPERATORS_REGEX = /([\^*/+-])/

const testCellFormula = (cellValue) => cellValue.replace(/\s/g, '').substring(0, 1) === '=' && CELL_FORMULA_REGEX.test(cellValue.replace(/\s/g, ''))

const highlightCells = (cellValue) => {
  clearHighlightedCells()

  const cellFunctionMatchArray = findMatchesForCellFunction(cellValue)

  if (cellFunctionMatchArray !== null) {
    const [_wholeMatch, _cellFunctionName, startRange, endRange] = cellFunctionMatchArray

    filterCellsFromRange({ startRange, endRange })
      .forEach(([key]) => {
        document.getElementById(key).style.backgroundColor = '#fff2c8'
      })
  } else if (testCellFormula(cellValue)) {
    ;cellValue
      .replace(/\s/g, '')
      .slice(1) // Remove the leading =
      .split(CELL_FORMULA_OPERATORS_REGEX)
      .filter(
        current => CELL_FORMULA_OPERATORS_ARRAY.includes(current)
          ? false
          : /^[A-Z]+[0-9]+$/i.test(current)
            ? true
            : false
      )
      .forEach(
        key => {
          document.getElementById(key.toUpperCase()).style.backgroundColor = '#fff2c8'
        }
      )
  }
}

// Apply exponent calculations
const applyExponentiation = (mathArray) => mathArray
  .slice(1)
  .reduce(
    (acc, current, index, array) => index % 2 === 0
      ? ['^'].includes(current)
        ? [
            ...acc.slice(0, -1),
            acc[acc.length - 1] ** array[index + 1]
          ]
        : [ ...acc, current ]
      : ['^'].includes(array[index - 1])
        ? acc
        : [ ...acc, current ],
    [ mathArray[0] ]
  )


// Apply multiplication and division calculations
const applyMultiplicationAndDivision = (mathArray) => mathArray
  .slice(1)
  .reduce(
    (acc, current, index, array) => index % 2 === 0
      ? ['*', '/'].includes(current)
        ? current === '*'
          ? [
              ...acc.slice(0, -1),
              acc[acc.length - 1] * array[index + 1]
            ]
          : [
              ...acc.slice(0, -1),
              acc[acc.length - 1] / array[index + 1]
            ]
        : [ ...acc, current ]
      : ['*', '/'].includes(array[index - 1])
        ? acc
        : [ ...acc, current ],
    [ mathArray[0] ]
  )

// Apply addition and subtraction calculations
const applyAdditionAndSubtraction = (mathArray) => mathArray
  .slice(1)
  .reduce(
    (acc, current, index, array) => index % 2 === 0
      ? ['+', '-'].includes(current)
        ? current === '+'
          ? [
              ...acc.slice(0, -1),
              acc[acc.length - 1] + array[index + 1]
            ]
          : [
              ...acc.slice(0, -1),
              acc[acc.length - 1] - array[index + 1]
            ]
        : [ ...acc, current ]
      : ['+', '-'].includes(array[index - 1])
        ? acc
        : [ ...acc, current ],
    [ mathArray[0] ]
  )

const calculateCell = (cellValue) => {
  const cellFunctionMatchArray = findMatchesForCellFunction(cellValue)

  if (cellFunctionMatchArray !== null) {
    const [_wholeMatch, cellFunctionName, startRange, endRange] = cellFunctionMatchArray

    const isAverage = cellFunctionName === 'average'

    return filterCellsFromRange({ startRange, endRange })
      .map(
        ([_key, value]) => convertCellToNumber(value)
      )
      .reduce(
        (acc, current, index, array) => (isAverage && index === array.length - 1)
          ? (acc + current) / array.length
          : acc + current,
        0
      )
  } else if (testCellFormula(cellValue)) {
    const mathArray = cellValue
      .replace(/\s/g, '')
      .slice(1) // Remove the leading =
      .split(CELL_FORMULA_OPERATORS_REGEX)
      .map(
        current => CELL_FORMULA_OPERATORS_ARRAY.includes(current)
          ? current
          : /^[A-Z]+[0-9]+$/i.test(current)
            ? convertCellToNumber(state[current.toUpperCase()])
            : convertCellToNumber(current)
      )

    const mathArraySolved = applyAdditionAndSubtraction(
      applyMultiplicationAndDivision(
        applyExponentiation(mathArray)
      )
    )

    if (mathArraySolved.length !== 1) {
      console.error(`Could not parse math formula: ${cellValue}`)
      return NaN
    }

    return mathArraySolved[0]
  } else {
    return cellValue
  }
}

const renderSheet = ({ appRoot }) => {
  const grid = document.createElement('div')
  grid.className = 'grid'
  grid.style.gridTemplateColumns = `repeat(calc(${COLUMN_LENGTH} + 1), ${CELL_WIDTH})`

  const cell = document.createElement('div')
  cell.innerText = ''
  cell.className = 'header'

  grid.appendChild(cell)

  // Catch keydown events that bubble up from cell input fields
  grid.addEventListener(
    'keydown', // Use keydown and preventDefault() to prevent browser default behavior 
    (event) => {
      const element = event.target

      if ((event.ctrlKey || event.metaKey)) {
        switch (event.key) {
          case 'b':
            element.style.fontWeight = element.style.fontWeight === '700' ? '400' : '700'
            event.preventDefault()
            break
          case 'i':
            element.style.fontStyle = element.style.fontStyle === 'italic' ? 'normal' : 'italic'
            event.preventDefault()
            break
          case 'u':
            element.style.textDecoration = element.style.textDecoration === 'underline' ? 'none' : 'underline'
            event.preventDefault()
            break
        }
      }
    }
  )

  // Catch keyup events that bubble up from cell input fields
  grid.addEventListener(
    'keyup', // Use keyup instead of input to prevent loop triggered by programmatically changing the element value
    (event) => {
      const element = event.target

      if (element.value.substring(0, 1) === '=' && element.value.toUpperCase().includes(element.id)) {
        state[element.id] = '💣'
        return
      }

      state[element.id] = element.value

      Object.entries(state).filter(([_key, value]) => value.substring(0, 1) === '=')
        .forEach(
          ([key, _value]) => {
            const el = document.getElementById(key)
            if (el !== document.activeElement) {
              el.value = calculateCell(state[key])
            }
          }
        )

      highlightCells(element.value)
    }
  )

  // Catch focusout events that bubble up from cell input fields
  grid.addEventListener(
    'focusout',
    (event) => {
      const element = event.target

      element.value = calculateCell(state[element.id])

      clearHighlightedCells()
    }
  )

  // Catch focusin events that bubble up from cell input fields
  grid.addEventListener(
    'focusin',
    (event) => {
      const element = event.target

      element.value = state[element.id]

      highlightCells(element.value)
    }
  )

  header.forEach(
    current => {
      const cell = document.createElement('div')
      cell.innerText = current
      cell.className = 'header'

      grid.appendChild(cell)
    }
  )

  Object.entries(state)
    .forEach(
      ([currentKey, currentValue], index) => {
        if ((index + 1) % ROW_LENGTH === 1) {
          const cell = document.createElement('div')
          cell.innerText = (index / ROW_LENGTH) + 1
          cell.className = 'header'
    
          grid.appendChild(cell)
        }

        const cell = document.createElement('input')
        cell.value = currentValue.substring(0, 1) === '='
          ? calculateCell(state[currentKey])
          : currentValue
        cell.id = currentKey
        cell.className = 'cell'

        grid.appendChild(cell)
      }
    )

  appRoot.appendChild(grid)
}

const renderFooter = ({ appRoot }) => {
  const footer = document.createElement('div')
  footer.className = 'footer'

  const notes = document.createElement('p')
  notes.innerHTML = `
    Cell formatting: bold (⌘-B), italic (⌘-I), underline (⌘-U).<br>
    Cell formulas: =A1+A2 (supported operators: ^, *, /, +, -)<br>
    Cell functions: =sum(A1:A10), =average(A1:A10)
  `

  const button = document.createElement('button')
  button.innerText = 'Refresh'
  button.className = 'refresh-button'
  button.addEventListener(
    'click',
    () => {
      renderApp({ appRoot })
    }
  )

  appRoot.appendChild(footer)
  footer.appendChild(notes)
  footer.appendChild(button)
}

const renderApp = ({ appRoot }) => {
  appRoot.innerHTML = ''
  renderSheet({ appRoot })
  renderFooter({ appRoot })
}

// Check that the DOM is loaded
document.addEventListener('DOMContentLoaded', function () {
  // Render grid
  renderApp({ appRoot: document.getElementById('app') })
})
