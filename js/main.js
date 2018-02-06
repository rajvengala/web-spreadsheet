class Worksheet {

    constructor(rows, columns) {
        this.maxRows = rows;
        this.maxColumns = columns;
        this.worksheetData = Array(this.maxRows).fill().map(() => Array(this.maxColumns).fill());
        this.selectedRow = 0; // used to track the currently selected row for add/delete operations
        this.rowNum = 0;
    }

    createSheet() {
        // register global event handler
        document.addEventListener('keydown', event => {
            if (event.code === 'Escape') { // Escape key
                let rowContextMenu = document.getElementById('row-context-menu');
                if (rowContextMenu) {
                    rowContextMenu.style.visibility = 'hidden';
                }
            }
        });

        // create header row
        let sheetHeader = document.getElementById('sheet-header');
        for (let col = 1; col <= this.maxColumns; col++) {
            // create header cell
            let headerCell = document.createElement('input');
            let headerCellId = 'header-cell-' + col;
            headerCell.setAttribute('id', headerCellId);
            headerCell.setAttribute('readonly', 'readonly');
            headerCell.setAttribute('class', 'cell');
            if (col > 1) {
                headerCell.setAttribute('value', Utils.getColumnNameInWorksheet(col - 1));
            }

            // add header row to header container
            sheetHeader.appendChild(headerCell);
        }

        // create rest of the rows
        let sheetBody = document.getElementById('sheet-body');
        for (let row = 1; row <= this.maxRows; row++) {
            this.createBodyRow(row, sheetBody, true); // true - append row;
        }

        // make first body cell editable
        document.getElementById('body-cell-row-1-col-2').removeAttribute('readonly');
        document.getElementById('body-cell-row-1-col-2').focus();

    }

    bodyCellEventHandler(event) {
        const currentCell = event.target.id;
        let bodyCellIdTokens = currentCell.split('-');
        let rowNum = parseInt(bodyCellIdTokens[3]);
        let colNum = parseInt(bodyCellIdTokens[5]);
        let currentRowNum = rowNum;
        let currentColNum = colNum;
        this.selectedRow = currentRowNum;

        // Mouse event
        if (event.button !== undefined) {
            switch (event.button) {
                case 0: // left mouse click
                    // make cell editable only if it is not in first column
                    if (currentCell.indexOf('col-1') === -1) {
                        document.getElementById(document.activeElement.id).setAttribute('readonly', 'readonly');
                        document.getElementById(event.target.id).removeAttribute('readonly');
                        document.getElementById(event.target.id).focus();
                    } else {
                        // close any row context menu, if one is visible
                        if (document.getElementById('row-context-menu')) {
                            const escapeEvent = new KeyboardEvent('keydown', {
                                'key': 'Escape',
                                'code': 'Escape'
                            });
                            document.dispatchEvent(escapeEvent);
                        }
                    }
                    break;
                case 2: // right mouse click
                    // if col is 1, show row context menu
                    if (currentColNum === 1) {
                        // unselected previously highlighted row
                        let highlightedRow = document.getElementById('selected-row');
                        if (document.getElementById('selected-row')) {
                            highlightedRow.removeAttribute('id');
                        }

                        // highlight selected row
                        let rowSeqNum = this.findSelectedRowSeqNum(this.selectedRow);
                        let selectedRow = document.getElementById('sheet-body').children[rowSeqNum - 1];
                        selectedRow.setAttribute('id', 'selected-row');

                        if (!document.getElementById('row-context-menu')) {
                            let contextMenuEle = document.createElement('div');
                            contextMenuEle.setAttribute('id', 'row-context-menu');
                            let listEle = document.createElement('ul');
                            let addRowEle = document.createElement('li');
                            addRowEle.setAttribute('id', 'add-row');
                            addRowEle.innerText = 'Add Row';
                            let delRowEle = document.createElement('li');
                            delRowEle.setAttribute('id', 'del-row');
                            delRowEle.innerText = 'Delete Row';
                            listEle.appendChild(addRowEle);
                            listEle.appendChild(delRowEle);
                            contextMenuEle.appendChild(listEle);

                            addRowEle.addEventListener('click', (event) => {
                                const escapeEvent = new KeyboardEvent('keydown', {
                                    'key': 'Escape',
                                    'code': 'Escape'
                                });
                                document.dispatchEvent(escapeEvent);
                                let rowSeqNum = this.findSelectedRowSeqNum(this.selectedRow);
                                this.insertRow(event, rowSeqNum);
                            });

                            delRowEle.addEventListener('click', (event) => {
                                const escapeEvent = new KeyboardEvent('keydown', {
                                    'key': 'Escape',
                                    'code': 'Escape'
                                });
                                document.dispatchEvent(escapeEvent);
                                let rowSeqNum = this.findSelectedRowSeqNum(this.selectedRow);
                                this.deleteRow(event, rowSeqNum);
                            });

                            let body = document.getElementsByTagName('body')[0];
                            body.appendChild(contextMenuEle);

                        }
                        let rowContextMenu = document.getElementById('row-context-menu');
                        rowContextMenu.style.top = (event.clientY) + 'px';
                        rowContextMenu.style.left = (event.clientX + 5) + 'px';
                        rowContextMenu.style.visibility = 'visible';
                    }
            }
        } else {
            const keyName = event.key;
            switch (keyName) {
                case 'ArrowUp':
                    // decrease row num
                    if (rowNum !== 1)
                        rowNum -= 1;
                    break;
                case 'ArrowRight':
                    // increase col num
                    if (colNum !== this.maxColumns)
                        colNum += 1;
                    break;
                case 'ArrowDown':
                    // decrease row num
                    if (rowNum !== this.maxRows)
                        rowNum += 1;
                    break;
                case 'ArrowLeft':
                    // decrease col num
                    if (colNum > 2)
                        colNum -= 1;
                    break;
            }

            // Stop default behavior for arrow key press actions
            if (keyName && keyName.indexOf('Arrow') !== -1) {
                event.preventDefault();
                event.stopPropagation();

                let newBodyCellIdToFocus = 'body-cell-row-' + rowNum + '-col-' + colNum;
                // update the model with the cell value
                let cellVal = document.getElementById(currentCell).value;

                this.worksheetData[currentRowNum - 1][currentColNum - 2] = cellVal;
                document.getElementById(currentCell).setAttribute('readonly', 'readonly');
                document.getElementById(newBodyCellIdToFocus).removeAttribute('readonly');
                document.getElementById(newBodyCellIdToFocus).focus();
            }
        }
    }

    // find the row sequence number from the id attribute of the selected row
    findSelectedRowSeqNum(rowId) {
        let sheetBodyRow = document.getElementsByClassName('sheet-body-row');
        let indexOfRowId = Array.from(sheetBodyRow).findIndex((ele) => {
            let cellId = ele.getElementsByTagName('input')[0].id;
            return (cellId.indexOf('row-' + rowId) !== -1);
        });
        return indexOfRowId + 1;
    }

    createBodyRow(row, sheetBody, toAppend) {
        let referenceRow;
        if (!toAppend) {
            referenceRow = sheetBody.children[row];
            row = this.maxRows;
        }

        // create container for row
        let sheetBodyRow = document.createElement('div');
        sheetBodyRow.setAttribute('class', 'sheet-body-row');

        for (let col = 1; col <= this.maxColumns; col++) {
            // create body cell
            let bodyCell = document.createElement('input');
            let bodyCellId = 'body-cell-row-' + row + '-col-' + col;
            bodyCell.setAttribute('id', bodyCellId);
            bodyCell.setAttribute('class', 'cell');
            if (col === 1) {
                bodyCell.setAttribute('class', 'cell body-cell-col-1');
                bodyCell.setAttribute('value', row.toString());
                bodyCell.addEventListener('rowsChanged', (event) => {
                    this.rowNum += 1;
                    bodyCell.setAttribute('value', this.rowNum.toString());
                });
            }
            bodyCell.setAttribute('readonly', 'readonly');
            bodyCell.addEventListener('keydown', (event) => this.bodyCellEventHandler(event));
            bodyCell.addEventListener('contextmenu', (event) => event.preventDefault());
            bodyCell.addEventListener('mousedown', (event) => {
                this.bodyCellEventHandler(event)
            });

            // add cells to row
            sheetBodyRow.appendChild(bodyCell);
        }
        // add row to sheet
        if (toAppend) {
            sheetBody.appendChild(sheetBodyRow);
        } else {
            sheetBody.insertBefore(sheetBodyRow, referenceRow);
        }
    }

    deleteRow(event, rowNum) {
        this.maxRows -= 1;

        let sheetBody = document.getElementById('sheet-body');
        document.getElementById('sheet-body').children[rowNum - 1].remove();
        this.worksheetData.splice(rowNum - 1, 1);

        let rowNumCols = document.getElementsByClassName('cell body-cell-col-1');
        this.rowNum = 0; // reset rowNum
        Array.from(rowNumCols).map((ele) => {
            ele.dispatchEvent(new CustomEvent('rowsChanged'));
        });
    }

    insertRow(event, rowNum) {
        this.maxRows += 1;
        let sheetBody = document.getElementById('sheet-body');
        this.createBodyRow(rowNum, sheetBody, false); // false - insert row at position rowNum
        this.worksheetData.push(Array(this.maxColumns).fill());

        let rowNumCols = document.getElementsByClassName('cell body-cell-col-1');
        this.rowNum = 0; // reset rowNum
        Array.from(rowNumCols).map((ele) => {
            ele.dispatchEvent(new CustomEvent('rowsChanged'));
        });
    }
}

class Utils {
    // Generate column names for header cells
    static getColumnNameInWorksheet(columnNum) {
        let colName = '';

        let colNameCharCount = Math.floor(columnNum / 26); // 26 - Number of english alphabets
        let colNameChar = columnNum % 26;

        if (colNameCharCount > 0 && colNameChar !== 0) {
            colName += String.fromCharCode(colNameCharCount + 64);
        }

        if (colNameChar === 0) colNameChar = 26;
        return colName + String.fromCharCode(colNameChar + 64);
    }
}

window.onload = () => {
    const rows = 50;
    const columns = 50;
    let worksheet = new Worksheet(rows, columns);
    worksheet.createSheet();
};
