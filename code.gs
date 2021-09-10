function onOpen() {

    var ui = SpreadsheetApp.getUi()
    ui.createMenu("Antro's Actions")
      .addItem("Format Array Data Region", "formatRegion")
      .addItem("Format Array Data Region & Log", "formatRegionLog")
      .addToUi();     
  }
  function formatRegionLog(){
    formatRegion( true )
  }
  function formatRegion( doLog = false ){
  
    // get the current sheet name and cell address and formula
    const ss = SpreadsheetApp.getActive()
    const shData = ss.getActiveSheet()
    const shName = shData.getSheetName()
    const activeCell = ss.getActiveCell()
    const cCell = activeCell.getA1Notation()
    const cRow = activeCell.getRow()
    const cCol = activeCell.getColumn()
    let form = activeCell.getFormula()
    if ( form === "" ){
      showPrompt("The current cell (" + cCell + ") doesn't contain a formula.")
      return
    }
    let shConfig
    let configCell
    let dataRange
    if ( doLog ) {
      // get config sheet
      try {
        shConfig = ss.getSheetByName( "config" )
        configCell = shConfig.getRange( "D2" )
      } catch(e){
        showPrompt("There doesn't appear to be a sheet called config. This required, and needs to be in the correct format as well.")
        return
      }
    }
    // if formula, continueand remove leading =
    form = form.slice(1)
    // set the name of the sheet and cell address of curent active cell
    // which is where the original formula was
    // make row and column formulas for current cell
    const rowsForm = "=ROWS(" + form + ")"
    const colsForm = "=COLUMNS(" + form + ")"
    activeCell.setFormula( rowsForm )
    const numRows = activeCell.getValue()
    activeCell.setFormula( colsForm )
    const numCols = activeCell.getValue()
    activeCell.setFormula( form )
    
    // get current cell background colour and border style
    let bgColour = activeCell.getBackground()
    let brStyle = activeCell.getBorder().getLeft().getBorderStyle()
    let brColour = activeCell.getBorder().getLeft().getColor().asRgbColor().asHexString()
    if ( doLog ) {
      // get cell formatting from config C2
      bgColour = configCell.getBackground()
      brStyle = configCell.getBorder().getLeft().getBorderStyle()
      brColour = configCell.getBorder().getLeft().getColor().asRgbColor().asHexString()
    }
    // set background of data region
    shData.getRange( cRow, cCol, numRows, numCols ).setBackground( bgColour )
    try {
      // apply border settings to data region on original sheet, or clear borders
      activeCell.setBorder(false, false, false, false, false, false);
      shData.getRange( cRow, cCol, numRows, numCols ).setBorder(true, true, true, true, null, null, brColour, brStyle)
    } catch(e){
      shData.getRange( cRow, cCol, numRows, numCols ).setBorder(false, false, false, false, false, false);
    }
  
  if ( doLog ) {
      shConfig.getRange("A2").setValue( shName + "!" + cCell )
      // check if row on config sheet already exists for this sheet/cell combination
      // if so it returns the row number of the entry
      let rowNumIfExists = shConfig.getRange("B2").getValue()
      const rowData = [ shName, cCell, shName + "!" + cCell, form, numRows, numCols ]
      if ( rowNumIfExists === "" ){ 
        // entry doesn't exist, so append to config sheet
        // add blank entry to array, to leave space for SEQ formula to populate on sheet
        rowData.unshift("")
        shConfig.appendRow( rowData )
        rowNumIfExists = shConfig.getLastRow()
      } else {
        // update config sheet
        shConfig.getRange( rowNumIfExists, 2, 1, 6).setValues( [rowData] )
      }
      shConfig.getRange( "A" + rowNumIfExists ).setBackground( bgColour )
      try {
        shConfig.getRange( "A" + rowNumIfExists ).setBorder(true, true, true, true, null, null, brColour, brStyle)
      } catch(e){
        shConfig.getRange( "A" + rowNumIfExists ).setBorder(false, false, false, false, false, false);
      }
    }
  }
  function showPrompt( msg ){
    const ui = SpreadsheetApp.getUi()
    const result = ui.alert( "Formatting Array Formula Data Region", msg, ui.ButtonSet.OK);
  }