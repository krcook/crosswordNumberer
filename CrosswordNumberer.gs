// code to add the custom menu
function onOpen() {
  const ui = DocumentApp.getUi();
  ui.createMenu('Crosswords')
      .addItem('renumber Crossword', 'numberCrossword')
      .addToUi();
}

function numberCrossword() {
  var doc = DocumentApp.getActiveDocument();
  var element = null;
  var cursor = doc.getCursor();
  if (cursor !=  null)
  {
    element = cursor.getElement();
  }
  else
  {
    var selection = doc.getSelection();
    if (selection != null)
    {
      element = selection.getRangeElements()[0].getElement();
    }
    else
    {
      DocumentApp.getUi().alert("Can't get element for some reason!");
      return;
    }
  }
  while (element.getType() != "TABLE" && element.getType() != "BODY_SECTION"){
    element = element.getParent();
  } 
  // Logger.log(element);

  if (element.getType() != "TABLE") {
    DocumentApp.getUi().alert("Crossword grid not currently selected");
    return;
  }

  var nextNumber = 1;

  var xtable = element.asTable();
  var rowCount = xtable.getNumRows();

  var previousRow = null;
  var row = null;
  var nextRow = xtable.getRow(0);

  for (var rnum = 0; rnum < rowCount; rnum++)
  {
    previousRow = row;
    row = nextRow;
    if (rnum < rowCount - 1)
    {
      nextRow = xtable.getRow(rnum + 1);
    }
    else
    {
      nextRow = null;
    }

    var cellCount = row.getNumCells(); 
    for (var cnum = 0; cnum < cellCount; cnum++)
    {
      var cell = row.getCell(cnum);
      var bg = cell.getBackgroundColor();

      if (bg != null)  // it's a black cell - don't need to do anything  (null is colour = none = white)
      {
        continue;
      }

      var bg_above = '#000000';
      if (previousRow != null)  // not the first row
      {
        bg_above = previousRow.getCell(cnum).getBackgroundColor();
      }

      var bg_below = '#000000';
      if (nextRow != null)  // not the last row
      {
        bg_below = nextRow.getCell(cnum).getBackgroundColor();
      }

      var bg_left = '#000000';
      if (cnum > 0)  // not the first cell
      {
        bg_left = row.getCell(cnum - 1).getBackgroundColor();
      }

      var bg_right = '#000000';
      if (cnum < cellCount - 1)  // not the last cell
      {
        bg_right = row.getCell(cnum + 1).getBackgroundColor();
      }

      var needsNumber = false;

      if (bg_above != null && bg_below == null) // (null is colour = none = white, anything else considered black)
      {
        needsNumber = true;
      }

      if (bg_left != null && bg_right == null) //   (null is colour = none = white, anything else considered black)
      {
        needsNumber = true;
      }

      var cellChildCount = cell.getNumChildren();
      for (var childNum = 0; childNum < cellChildCount; childNum++)
      {
        var para = cell.getChild(childNum).asParagraph();
        if (para != null)
        {
          if (needsNumber)
          {
            para.setText(nextNumber);
            text = para.getChild(0).asText();
            text.setBackgroundColor(null);
            text.setForegroundColor('#000000');
            text.setFontSize(9);
            text.setFontFamily('Arial');
            nextNumber++;
          }
          else
          {
            para.clear();
          }
          break; // don't look for any more paragraphs!
        }
      }
    }
  }
}
