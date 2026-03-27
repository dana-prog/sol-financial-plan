function onEdit(event) {
  SOLLibrary.logArgs('Main', 'onEvent', {event: event});
}

function onTestEdit(event) {
  const constructionTypesRange = getConstructionCostsTypeColRange();

  // if (!SOLLibrary.isInside(event.range, constructionTypesRange)) {
  //   return;
  // }

  SOLLibrary.logArgs('Main', 'onEdit', {event}, 'onEdit inside construction types range');

  const oldUnitType = event.oldValue;
  const newUnitType = event.value || '';

  if (oldUnitType === newUnitType) {
    return; // TODO: check if this condition is needed
  }

  const newUnitTypes = getUnitTypes();
  const oldUnitTypes = newUnitTypes.map(type => type === newUnitType ? oldUnitType : type);
  SOLLibrary.logArgs('Main', 'onEdit', {
    oldUnitType,
    newUnitType,
    oldTypes: oldUnitTypes,
    newTypes: newUnitTypes
  });

  validateUnitTypesParamsSync(oldUnitTypes);
  SOLLibrary.log('Main', 'onEdit', 'finished validation')

  if (oldUnitType !== newUnitType && oldUnitType !== '' && newUnitType !== '') {
    SOLLibrary.log('Main', 'onEdit', 'calling update');
    updateTimelineUnitTypeParams(oldUnitType, newUnitType);
    return;
  }

  const typeIndex = oldUnitTypes.indexOf(oldUnitType);
  if (typeIndex === -1) {
    SOLLibrary.logArgs('Main', 'onEdit', {
      newValue: newUnitType,
      oldValue: oldUnitType,
      oldValues: oldUnitTypes
    }, `Param '${oldUnitType}' was not found.`);
    throw new Error(`${oldUnitType} was not found.`);
  }
  if (oldUnitType === '') {
    addTimelineConstructionTypeParams(newUnitType, typeIndex + 1);
    return;
  }

  if (newUnitType === '') {
    SOLLibrary.logArgs('Main', 'onEdit', {oldUnitType}, 'calling removeTimelineConstructionTypeParams');
    removeTimelineConstructionTypeParams(typeIndex);
  }
}

/**
 * An event handler called when the spreadsheet is opened. Initializes the SOL menu.
 */
// noinspection JSUnusedGlobalSymbols
function onOpen() {
  _createMenu();
}

/**
 * Creates the SOL menu.
 * @private
 */
function _createMenu() {
  const ui = SpreadsheetApp.getUi();


  ui.createMenu('SOL')
    .addItem('Export as XLSX (Values Only)', '_exportValuesXSLX')
    .addToUi();
}

function _exportValuesXSLX() {
  SOLLibrary.exportValuesXSLX('deleteTmpExportResources');
}

// do not remove this line. see SOLLibrary.exportValuesXSLX docs for details
const deleteTmpExportResources = SOLLibrary.deleteTmpExportResources;
