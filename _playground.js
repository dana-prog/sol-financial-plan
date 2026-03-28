function test() {
  const unitTypes = getUnitTypes();
  const res = {};
  for (let i = 0; i < unitTypes.length; i++) {
    const unitType = unitTypes[i];
    res[unitType] = {
      costs: getConstructionCostsTotalUnitCount(unitType),
      timeline: getTimelineTotalUnitCount(unitType)
    };
  }
  SOLLibrary.logArgs('_playground', 'test', {res});
}