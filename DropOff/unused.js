
function extractVehicleDriverHelper(text) {
  // Use regular expressions to extract vehicle, driver, and helper
  var vehicleMatch = text.match(/BUS (\d+) - PLATE \(([\w\d]+)\) - DRIVER: ([\w\s]+) - HELPER: ([\w\s]+)/);

  if (vehicleMatch) {
    var vehicleNumber = vehicleMatch[1];
    var vehiclePlate = vehicleMatch[2];
    var driverName = vehicleMatch[3];
    var helperName = vehicleMatch[4];

    return {
      vehicleNumber: vehicleNumber,
      vehiclePlate: vehiclePlate,
      driverName: driverName,
      helperName: helperName
    };
  } else {
    // Return null or an empty object if the text doesn't match the expected format
    return null;
  }
}