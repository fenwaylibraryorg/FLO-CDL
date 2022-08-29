function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var menuEntries = [ 
    {name: "Update List of Reserves", functionName: "listFilesAndFolders"},
  ];
  ss.addMenu("Update Reserves", menuEntries);
}
