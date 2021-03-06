// Loading Library
const ExcelHandler = require('ExcelHandler');

function UpdateProdOrderQuantity(){
  // Instantiation Step
  let ExcelHdl = new ExcelHandler('ExcelFile.xlsx').open().sheet('DATA').rowStartAt(2).cols({
    "Updated": "F"
  });

  // Dynamic methods available from header line :
  let Warehouse = ExcelHdl.Warehouse(2); // Col A -> Warehouse -> line 2 = SA1
  let Material  = ExcelHdl.Material(2);  // Col B -> Material  -> line 2 = K00289

  // This one, Method has been adjusted to meet JavaScript Function naming convention rule
  // by trimming space
  let ProdOrd   = ExcelHdl.ProdOrd(2);   // Col C -> Prod Ord  -> Line 2 = 20001467

  // Updating Cell
  ExcelHdl.Quantity(2, 12);   // Col E -> Quantity -> Line 2 become 12 instead of 6
  
  // Set Update Date
  ExcelHdl.Updated(2, new Date());

  // Save modifications
  ExcelHdl.save();
}