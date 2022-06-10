#!/usr/bin/env node

clog = console.log;
cdir = function ($val) {
    console.dir($val, {
        depth: null
    })
};

const TableJs = require("./bin/TableJs");



// -------------------------------------------------------------------------
// ---[ I ]-----------------------------------------------------------------
// -------------------------------------------------------------------------
clog("-------------------------------------------------------");
console.log(">>>> Initializing a new TableJs: \n\n");

let cameras = new TableJs(
    // List of Fields
    ['Brand', 'Camera', 'Date', 'Format', 'Purpose'],

    // Indicating which field compose the keys
    ['Camera', 'Brand'],

    // Table Data
    [
        [ 'Nikon', 'D3', '2007', 'Full Frame', 'Professional' ],
        [ 'Nikon', 'D750', '2014', 'Full Frame', 'Action' ],
        [ 'Nikon', 'D800', '2011', 'Full Frame', 'Semi-Professional' ],
        [ 'Nikon', 'D810A', '2015', 'Full Frame', 'Astro' ],
        [ 'Nikon', 'D7100', '2013', 'APS-C', 'Expert' ],
        [ 'Nikon', 'D6', '2020', 'Full Frame', 'Professional' ],
        [ 'Canon', '1Ds Mark III', '2007', 'Full Frame', 'Professional' ],
        [ 'Canon', '5D Mark II', '2008', 'Full Frame', 'Semi-Professional' ],
        [ 'Canon', '60Da', '2012', 'APS-C', 'Astro' ],
        [ 'Canon', '250D', '2019', 'APS-C', 'Compact' ],
    ]
);
// console.log(cameras);
// return false;
clog("\n-------------------------------------------------------");



// -------------------------------------------------------------------------
// ---[ II ]----------------------------------------------------------------
// -------------------------------------------------------------------------
console.log(">>>> Get distinct values: \n\n");

console.log('Brand List:', cameras.Brand());
console.log('Camera List:', cameras.Camera());
console.log('Date List:', cameras.Date());
console.log('Format List:', cameras.Format());
console.log('Purpose List:', cameras.Purpose());
clog("\n-------------------------------------------------------");



// -------------------------------------------------------------------------
// ---[ III ]---------------------------------------------------------------
// -------------------------------------------------------------------------
console.log(">>>> Get rows: \n\n");

let NikonProCam = cameras.Brand('Nikon').Purpose('Professional');
console.log('Table Result:', NikonProCam);
console.log('New Camera List:', NikonProCam.Camera());
clog("\n-------------------------------------------------------");



// -------------------------------------------------------------------------
// ---[ IV ]----------------------------------------------------------------
// -------------------------------------------------------------------------
console.log(">>>> Get field value of row: \n\n");

// Return the camera name of the first entry
let cameraName = NikonProCam[0].Camera();
console.log("Camera name of the first row:", cameraName);


NikonProCam.forEach(function($row){
    console.log('Camera: ', $row.Camera());
});
clog("\n-------------------------------------------------------");



// -------------------------------------------------------------------------
// ---[ V ]-----------------------------------------------------------------
// -------------------------------------------------------------------------
console.log(">>>> Appending new rows: \n\n");

cameras.push('Sony Alpha');
console.log("Updated Cameras table:", cameras);

cameras.push([
    "Sony Alpha", "α 9 II", "2019", "Full Frame", "Sport-Pro"
]);
console.log("Updated 2 Cameras table:", cameras);
console.log("Updated Cameras List:", cameras.Camera());

clog("\n-------------------------------------------------------");



// -------------------------------------------------------------------------
// ---[ VI ]----------------------------------------------------------------
// -------------------------------------------------------------------------
console.log(">>>> Setting (Updating) field value: \n\n");

let sonyCamera = cameras.Brand('Sony Alpha');

// Update Row where there is no Camera Name
let inc = 1;
sonyCamera.Camera('').forEach(function($row){
    // Camera field is a component of key,
    // So we have to set a "unique" name for brand (here sony)
    $row.Camera(`Camera ${inc}`);
    inc++;
});

console.log("Updated Sony Alpha Camera Table", cameras.Brand('Sony Alpha'));
console.log("Updated Camera Table", cameras);

clog("\n-------------------------------------------------------");



// -------------------------------------------------------------------------
// ---[ VII ]---------------------------------------------------------------
// -------------------------------------------------------------------------
console.log(">>>> Setting (Updating) field value of result of rows: \n\n");

cameras.Format('Full Frame').update({
    Format: 'Full Frame(24x36)'
});
console.log("Updated Table:", cameras);

console.log("\n-------------------------------------------------------");



// -------------------------------------------------------------------------
// ---[ VIII ]--------------------------------------------------------------
// -------------------------------------------------------------------------
console.log(">>>> Make a copy of table: \n\n");

let D3Cam  = cameras.Camera('D3');          // Table with 1 row
let D3Cam2 = D3Cam;                         // This is not a copy
D3Cam2[0].Camera('D3Rename');               // Rename the camera name

console.log("D3Cam:  ", D3Cam);             // D3 --> D3Rename
console.log("D3Cam2: ", D3Cam2);            // D3 --> D3Rename
console.log("Cameras Table: ", cameras);    // D3 --> D3Rename


let D6Cam  = cameras.Camera('D6');          // Table with 1 row
let D6Cam2 = D6Cam.copy();                  // Make a true copy of the table
D6Cam2[0].Camera('D6Rename');               // Rename the camera name

console.log("D6Cam:  ", D6Cam);             // D6 --> D6
console.log("D6Cam2: ", D6Cam2);            // D6 --> D6Rename
console.log("Cameras Table: ", cameras);    // D6 --> D6

console.log("\n-------------------------------------------------------");



// -------------------------------------------------------------------------
// ---[ IX ]----------------------------------------------------------------
// -------------------------------------------------------------------------
console.log(">>>> Delete rows from the table: \n\n");

cameras.Format('APS-C').delete();
console.log(cameras);

console.log("\n-------------------------------------------------------");
