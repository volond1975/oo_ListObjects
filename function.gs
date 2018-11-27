function RangeIntersect (R1, R2) {
  return (R1.getLastRow() >= R2.getRow()) && (R2.getLastRow() >= R1.getRow()) && (R1.getLastColumn() >= R2.getColumn()) && (R2.getLastColumn() >= R1.getColumn());
}
//Helper Functions
//Going GAS
//From VBA to Google Apps Script
//isObject Test if an item is an object.
function isObject (obj) {
 return obj === Object(obj);
}
//
//isArray Test if an item is an array.
function isArray (arg) {
 return Array.isArray (arg);
}
//
//isUndefined Test if an item is undefined.
function isUndefined ( arg) {
 return typeof arg === typeof undefined;
 }
//
//fixOptional 
//При необходимости добавьте необязательный аргумент, если он отсутствует.
function fixOptional(arg, defaultValue) {
 return isUndefined(arg) ?
 defaultValue : arg;
 }
