/**
 * Calculate GS1 check digit.
 * @param {string} source Numeric barcode.
 * @return {number} GS1 check digit.
 * @customfunction
*/
function GS1_Check(source) {
  //Convert any input to string
  source = source.toString();

  //Validate input
  var regExp = new RegExp("[^0-9]");
  if (regExp.test(source)) {
    throw "Numeric values only";
  }

  // Loop through each digit to get sum
  var count = 0;
  for (let i = 0; i < source.length; i++) {
    count += parseInt(source.substring(i,i+1));
    //Length is even and even-numbered digit positions (2nd, 4th, 6th, etc.)
    //Length is odd and odd-numbered digit positions (1st, 3rd, 5th, etc.)
    if ((source.length % 2 == 0 && i % 2 != 0) || (source.length % 2 != 0 && i % 2 == 0)) {
      count += 2 * parseInt(source.substring(i,i+1));
    }
  }

  //Calculate check digit
  var Check_Digit = 10 - (count % 10);
  if (Check_Digit == 10) {
    Check_Digit = 0;
  }

  return Check_Digit;
}

CustomFunctions.associate("GS1_CHECK", GS1_Check);

/**
 * Return sparkline barcode format options.
 * @return {string} Barcode format options for sparkline function.
 * @customfunction
*/
function BarcodeOpt() {
  // Javascript translation array: {"charttype","bar";"color1","black";"color2","white"}
  return [["charttype","bar"],["color1","black"],["color2","white"]];
}

CustomFunctions.associate("BARCODEOPT", BarcodeOpt);

/**
 * Generate raw Code 11 barcode.
 * @param {string} source Digits to encode.
 * @return {number} Raw Code 11 barcode.
 * @customfunction
*/
function Code11(source) {
  var Code11chars = "0123456789-";
  var Code11Table = ["111121", "211121", "121121", "221111", "112121",
            "212111", "122111", "111221", "211211", "211111", "112111"];
  
  //Convert any input to string
  source = source.toString();

  //Validate input
  for (let i = 0; i < source.length; i++) {
    if (Code11chars.includes(source.substring(i,i+1)) == false) {
      throw "Invalid character found: "+source.substring(i,i+1);
    }
  }

  //Start characters
  var dest = [[1],[1],[2],[2],[1],[1]];
  //Middle characters
  for (let j = 0; j < source.length; j++) {
    for (let k = 0; k < 6; k++) {
      dest.push([parseInt(Code11Table[parseInt(Code11chars.search(source.substring(j,j+1)))][k])]);
    }
  }
  //End characters
  dest.push([1],[1],[2],[2],[1]);
  return dest;
}

CustomFunctions.associate("CODE11", Code11);
