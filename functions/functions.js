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

/**
 * Generate raw Code 39 barcode.
 * @param {string} source  Text to encode.
 * @param {boolean} CHECK_DIGIT [OPTIONAL] Add check digit.
 * @return {number} Raw Code 39 barcode.
 * @customfunction
*/
function Code39(source, CHECK_DIGIT = false) {
  var Code39chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%";
  var Code39Table = ["1112212111", "2112111121", "1122111121", "2122111111", "1112211121",
    "2112211111", "1122211111", "1112112121", "2112112111", "1122112111", "2111121121",
    "1121121121", "2121121111", "1111221121", "2111221111", "1121221111", "1111122121",
    "2111122111", "1121122111", "1111222111", "2111111221", "1121111221", "2121111211",
    "1111211221", "2111211211", "1121211211", "1111112221", "2111112211", "1121112211",
    "1111212211", "2211111121", "1221111121", "2221111111", "1211211121", "2211211111",
    "1221211111", "1211112121", "2211112111", "1221112111", "1212121111", "1212111211",
    "1211121211", "1112121211"];

  //Convert any input to string
  source = source.toString();

  //Convert all letters to uppercase
  source = source.toUpperCase();

  //Validate input
  for (let i = 0; i < source.length; i++) {
    if (Code39chars.includes(source.substring(i,i+1)) == false) {
      throw "Invalid character found: "+source.substring(i,i+1);
    }
  }

  var count = 0;
  //Start character (asterisk)
  var dest = [[1],[2],[1],[1],[2],[1],[2],[1],[1],[1]];
  //Middle characters
  for (let j = 0; j < source.length; j++) {
    for (let k = 0; k < 10; k++) {
      dest.push([parseInt(Code39Table[parseInt(Code39chars.search(source.substring(j,j+1)))][k])]);
      count += parseInt(Code39chars.search(source.substring(j,j+1)));
    }
  }
  //Check digit (Code 39 mod 43)
  if (CHECK_DIGIT == true) {
    count %= 43;
    for (let l = 0; l < 10; l++) {
      dest.push(parseInt(Code39Table[count][l]));
    }
  }
  //End character (asterisk)
  dest.push([1],[2],[1],[1],[2],[1],[2],[1],[1],[1]);
  return dest;
}

CustomFunctions.associate("CODE39", Code39);

/**
 * Generate raw Code 93 barcode.
 * @param {string} source Text to encode.
 * @return {number} Raw Code 93 barcode.
 * @customfunction
*/
function Code93(source) {
  var Code93chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%";
  var Code93Table = ["131112", "111213", "111312", "111411", "121113", "121212", "121311",
        "111114", "131211", "141111", "211113", "211212", "211311", "221112", "221211", "231111",
        "112113", "112212", "112311", "122112", "132111", "111123", "111222", "111321", "121122",
        "131121", "212112", "212211", "211122", "211221", "221121", "222111", "112122", "112221",
        "122121", "123111", "121131", "311112", "311211", "321111", "112131", "113121", "211131",
        "121221", "312111", "311121", "122211"];
  
  //Convert any input to string
  source = source.toString();

  //Convert all letters to uppercase
  source = source.toUpperCase();

  //Validate input
  for (let i = 0; i < source.length; i++) {
    if (Code93chars.includes(source.substring(i,i+1)) == false) {
      throw "Invalid character found: "+source.substring(i,i+1);
    }
  }
  
  //Start character
  var dest = [[1],[1],[1],[1],[4],[1]];
  //Calculate check digit C
  var c = 0;
  var weight = 1;
  for (let j = source.length-1; j > -1; j--) {
    c += parseInt(Code93chars.search(source.substring(j,j+1))) * weight;
    weight += 1;
    if (weight == 21) {
      weight = 1;
    }
  }
  c %= 47;
  source += Code93chars.substring(c,c+1);
  //Calculate check digit K
  var k = 0;
  weight = 1;
  for (let l = source.length-1;l > -1; l--) {
    k += parseInt(Code93chars.search(source.substring(l,l+1))) * weight;
    weight += 1;
    if (weight == 16) {
      weight = 1;
    }
  }
  k %= 47;
  source += Code93chars.substring(k,k+1);
  //Middle characters with check digits C and K
  for (let m = 0; m < source.length; m++) {
    for (let n = 0; n < 6; n++) {
      dest.push([parseInt(Code93Table[parseInt(Code93chars.search(source.substring(m,m+1)))][n])]);
    }
  }
  //End character
  dest.push([1],[1],[1],[1],[4],[1],[1]);
  return dest;
}

CustomFunctions.associate("CODE93", Code93);

var ITFtable = ["11221", "21112", "12112", "22111", "11212",
              "21211", "12211", "11122", "21121", "12121"];

/**
 * Generate raw ITF barcode.
 * @param {string} source Even number of digits to encode.
 * @return {number} Raw ITF barcode.
 * @customfunction
*/
function ITF(source) {  
  //Convert any input to string
  source = source.toString();

  //Validate input
  var regExp = new RegExp("[^0-9]");
  if (regExp.test(source)) {
    throw "Numeric values only";
  }

  //ITF requires an even number of digits. If odd, add a zero to the beginning
  if (source.length % 2 != 0) {
    source = "0" + source;
  }

  //Start character
  var dest = [[1],[1],[1],[1]];
  //Middle characters
  for (let i = 0; i < source.length; i+=2) {
    //Interleave 2 digits at a time (1st digit is bars, 2nd digit is spaces)
    for (let j = 0; j < 5; j++) {
      dest.push([parseInt(ITFtable[source.substring(i,i+1)][j])]);
      dest.push([parseInt(ITFtable[source.substring(i+1,i+2)][j])]);
    }
  }
  //End characters
  dest.push([2],[1],[1]);
  return dest;
}

CustomFunctions.associate("ITF", ITF);

/**
 * Generate raw ITF-14 barcode.
 * @param {string} source Even number of digits to encode.
 * @return {number} Raw ITF-14 barcode.
 * @customfunction
*/
function ITF_14(source) {
  //Convert any input to string
  source = source.toString();

  //Validate input
  var regExp = new RegExp("[^0-9]");
  if (regExp.test(source)) {
    throw "Numeric values only";
  }
  if (source.length < 13 || source.length > 14)  {
    throw "Improper ITF-14 barcode length (13-14 digits)";
  } else if (source.length == 14 && GS1_Check(parseInt(source.substring(0,13))) != parseInt(source.substring(13,14))) {
    throw "Invalid check digit ("+GS1_Check(parseInt(source.substring(0,13)))+")";
  }

  //Calculate check digit
  if (source.length == 13) {
    source += GS1_Check(parseInt(source.substring(0,13)));
  }

  //Start character
  var dest = [[1],[1],[1],[1]];
  //Middle characters
  for (let i = 0; i < source.length; i+=2) {
    //Interleave 2 digits at a time (1st digit is bars, 2nd digit is spaces)
    for (let j = 0; j < 5; j++) {
      dest.push([parseInt(ITFtable[source.substring(i,i+1)][j])]);
      dest.push([parseInt(ITFtable[source.substring(i+1,i+2)][j])]);
    }
  }
  //End characters
  dest.push([2],[1],[1]);
  return dest;
}

CustomFunctions.associate("ITF_14", ITF_14);

//Global Variables
var UPCParity0 = ["BBBAAA", "BBABAA", "BBAABA", "BBAAAB", "BABBAA", "BAABBA", "BAAABB", 
                  "BABABA", "BABAAB", "BAABAB"]; //Number set for UPC-E symbol (EN Table 4)
var UPCParity1 = ["AAABBB", "AABABB", "AABBAB", "AABBBA", "ABAABB", "ABBAAB", "ABBBAA", 
                  "ABABAB", "ABABBA", "ABBABA"]; //Not covered by BS EN 797:1995
var EAN2Parity = ["AA", "AB", "BA", "BB"]; //Number sets for 2-digit add-on (EN Table 6)
var EAN5Parity = ["BBAAA", "BABAA", "BAABA", "BAAAB", "ABBAA", "AABBA", "AAABB", "ABABA", 
                  "ABAAB", "AABAB"]; //Number set for 5-digit add-on (EN Table 7)
var EAN13Parity = ["AAAAA", "ABABB", "ABBAB", "ABBBA", "BAABB", "BBAAB", "BBBAA", "BABAB", 
                  "BABBA", "BBABA"]; //Left hand of the EAN-13 symbol (EN Table 3)
var EANsetA = ["3211", "2221", "2122", "1411", "1132", "1231", "1114", "1312", "1213", 
              "3112"]; //Representation set A and C (EN Table 1)
var EANsetB = ["1123", "1222", "2212", "1141", "2311", "1321", "4111", "2131", "3121", 
              "2113"]; //Representation set B (EN Table 1)

/**
 * Generate raw EAN-8 or UPC-A barcode.
 * @param {string} source  Digits to encode (EAN-8 is 8, UPC-A is 11-12).
 * @return {number} Raw EAN-8 or UPC-A barcode.
 * @customfunction
*/
function UPCA(source) {
  //Convert any input to string
  source = source.toString();

  //Validate input
  var regExp = new RegExp("[^0-9]");
  if (regExp.test(source)) {
    throw "Numeric values only";
  }

  if (source.length != 11 && source.length != 12 && source.length != 8) {
    throw "Improper EAN-8 or UPC-A barcode length (8 or 11-12 digits)";
  } else if (source.length == 12 && GS1_Check(parseInt(source.substring(0,11))) != parseInt(source.substring(11,12))) {
    throw "Invalid check digit ("+GS1_Check(parseInt(source.substring(0,11)))+")";
  }

  //Calculate check digit (UPC-A only)
  if (source.length == 11) {
    source += GS1_Check(parseInt(source.substring(0,11)));
  }

  var half_way = source.length / 2;
  //Start characters
  var dest = [[1],[1],[1]];
  //Middle characters
  for (let i = 0; i < source.length; i++) {
    if (i == half_way) {
      dest.push([1],[1],[1],[1],[1]);
    }
    for (let j = 0; j < 4; j++) {
      dest.push([parseInt(EANsetA[source.substring(i,i+1)][j])]);
    }
  }
  //End characters
  dest.push([1],[1],[1]);
  return dest;
}

CustomFunctions.associate("UPCA", UPCA);

/**
 * Generate raw UPC-E barcode.
 * @param {string} source  Digits to encode.
 * @return {number} Raw UPC-E barcode.
 * @customfunction
*/
function UPCE(source) {
  //Convert any input to string
  source = source.toString();

  //Validate input
  var regExp = new RegExp("[^0-9]");
  if (regExp.test(source)) {
    throw "Numeric values only";
  }

  //Two number systems can be used - system 0 and system 1
  if (source.length == 7) {
    if (source.substring(0,1) > 1) {
      source = "0" + source.substring(1,7);
    }
  } else if(source.length == 6) {
    //Default number system is 0
    source = "0" + source;
  } else {
    throw "Improper UPC-E barcode length (6-7 digits)";
  }
  
  //Expand the zero-compressed UPCE code to make a UPCA equivalent (EN Table 5)
  var emode = source.substring(6,7);
  var equivalent = source.substring(0,3);
  switch(parseInt(emode)) {
    case 0:
    case 1:
    case 2:
      equivalent += emode+"0000"+source.substring(3,6);
      break;
    case 3:
      equivalent += source.substring(3,4)+"00000"+source.substring(4,6);
      break;
    case 4:
      equivalent += source.substring(3,5)+"00000"+source.substring(5,6);
      break;
    case 5:
    case 6:
    case 7:
    case 8:
    case 9:
      equivalent += source.substring(3,6)+"0000"+emode;
      break;
  }
  
  //Calculate check digit
  var CHECK_DIGIT = GS1_Check(equivalent);
  equivalent += CHECK_DIGIT;
  
  //Use number system and check digit to choose a parity scheme
  var parity;
  if (equivalent.substring(0,1) == 1) {
    parity = UPCParity1[CHECK_DIGIT];
  } else {
    parity = UPCParity0[CHECK_DIGIT];
  }
  
  //Start characters
  var dest = [[1],[1],[1]];
  //Middle characters
  for (let i = 0; i < parity.length; i++) {
    for (let j = 0; j < 4; j++) {
      if (parity.substring(i,i+1) == "A") {
        dest.push([parseInt(EANsetA[source.substring(i+1,i+2)][j])]);
      } else {
        dest.push([parseInt(EANsetB[source.substring(i+1,i+2)][j])]);
      }
    }
  }
  //End characters
  dest.push([1],[1],[1],[1],[1],[1]);
  return dest;
}

CustomFunctions.associate("UPCE", UPCE);

/**
 * Generate raw EAN-13 barcode.
 * @param {string} source  Digits to encode.
 * @return {number} Raw EAN-13 barcode.
 * @customfunction
*/
function EAN_13(source) {
  //Convert any input to string
  source = source.toString();

  //Validate input
  var regExp = new RegExp("[^0-9]");
  if (regExp.test(source)) {
    throw "Numeric values only";
  }

  if (source.length < 12 || source.length > 13) {
    throw "Improper EAN-13 barcode length (12-13 digits)";
  } else if (source.length == 13 && GS1_Check(parseInt(source.substring(0,12))) != parseInt(source.substring(12,13))) {
    throw "Invalid check digit ("+GS1_Check(parseInt(source.substring(0,12)))+")";
  }

  //Calculate check digit
  if (source.length == 12) {
    source += GS1_Check(source);
  }
  
  //Get parity for first half of symbol
  var parity = EAN13Parity[source.substring(0,1)];
  var half_way = 7;

  //Start characters
  var dest = [[1],[1],[1]];
  //Middle characters
  for (let i = 1; i < source.length; i++) {
    if (i == half_way) {
      dest.push([1],[1],[1],[1],[1]);
    }
    for (let j = 0; j < 4; j++) {
      if (i > 1 && i < half_way) {
        if (parity.substring(i-2,i-1) == "A") {
          dest.push([parseInt(EANsetA[source.substring(i,i+1)][j])]);
        } else {
          dest.push([parseInt(EANsetB[source.substring(i,i+1)][j])]);
        }
      } else {
        dest.push([parseInt(EANsetA[source.substring(i,i+1)][j])]);
      }
    }
  }
  //End characters
  dest.push([1],[1],[1]);
  return dest;
}

CustomFunctions.associate("EAN_13", EAN_13);

/**
 * Generate raw EAN-5 barcode.
 * @param {string} source  Digits to encode.
 * @return {number} Raw EAN-5 barcode.
 * @customfunction
*/
function EAN_5(source) {
  //Convert any input to string
  source = source.toString();

  //Validate input
  var regExp = new RegExp("[^0-9]");
  if (regExp.test(source)) {
    throw "Numeric values only";
  }

  if (source.length != 5) {
    throw "Improper EAN-5 barcode length (5 digits)";
  }

  //Determine parity
  var parity_sum = 3 * (parseInt(source.substring(0,1)) + parseInt(source.substring(2,3))+parseInt(source.substring(4,5)));
  parity_sum += 9 * (parseInt(source.substring(1,2)) + parseInt(source.substring(3,4)));
  parity_sum %= 10;
  var parity =  EAN5Parity[parity_sum];

  //Start characters
  var dest = [[1],[1],[2]];
  //Middle characters
  for (let i = 0; i < parity.length; i++) {
    for (let j = 0; j < 4; j++) {
      if (parity.substring(i,i+1) == "A") {
        dest.push([parseInt(EANsetA[source.substring(i,i+1)][j])]);
      } else {
        dest.push([parseInt(EANsetB[source.substring(i,i+1)][j])]);
      }
    }
    if (i < parity.length-1) {
      dest.push([1],[1]);
    }
  }
  return dest;
}

CustomFunctions.associate("EAN_5", EAN_5);

/**
 * Generate raw EAN-2 barcode.
 * @param {string} source  Digits to encode.
 * @return {number} Raw EAN-2 barcode.
 * @customfunction
*/
function EAN_2(source) {
  //Convert any input to string
  source = source.toString();

  //Validate input
  var regExp = new RegExp("[^0-9]");
  if (regExp.test(source)) {
    throw "Numeric values only";
  }

  if (source.length != 2) {
    throw "Improper EAN-2 barcode length (2 digits)";
  }

  //Determine parity
  var parity_sum = (parseInt(source.substring(0,1)) * 10) + parseInt(source.substring(1,2));
  parity_sum %= 4;
  parity = EAN2Parity[parity_sum];

  //Start characters
  var dest = [[1],[1],[2]];
  //Middle characters
  for (let i = 0; i < parity.length; i++) {
    for (let j = 0; j < 4; j++) {
      if (parity.substring(i,i+1) == "A") {
        dest.push([parseInt(EANsetA[source.substring(i,i+1)][j])]);
      } else {
        dest.push([parseInt(EANsetB[source.substring(i,i+1)][j])]);
      }
    }
    if (i < parity.length-1) {
      dest.push([1],[1]);
    }
  }
  return dest;
}

CustomFunctions.associate("EAN_2", EAN_2);
