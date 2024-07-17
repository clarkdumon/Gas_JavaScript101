const romanDict = {
    M: 1000,
    CM: 900,
    D: 500,
    CD: 400,
    C: 100,
    XC: 90,
    L: 50,
    XL: 40,
    X: 10,
    IX: 9,
    V: 5,
    IV: 4,
    I: 1
};

function romanToNumber(roman) {

let total = 0;

for(let i=0, i<roman.lenth,i++){
    if{romanDict[roman[i]]<romanDict[roman[i+1]]}

    total -= romanDict[roman[i]];

}else{
    total +=romanDict[roman[i]];
}
return total;

}
