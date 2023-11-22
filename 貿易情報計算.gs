// function CopyQuotaitionItem(){
//   let quotationData = quotationSheet.getRange("B14:E").getValues();
//   // 入力されている値のみを取得（nullおよび空文字も含む）
//   let filteredData = quotationData.filter(function(row) {
//     return row.some(cell => cell !== null);
//   });
//   // 貿易情報シートのA列からD列、4行目から始まる列にコピー
//   tradeSheet.getRange(4, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
// }

// function CopyQuotaitionVolume() {  //見積シートの数量を貿易情報の単位にコピー
//   // 見積作成シートのG列14行目以降を取得
//   let quotationData = quotationSheet.getRange("G14:G").getValues();
  
//   // 入力されている値のみを取得
//   let filteredData = quotationData.filter(function(row) {
//     return row[0] !== "";
//   });
  
//   // 貿易情報シートのJ列4行目から始まる列にコピー
//   tradeSheet.getRange(4, 10, filteredData.length, 1).setValues(filteredData);
// }

// function CopyCostPrice() {  //見積シート単価（原価）F列にをコピー
//   // 見積作成シートのG列14行目以降を取得
//   let quotationData = quotationSheet.getRange("J14:J").getValues();
  
//   // 入力されている値のみを取得
//   let filteredData = quotationData.filter(function(row) {
//     return row[0] !== "";
//   });
  
//   // 貿易情報シートのF列4行目から始まる列にコピー
//   tradeSheet.getRange(4, 6, filteredData.length, 1).setValues(filteredData);
// }

function TotalCostOfGoods() {  //商品代合計　OK
  let lastRow = tradeSheet.getLastRow();
  // 4行目から最終行までのF列とJ列の値を取得する。
  let fValues = tradeSheet.getRange(4, 6, lastRow - 3, 1).getValues();
  let jValues = tradeSheet.getRange(4, 10, lastRow - 3, 1).getValues();
  console.log(fValues, 'fValues');
  console.log(jValues, 'jValues');
   for (let i = 0; i < fValues.length; i++) {
      let jValue = jValues[i][0]; // Column J
      let fValue = fValues[i][0]; // Column F
      // JとFに値があるかどうかをチェックし、Gを計算して更新する。
      if (jValue !== "" && fValue !== "") {
       tradeSheet.getRange(4 + i, 7).setValue(jValue * fValue); // Column G
    }
  }
}

function OverseasFareSetValues() {   //海上運賃 OK
  let lastRow = tradeSheet.getLastRow();

  // R列～X列の2行目の値を取得
  let fixedValues = [];
  for (col = 18; col <= 24; col++) {  // 列R（18）から列X（24）まで
    fixedValues[col] = tradeSheet.getRange(2, col).getValue();
  }

  // K列の4行目以降の値を取得
  let kValues = tradeSheet.getRange(4, 11, lastRow - 3, 1).getValues();  // 列K（11）から始まり、4行目以降
  
  // 各列の2行目の値とK列の4行目以降の値を掛け算して、その結果を各列の4行目以降に表示
  for (col = 18; col <= 24; col++) {  // 列R（18）から列X（24）まで
    for (row = 4; row <= lastRow; row++) {
      let result = fixedValues[col] * kValues[row - 4][0];
      // 結果を表示
      tradeSheet.getRange(row, col).setValue(result);
    }
    //R列からAB列までの各行の値の合計を計算する。
    for (let row = 4; row <= lastRow; row++) {
      let sumResult = 0;
      for (let col = 18; col <= 28; col++) {  // 列R（18）から列AB（28）まで
        let value = tradeSheet.getRange(row, col).getValue();
          sumResult += value !== "" ? value : 0;
      }
        // 合計結果をH列に表示
        tradeSheet.getRange(row, 8).setValue(sumResult);
    }
  }
}

function CustomsClearanceChargesSetValues() {   //通関費 OK
  let lastRow = tradeSheet.getLastRow();

  // AC列～AI列の2行目の値を取得
  let fixedValues = [];
    for (col = 33; col <= 35; col++) {  // 列AG（33）から列AI（35）まで
      fixedValues[col] = tradeSheet.getRange(2, col).getValue();
    }
      // G列の4行目以降の値を取得
      let gValues = tradeSheet.getRange(4, 7, lastRow - 3, 1).getValues();  // 列G（7）から始まり、4行目以降
  
      // 各列の2行目の値とK列の4行目以降の値を掛け算して、その結果を各列の4行目以降に表示
    for (col = 33; col <= 35; col++) {  
      for (row = 4; row <= lastRow; row++) {
        let result = fixedValues[col] * gValues[row - 4][0];
        // 結果を表示
        tradeSheet.getRange(row, col).setValue(result);
    }
  }
    //AC列からAI列までの各行の値の合計を計算する。
    for (let row = 4; row <= lastRow; row++) {
      let sumResult = 0;
      for (let col = 29; col <= 35; col++) {  // 列AC（29）から列AI（35）まで
        let value = tradeSheet.getRange(row, col).getValue();
            sumResult += value !== "" ? value : 0;
      }
        // 合計結果をI列に表示
        tradeSheet.getRange(row, 9).setValue(sumResult);
    }
}  

function CostPrice(){   //原価
  let lastRow = tradeSheet.getLastRow();
    for (let row = 4; row <= lastRow; row++) {
      let sum = 0;
      //G列からI列の値を合計する
      for (let col = 7; col <= 9; col++) {
        let value = tradeSheet.getRange(row, col).getValue();
        // Check if the value is a number before adding it to the sum
        if (!isNaN(value)) {
          sum += value;
        }
      }
      // M列に合計を表示
      tradeSheet.getRange(row, 13).setValue(sum);
    }
}

function SellingPrice(){    //売値
  let lastRow = tradeSheet.getLastRow();
  // 4行目から最終行までのN列とJ列の値を取得する。
  let nValues = tradeSheet.getRange(4, 14, lastRow - 3, 1).getValues();
  let jValues = tradeSheet.getRange(4, 10, lastRow - 3, 1).getValues();
   for (let i = 0; i < nValues.length; i++) {
      let nValue = nValues[i][0];
      let jValue = jValues[i][0]; 
      // JとNに値があるかどうかをチェックし、Oを計算して更新する。
      if (jValue !== "" && nValue !== "") {
       tradeSheet.getRange(4 + i, 15).setValue(nValue * jValue);
      }
    }
}

function GrossProfit(){    //粗利
  let lastRow = tradeSheet.getLastRow();
  // 4行目から最終行までのM列とJ列の値を取得する。
  let oValues = tradeSheet.getRange(4, 15, lastRow - 3, 1).getValues();
  let mValues = tradeSheet.getRange(4, 13, lastRow - 3, 1).getValues();
   for (let i = 0; i < oValues.length; i++) {
      let oValue = oValues[i][0];
      let mValue = mValues[i][0];
      // JとNに値があるかどうかをチェックし、Oを計算して更新する。
      if (oValue !== "" && mValue !== "") {
       tradeSheet.getRange(4 + i, 16).setValue(oValue - mValue);
     }
  }
}

function GrossMarginRatio(){    //粗利率
  let lastRow = tradeSheet.getLastRow();
  // 4行目から最終行までのN列とJ列の値を取得する。
  let pValues = tradeSheet.getRange(4, 16, lastRow - 3, 1).getValues();
  let oValues = tradeSheet.getRange(4, 15, lastRow - 3, 1).getValues();
   for (let i = 0; i < pValues.length; i++) {
      let pValue = pValues[i][0]; 
      let oValue = oValues[i][0]; 
      // PとOに値があるかどうかをチェックし、Qを計算して更新する。
      if (pValue !== "" && oValue !== "") {
       tradeSheet.getRange(4 + i, 17).setValue(pValue / oValue);

    }
  }
}
