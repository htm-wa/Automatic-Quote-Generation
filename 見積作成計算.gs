function UnitPriceGrossMargin() {
  // F列（6列目）が編集された場合かつ14行目以降である場合
  if (editedColumn == 6 && editedRange.getRow() >= 14) {
    let fValue = editedRange.getValue();
    // 数字以外が入力された場合はエラーを表示
    if (isNaN(fValue)) {
      sheet.toast("数字以外が入力されました。", "エラー", 5);
      return;
    }
    let kValue = quotationSheet.getRange(editedRange.getRow(), 11).getValue(); // K列の値を取得
    let result = ((fValue * kValue) / fValue) * 100;  //粗利率＝（（単価-原価）/単価）×100
    
    // M列に計算結果を表示
    quotationSheet.getRange(editedRange.getRow(), 13).setValue(result); // M列は13列目
  }
  // M列（13列目）が編集された場合かつ14行目以降である場合
  if (editedColumn == 13 && editedRange.getRow() >= 14) {
    let mValue = editedRange.getValue();
    // 数字以外が入力された場合はエラーを表示
    if (isNaN(mValue)) {
      sheet.toast("数字以外が入力されました。", "エラー", 5);
      return;
    }
    let kValue = quotationSheet.getRange(editedRange.getRow(), 11).getValue(); // K列の値を取得
    // 計算
    let result = (kValue / (1 - mValue / 100));   //単価＝原価/（1-（粗利率/100））

    // F列に計算結果を表示
    quotationSheet.getRange(editedRange.getRow(), 6).setValue(result); // F列は6列目
  }
}

function SellingPriceTotal() {  //売値合計
  let lastRow = quotationSheet.getLastRow();

  // F列とG列の14行目以降の各行の値を取得
  let fValues = quotationSheet.getRange("F14:F" + lastRow).getValues();
  let gValues = quotationSheet.getRange("G14:G" + lastRow).getValues();

  // 計算結果を格納する配列
  let results = [];

  // 各行の計算結果を求めてresults配列に格納
  for (i = 0; i < fValues.length; i++) {
    let result = fValues[i][0] * gValues[i][0];
    results.push([result]);
  }
  quotationSheet.getRange("I14:I" + (14 + results.length - 1)).setValues(results);
  
  // J列14行目以降の各行の値を取得
  let jRange = quotationSheet.getRange("J14:J" + (14 + results.length - 1));
  let jValues = jRange.getValues();

  // I列14行目の各行の値をJ列から引いて結果を表示
  for (var i = 0; i < results.length; i++) {
    let jValue = jValues[i][0];
    if (jValue !== "") {
      let subtractionResult = jValue - results[i][0];
      quotationSheet.getRange(14 + i, 9).setValue(subtractionResult); // I列は9列目
    }
  }
}

function TotalCost(){ //合計（原価）
  let lastRow = quotationSheet.getLastRow();

  // G列とK列の14行目以降の各行の値を取得
  let gValues = quotationSheet.getRange("G14:G" + lastRow).getValues();
  let kValues = quotationSheet.getRange("K14:K" + lastRow).getValues();

  // 計算結果を格納する配列
  let results = [];

  // 各行の計算結果を求めてresults配列に格納
  for (i = 0; i < gValues.length; i++) {
    let result = gValues[i][0] * kValues[i][0];
    results.push([result]);
  }

  // 結果をL列14行目にセット
  quotationSheet.getRange("L14:L" + (14 + results.length - 1)).setValues(results);
}

function CrudeInterestRate() {  //

}
