  let quotationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("見積作成");
  let itemSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("商品管理簿");
  let targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("見積書");
  let tradeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("貿易情報");

  const CATEGORYINFOROW = 4 // 商品情報の配列を5回繰り返す

function onEdit(e) {
  let activeSheet = e.source.getActiveSheet();  // 編集が行われたシート
  let sheetName = e.source.getSheetName();  //編集されたシートの名前を取得
  let editedRange = e.range;
  let editedRow = e.range.getRow(); //編集されたセルの行を取得
  let editedValue = e.range.getValue();  //編集されたセルの値
  let editedData = e.range.getSheet().getRange(editedRow, 1, 1, 5).getValues()[0];

  let quotation = quotationSheet.getSheetName();
  let rowData = itemSheet.getRange(2, 2, itemSheet.getLastRow() -1, itemSheet.getLastColumn() - 1).getValues(); //2列目の１行最後まで取得
  let usDollarLastRow = tradeSheet.getLastRow();
  let usDollar = tradeSheet.getRange(4, 5, targetSheet.getLastRow(), 2).getValues();
  let emptyArray = [];  //空配列作成
  console.log(usDollar);

  if(sheetName === quotation){ //編集されたシートでセルが編集されたら
    if (editedRange.getColumn() >=  1 && editedRange.getRow() >= 14 && editedValue !== undefined) { //編集されたセルのが１列目１０行目以降で編集されたセルが未定義でなかったら
      for (i = 0; i < rowData.length; i++) {
        if(rowData[i][0] === editedValue){ // A列の値と編集されたセルの値が一致するか確認
            console.log('rowData[i]][1]をpush');
            console.log('Comparing:', rowData[i][1], 'and', editedValue, '[i][1]');
            emptyArray.push(rowData[i][1]);   //条件が一致したらemptyArray[]の末尾に要素を追加
            
        } else{
          console.log('else1回目');
        }
        if(rowData[i][1] === editedValue  && rowData[i][2] !== '' && rowData[i][2] !== []){
          console.log('rowData[i][1]をpush') ;
          console.log('Comparing:', rowData[i][2], 'and', editedValue, '[i][2]');
          emptyArray.push(rowData[i][2]);
        } else {
          console.log('else2回目');
        }
        if(rowData[i][2] === editedValue && rowData[i][3] !== '' && rowData[i][3] !== []){
          console.log('rowData[i][3]をpush');
            console.log('Comparing:', rowData[i][3], 'and', editedValue, '[i][3]');
            emptyArray.push(rowData[i][3]);
           
        } else {
          console.log('else3回目');
        }
         if(rowData[i][3] === editedValue && rowData[i][4] !== '' && rowData[i][4] !== []){
            console.log('rowData[i][4]をpush');
            console.log('Comparing:', rowData[i][4], 'and', editedValue, '[i][4]');
            emptyArray.push(rowData[i][4]);
        } else {
            console.log('else4回目');
        }
      }
      let dropdownRange = editedRange.offset(0, 1);
        if(emptyArray.length !== 0){
          console.log(emptyArray, 'emptyArray');
          dropdownRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(emptyArray).build());
          console.log('ドロップダウン作成');
          editedRange.clearDataValidations();
        } else {
            console.log('ドロップダウン作らない');
           for (let i = 1; i <= rowData.length; i++) {
              let nextCell = editedRange.offset(0, i);
              if (nextCell) {
              // Add an empty string to the next cell
              nextCell.setValue('');
              
              
              } 
            } 
          }
   

// function MatchConfirmation(){
         let matchedUnitPrice = [];
         let matchedUsDollar = [];

          for(i = 0; i < rowData.length; i++){
            console.log('for動いた');
            console.log(rowData, 'rowData2');
            console.log(editedRow, 'editedRow');
            console.log(rowData[i][6],'rowData[i][6]');
            console.log(editedData, 'editedData');
            console.log(rowData[i][5], 'rowData[i][5]');
        
            if(
              rowData[i][0] === editedData[0] && 
              rowData[i][1] === editedData[1] && 
              rowData[i][2] === editedData[2] && 
              rowData[i][3] === editedData[3] && 
              rowData[i][4] === editedData[4]
              ){
                console.log('if動いた');
                matchedUnitPrice.push(rowData[i][6]);
                console.log(matchedUnitPrice, 'matchedUnitPrice');
                matchedUsDollar.push(rowData[i][5]);
                console.log(matchedUsDollar, 'matchedUsDollar');
                
              }
              console.log('else');
          }
          if (matchedUnitPrice.length > 0) {
            let numRows = matchedUnitPrice.length;
            let newRowTradeSheet = usDollarLastRow + 1;
            // console.log(rowData[i][6], 'rowData[i][6]');
            // console.log(rowData[i][5], 'rowData[i][5]');
            console.log(matchedUnitPrice, 'matchedUnitPrice2');
            console.log(matchedUsDollar, 'matchedUsDollar2');
            console.log(editedRow, 'editedRow');
            console.log('Match found. Setting value in quotationSheet.');
            quotationSheet.getRange("K" + editedRow).setValue(matchedUnitPrice);
            for (let i = 0; i < matchedUnitPrice.length; i++) {
              tradeSheet.getRange(newRowTradeSheet + i, 6, 1, 1).setValue(matchedUnitPrice[i]);
              tradeSheet.getRange(newRowTradeSheet + i, 5, 1, 1).setValue(matchedUsDollar[i]);
            // tradeSheet.getRange(newRowTradeSheet, 6, numRows, 1).setValue(matchedUnitPrice[0]);
            // tradeSheet.getRange(newRowTradeSheet, 5, numRows, 1).setValue(matchedUsDollar[0]);
            }
            
          }
        }
            //  editedRange.clearDataValidations(editedRange);
         
      // 
    // "商品管理簿 "のH列から "見積シート "のH列に値を追加
    // let targetRange = quotationSheet.getRange(10, 11,  quotationSheet.getLastRow(), 1); 
     } 
} 

// targetRange.setValues();
  
   


