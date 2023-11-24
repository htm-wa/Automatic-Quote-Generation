  const quotationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("見積作成");
  const quotation = quotationSheet.getSheetName();
  const quotationData = quotationSheet.getDataRange().getValues();
  
  const itemSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("商品管理簿");
  const rowData = itemSheet.getRange(2, 2, itemSheet.getLastRow() -1, itemSheet.getLastColumn() - 1).getValues(); //2列目の１行最後まで取得
  
  const estimateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("見積書");
  
  const tradeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("貿易情報");
  const usDollarLastRow = tradeSheet.getLastRow();
  
  const customerRegistry = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("顧客管理簿");
  // const CATEGORYINFOROW = 4 // 商品情報の配列を5回繰り返す

function onEdit(e){
  const activeSheet = e.source.getActiveSheet();  // 編集が行われたシート
  const sheetName = e.source.getSheetName();  //編集されたシートの名前を取得
  const editedRange = e.range;
  const editedRow = e.range.getRow(); //編集されたセルの行を取得
  const editedValue = e.range.getValue();  //編集されたセルの値
  const editedData = quotationData[editedRow -1];
  const emptyArray = [];  //空配列作成

  if(sheetName === quotation){ //編集されたシートでセルが編集されたら
    if (editedRange.getColumn() >=  1 && editedRange.getRow() >= 14 && editedValue !== undefined) { //編集されたセルのが１列目１０行目以降で編集されたセルが未定義でなかったら
      for (i = 0; i < rowData.length; i++) {
        if(rowData[i][0] === editedValue){ // A列の値と編集されたセルの値が一致するか確認
            emptyArray.push(rowData[i][1]);   //条件が一致したらemptyArray[]の末尾に要素を追加
            
        } else if(rowData[i][1] === editedValue  && rowData[i][2] !== ''){
            emptyArray.push(rowData[i][2]);
        } else if(rowData[i][2] === editedValue && rowData[i][3] !== ''){
            emptyArray.push(rowData[i][3]);
           
        } else if(rowData[i][3] === editedValue && rowData[i][4] !== ''){
            emptyArray.push(rowData[i][4]);
        } 
      }
      let dropdownRange = editedRange.offset(0, 1);
      if(emptyArray.length !== 0){
        dropdownRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(emptyArray).build());
        
      }  
    } 
  }
// }   

// function MatchConfirmation(e){

         const matchedUnitPrice = [];
         const matchedUsDollar = [];
         for(i = 0; i < rowData.length; i++){
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
                break;
              }
          }
          if (matchedUnitPrice.length > 0) {
            console.log(matchedUnitPrice.length, 'matchedUnitPrice.length');
            let numRows = matchedUnitPrice.length;
            let newRowTradeSheet = usDollarLastRow + 1;
              console.log(usDollarLastRow, 'usDollarLastRow');
              // console.log(rowData[i][5], 'rowData[i][5]');
              console.log(matchedUnitPrice, 'matchedUnitPrice2');
              console.log(matchedUsDollar, 'matchedUsDollar2');
              console.log(editedRow, 'editedRow');
              console.log('Match found. Setting value in quotationSheet.');
            quotationSheet.getRange("L" + editedRow).setValue(matchedUnitPrice);
              console.log(matchedUnitPrice[i],'matchedUnitPrice[i]')
            tradeSheet.getRange(newRowTradeSheet, 6, 1, 1).setValue(matchedUnitPrice);
              console.log(matchedUsDollar[i], 'matchedUsDollar[i]');
            tradeSheet.getRange(newRowTradeSheet, 5, 1, 1).setValue(matchedUsDollar);
              
              // break;
            // tradeSheet.getRange(newRowTradeSheet, 6, numRows, 1).setValue(matchedUnitPrice[0]);
            // tradeSheet.getRange(newRowTradeSheet, 5, numRows, 1).setValue(matchedUsDollar[0]);
            }
}
function KeywordSearch(){
  const keyword = quotationSheet.getRange("B2").getValue();
  const suggestionRange = quotationSheet.getRange("B1:K1");
  const targetColumn = 1;
  
  
  
  
  
  
  
  //検索候補クリア
  suggestionRange.clearContent();
}
            


 
  
   


