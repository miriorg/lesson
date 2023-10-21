const strNumberFormat: string = "#,##0,;[赤]-#,##0,";
const strFormulaQ: string = "=SUM(RC[-3]: RC[-1])";   //左3データの合計
const strFormulaH: string = "=RC[-5] + RC[-1]";       //左2Qの合計
const strFormulaTotal: string = "=RC[-10] + RC[-1]";  //左2Hの合計
const strZero = "0";
const arDefault: string[] = [
  strZero, strZero, strZero, strFormulaQ,
  strZero, strZero, strZero, strFormulaQ, strFormulaH,
  strZero, strZero, strZero, strFormulaQ,
  strZero, strZero, strZero, strFormulaQ, strFormulaH, strFormulaTotal
];
let arDefaults: string[][] = new Array();           // 入力欄初期値
let arColClients: string[][] = new Array();         // クライアント名一覧
let arColSalesMonth: string[][] = new Array();      // 
let arColGross: string[][] = new Array();           // 粗利計算式
let arColMargin: string[][] = new Array();          // 粗利率計算式
let arColQ: string[][] = new Array();               // Q列計算式
let arColH: string[][] = new Array();               // H列計算式
let arColTotal: string[][] = new Array();           // Total列計算式
let arLatehiding: string[] = new Array();           // 後で非表示にするシート名

interface objClient {
  strName: String;    //顧客名
  bCond: Boolean;     //集計対象フラグ
}
interface objHeader {
  strTitle: string;     //項目名
  strColor: string;     //背景色
  nWidth: number;       //項目幅
}
interface Subscription {
  customer: String; //顧客名
  service: String;  //サービス名
  plan: String;     //プラン名
  fromDate: Date;   // date format
  toDate: Date;     // date format
  sales: number;    //売上
  cost: number;     //原価
}
interface SubscriptionSales {
  customer: String;
  sales: number;
  cost: number;
  salesStr: string;
  costStr: string;
}

function log(strOutput: string) {
  var timeStamp = getNowWithString();
  console.log(`${timeStamp}:${strOutput}`);
}
function jsDate(dateSerial: number) {
  return new Date(Math.round((dateSerial - 25569) * 86400 * 1000));
}
function getNowDateWithString() {
  var dt = new Date();
  return `${dt.getFullYear()}${("00" + (dt.getMonth() + 1)).slice(-2)}${("00" + dt.getDate()).slice(-2)}`
}
function getNowWithString() {
    var dt = new Date();
  return `${("00" + (dt.getHours() + 1)).slice(-2)}${("00" + dt.getMinutes()).slice(-2)}${("00" + dt.getMilliseconds()).slice(-2)}`
}

//Clientsシートからクライアント名の一覧を取得
//以降はarClients経由で参照する
function getActiveClients(worksheet: ExcelScript.Worksheet): objClient[] {

  let arResults: objClient[] = new Array();
  let sheetClients: ExcelScript.Worksheet = worksheet;
  let rngClient = sheetClients.getRange("B2");
  let rngCond = sheetClients.getRange("C2");

  while (rngClient.getValue()) {
    let bCond = (rngCond.getValue() as Boolean);
    let strName = (rngClient.getValue() as String);
    //log(`Parsing:${bCond} / ${strName}`);
    let contract: objClient = {
      strName: strName as String,
      bCond: bCond as Boolean
    };
    //log(`Pushing:${contract.bCond} / ${contract.strName}`);
    if (bCond) {
      arResults.push(contract);
    }

    //次行へ…
    rngClient = rngClient.getOffsetRange(1, 0);
    rngCond = rngCond.getOffsetRange(1, 0);
  }
  log(`# of Clients: ${arResults.length}`);
  return arResults;
}

//フィルタを操作して顧客別・月別のupsellを集計する
function getUpsellMonth(
  tblUpsell: ExcelScript.Table, 
  strCustomer: String, 
  strYYYYMM: string, 
  sheetEscape: ExcelScript.Worksheet): string[] {

  // Toggle auto filter on table tblUpsell
  tblUpsell.getAutoFilter().apply(tblUpsell.getRange());

  // Apply checked items filter on table tblUpsell column 顧客
  tblUpsell.getColumnByName("顧客").getFilter().applyValuesFilter([`${strCustomer}`]);
  tblUpsell.getColumnByName("ステータス").getFilter().applyValuesFilter(["検収済み"]);
  tblUpsell.getColumnByName("検収年月").getFilter().applyValuesFilter([`${strYYYYMM}`]);

  // フィルターした値を取得
  let rng = tblUpsell.getRange();
  const values = rng.getVisibleView().getValues();

  // 転写先シートに貼り付け
  sheetEscape.getRangeByIndexes(0, 0, values.length, values[0].length).setValues(values);
  // 転写先の「全体売上/全体原価」の範囲
  let rngNew = sheetEscape.getRange(`J2:K${values.length}`);	//J2から始まり、対象は全体売上/全体原価
  rngNew.select;

  //フィルタ結果が0行の場合、列のヘッダが帰ってくるのでそれを確認
  const nReffer = rngNew.getValues();
  if (nReffer[0][0] == "全体売上") {
    return ["0", "0"];
  }
  let strSales = "=";
  let strCost = "=";

  for (const el of nReffer) {
    let strloopSales: string = (el[0] as string);
    let strloopCost: string = (el[1] as string);
    //log(`Read from Sheet for ${strCustomer} on ${strYYYYMM}: ${strSales} / ${strCost}`);

    strSales += `+${strloopSales}`;
    strCost += `+${strloopCost}`;
  }

  //後片付け：待避シートに複製されたフィルタ結果を消去する
  let rngClear = sheetEscape.getRangeByIndexes(0, 0, values.length, values[0].length);
  rngClear.clear(ExcelScript.ClearApplyTo.all);

  return [strSales, strCost];
}

//処理速度向上目的でスクリプト全般で利用する頻度の高い変数の初期設定を実施
//初期設定回数・変数定義処理の回数を減らす事でAPI呼び出しの回数なども削減できる
function initVars(arClients: objClient[]) {
  const nHeight = arClients.length
  //const strFormulaGross = `=IF(R[${(nHeight + 1) * -2}]C,R[${(nHeight + 1) * -1}]C/R[${(nHeight + 1) * -2}]C,"-")`;
  const strFormulaGross = `=R[${(nHeight + 1) * -2}]C-R[${(nHeight + 1) * -1}]C`;
  const strFormulaMargin = `=IF(R[${(nHeight + 1) * -3}]C,R[${(nHeight + 1) * -1}]C/R[${(nHeight + 1) * -3}]C,"-")`;
  let strFormula = `=R[${nHeight * -2}]C-R[${-nHeight}]C`;
  const nWidth = 3 * 4 + 4 + 2 + 1;

  for (const el of arClients) {
    arColGross.push([
      strFormulaGross, strFormulaGross, strFormulaGross, strFormulaGross, strFormulaGross,
      strFormulaGross, strFormulaGross, strFormulaGross, strFormulaGross, strFormulaGross,
      strFormulaGross, strFormulaGross, strFormulaGross, strFormulaGross, strFormulaGross,
      strFormulaGross, strFormulaGross, strFormulaGross, strFormulaGross
    ]);         //粗利計算式クライアント数分
    arColMargin.push([
      strFormulaMargin, strFormulaMargin, strFormulaMargin, strFormulaMargin, strFormulaMargin,
      strFormulaMargin, strFormulaMargin, strFormulaMargin, strFormulaMargin, strFormulaMargin,
      strFormulaMargin, strFormulaMargin, strFormulaMargin, strFormulaMargin, strFormulaMargin,
      strFormulaMargin, strFormulaMargin, strFormulaMargin, strFormulaMargin
    ]);       //粗利率計算式クライアント数分
    arDefaults.push([
      strZero, strZero, strZero, strFormulaQ,
      strZero, strZero, strZero, strFormulaQ, strFormulaH,
      strZero, strZero, strZero, strFormulaQ,
      strZero, strZero, strZero, strFormulaQ, strFormulaH, strFormulaTotal
    ]);     //売上・原価欄の初期値

    arColClients.push([el.strName + ""]);       //クライアント名一覧
  }
}

//Client一覧の出力(セクション毎に1回、1シート4回呼び出される)
function colClients(rng: ExcelScript.Range, arClients: objClient[]) {
  let nHeight = arClients.length;
  rng = rng.getAbsoluteResizedRange(nHeight, 1);
  rng.clear(ExcelScript.ClearApplyTo.formats);
  rng.setValues(arColClients);
}

// Upsellシート生成
function generateUpsell(
  sheet: ExcelScript.Worksheet, 
  table: ExcelScript.Table, 
  arClients: objClient[], 
  sheetEscape: ExcelScript.Worksheet) {
  let nHeight = arClients.length;
  const strFormulaMonth = `=SUM(R[1]C:R[${nHeight}]C)`;
  const strFormulaGross = `=R[${(nHeight + 1) * -2}]C-R[${(nHeight + 1) * -1}]C`; /****/
  let strFormulaRate = `=IF(R[${(nHeight + 1) * -3}]C,R[${(nHeight + 1) * -1}]C/R[${(nHeight + 1) * -3}]C,"-")`;
  let arFormulaSales: string[][] = new Array();
  let arFormulaCosts: string[][] = new Array();
  let nCost: number;
  let nSales: number;
  let strValues: string[] = new Array();
  let tblCosts: string[][] = new Array();
  let tblSales: string[][] = new Array();
  //シートの集計年は「先月における年」とする
  //年度を意識するなら「getMonth() - 4」に変更?
  var dt = new Date();
  dt.setMonth(dt.getMonth() - 1);
  const YYYY = dt.getFullYear();

  log(`correct Sales&Costs.`);
  for (const el of arClients) { //全クライアントに対して
    let arSales: string[] = new Array();        //月のクライアント毎の売上(表の横方向)
    let arCosts: string[] = new Array();        //月のクライアント毎の原価(表の横方向)
    for (const MM of ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]) {
      strValues = getUpsellMonth(table, el.strName, `${YYYY}/${MM}`, sheetEscape);
      arSales.push(strValues[0]);             //一月の売上をスタック
      arCosts.push(strValues[1]);             //一月の原価をスタック
    }
    tblSales.push(arSales);                     //月次売上1年分を1クライアント分としてStack
    tblCosts.push(arCosts);                     //月次売上1年分を1クライアント分Stack
  }
  setupFormula(sheet, tblSales, tblCosts);           //Totalシートは売上も原価も計算式は一緒
  table.clearFilters();
}
//Totalシートの作成
function generateTotal(sheet: ExcelScript.Worksheet, arClients: objClient[]) {
    const strFormula: string = "=Kaetec!RC+Upsell!RC+Subsc!RC";
    let arFormula: string[][] = new Array();
    let arMonth: string[] = [                        ////月次の数値は全てKaetec+Upsell+Subsc
      strFormula, strFormula, strFormula, strFormula, 
      strFormula, strFormula, strFormula, strFormula, 
      strFormula, strFormula, strFormula, strFormula
    ];

    for(const el of arClients){                     //全クライアントに対して
      arFormula.push(arMonth);                    //1クライアント分stack
    }
    setupFormula(sheet, arFormula, arFormula);         //Totalシートは売上も原価も計算式は一緒
}

function initSalesArea(rngPos: ExcelScript.Range, arClients: objClient[]){
  // 全クライアント分のデータ領域
  //高さ:クライアント数 / 幅：12ヶ月&4期&2半年+1トータル
  var rngDef: ExcelScript.Range = (rngPos.getOffsetRange(0, 1)).getAbsoluteResizedRange(arClients.length, 12 + 4 + 2 + 1);
  //log(`initial setup for: ${rngDef.getAddress()} - ${rngDef.getRowCount()} / ${rngDef.getColumnCount()}`);
  //log(`initial setup for: ${arDefaults.length} / ${arDefaults[0].length}`);
  rngDef.clear(ExcelScript.ClearApplyTo.formats);
  rngDef.setFormulasR1C1(arDefaults);
  rngDef.setNumberFormatLocal(strNumberFormat);
}

//シート初期設定
function initSheet(sheet: ExcelScript.Worksheet, arClients: objClient[]) {
  let nHeight = arClients.length;
  let strFormulaM = `=SUM(R[1]C:R[${nHeight}]C)`;
  let strFormulaRate = `=IF(R[${(nHeight + 1) * -3}]C,R[${(nHeight + 1) * -1}]C/R[${(nHeight + 1) * -3}]C,"-")`;

  let arFormula: string[] = [
      strFormulaM, strFormulaM, strFormulaM, strFormulaQ,
      strFormulaM, strFormulaM, strFormulaM, strFormulaQ, strFormulaH,
      strFormulaM, strFormulaM, strFormulaM, strFormulaQ,
      strFormulaM, strFormulaM, strFormulaM, strFormulaQ, strFormulaH, strFormulaTotal
  ]
  let arFormulaRate: string[] = [
      strFormulaRate, strFormulaRate, strFormulaRate, strFormulaRate,
      strFormulaRate, strFormulaRate, strFormulaRate, strFormulaRate, strFormulaRate,
      strFormulaRate, strFormulaRate, strFormulaRate, strFormulaRate,
      strFormulaRate, strFormulaRate, strFormulaRate, strFormulaRate, strFormulaRate, strFormulaRate
  ]

  let formulasRate: string[][] = new Array();
  formulasRate.push(arFormulaRate);
  let formulas: string[][] = new Array();
  formulas.push(arFormula);

  //表のヘッダー行を作成
  let rng = sheet.getRange("A1");
  const strColor = "#A5A5A5";
  let arHeader: objHeader[] = [
      { strTitle: "単位：千円", strColor: "#70AD47", nWidth: 16.5 },
      { strTitle: "1月", strColor: "#70AD47", nWidth: 8.08 },
      { strTitle: "2月", strColor: "#70AD47", nWidth: 8.08 },
      { strTitle: "3月", strColor: "#70AD47", nWidth: 8.08 },
      { strTitle: "1Q", strColor: "#5B9BD5", nWidth: 8.08 },
      { strTitle: "4月", strColor: "#70AD47", nWidth: 8.08 },
      { strTitle: "5月", strColor: "#70AD47", nWidth: 8.08 },
      { strTitle: "6月", strColor: "#70AD47", nWidth: 8.08 },
      { strTitle: "2Q", strColor: "#5B9BD5", nWidth: 8.08 },
      { strTitle: "1H", strColor: "#5B9BD5", nWidth: 8.08 },
      { strTitle: "7月", strColor: "#70AD47", nWidth: 8.08 },
      { strTitle: "8月", strColor: "#70AD47", nWidth: 8.08 },
      { strTitle: "9月", strColor: "#70AD47", nWidth: 8.08 },
      { strTitle: "3Q", strColor: "#5B9BD5", nWidth: 8.08 },
      { strTitle: "10月", strColor: "#70AD47", nWidth: 8.08 },
      { strTitle: "11月", strColor: "#70AD47", nWidth: 8.08 },
      { strTitle: "12月", strColor: "#70AD47", nWidth: 8.08 },
      { strTitle: "4Q", strColor: "#5B9BD5", nWidth: 8.08 },
      { strTitle: "2H", strColor: "#5B9BD5", nWidth: 8.08 },
      { strTitle: "Total", strColor: "#FFC000", nWidth: 8.08 }
  ];

  //arHeaderの定義を使って上部ヘッダの作成
  sheet.activate;
  for (const el of arHeader) {
    rng.setValue(el.strTitle);
    rng.getFormat().getFill().setColor(el.strColor);
    rng.getFormat().getFont().setColor("#FFFFFF");
    rng.getFormat().getFont().setBold(true);
    rng.getFormat().setColumnWidth(6.45 * el.nWidth);   // elem[2] * 6.2539
    rng = rng.getOffsetRange(0, 1); // ひとつ右へ
  }
  rng = sheet.getRange("A2");
  rng.clear(ExcelScript.ClearApplyTo.formats);
  rng.setValue("売上");
  rng.getFormat().getFill().setColor(strColor);
  rng.getFormat().getFont().setBold(true);

  let rngHeader = sheet.getRangeByIndexes(1 + (nHeight + 1) * 0, 1, 1, 12 + 4 + 2 + 1);
  rngHeader.getFormat().getFill().setColor(strColor);
  rngHeader.getFormat().getFont().setBold(true);
  rngHeader.setFormulasR1C1(formulas);
  (rngHeader.getResizedRange(nHeight, 0)).setNumberFormatLocal(strNumberFormat);

  rng = rng.getOffsetRange(1, 0); // ひとつ下へ
  colClients(rng, arClients);     //クライアント名
  initSalesArea(rng, arClients);  //初期化
  
  rng = rng.getOffsetRange(nHeight, 0); // クライアント数分下へ
  rng.setValue("原価");
  rng.getFormat().getFill().setColor(strColor);
  rng.getFormat().getFont().setBold(true);

  rngHeader = sheet.getRangeByIndexes(1 + (nHeight + 1) * 1, 1, 1, 12 + 4 + 2 + 1);
  rngHeader.getFormat().getFill().setColor(strColor);
  rngHeader.getFormat().getFont().setBold(true);
  rngHeader.setFormulasR1C1(formulas);
  (rngHeader.getResizedRange(nHeight, 0)).setNumberFormatLocal(strNumberFormat);

  rng = rng.getOffsetRange(1, 0); // ひとつ下へ
  colClients(rng, arClients);     //クライアント名
  initSalesArea(rng, arClients);  //初期化

  rng = rng.getOffsetRange(arClients.length, 0); // クライアント数分下へ
  rng.setValue("粗利");
  rng.getFormat().getFill().setColor(strColor);
  rng.getFormat().getFont().setBold(true);

  rngHeader = sheet.getRangeByIndexes(1 + (nHeight + 1) * 2, 1, 1, 12 + 4 + 2 + 1);
  rngHeader.getFormat().getFill().setColor(strColor);
  rngHeader.getFormat().getFont().setBold(true);
  rngHeader.setFormulasR1C1(formulas);
  (rngHeader.getResizedRange(nHeight, 0)).setNumberFormatLocal(strNumberFormat);

  rngHeader = rngHeader.getOffsetRange(1, 0);             // ひとつ下へ
  rngHeader = rngHeader.getResizedRange(nHeight-1, 0);    //クライアント数分範囲を拡大
  rngHeader.setNumberFormatLocal(strNumberFormat);
  //log(`RANGE:${rngHeader.getRowCount()}/${rngHeader.getColumnCount()}:${arColGross.length}/${arColGross[0].length}`)
  rngHeader.setFormulasR1C1(arColGross);

  rng = rng.getOffsetRange(1, 0); // ひとつ下へ
  colClients(rng, arClients);

  rng = rng.getOffsetRange(nHeight, 0); // クライアント数分下へ
  rng.setValue("粗利率");
  rng.getFormat().getFill().setColor(strColor);
  rng.getFormat().getFont().setBold(true);

  //ヘッダ行 一行分の領域
  rngHeader = sheet.getRangeByIndexes(1 + (nHeight + 1) * 3, 1, 1, 12 + 4 + 2 + 1);
  rngHeader.getFormat().getFill().setColor(strColor);
  rngHeader.getFormat().getFont().setBold(true);
  rngHeader.setPredefinedCellStyle("Percent");
  rngHeader.setFormulasR1C1(formulasRate);

  rngHeader = rngHeader.getOffsetRange(1, 0); // ひとつ下へ
  //クライアント行数分拡げる
  rngHeader = rngHeader.getResizedRange(nHeight-1, 0);

  log(`RANGE:${rngHeader.getRowCount()}/${rngHeader.getColumnCount()} : ${arColMargin.length}/${arColMargin[0].length}`)

  rngHeader.setNumberFormatLocal(strNumberFormat);
  rngHeader.setFormulasR1C1(arColMargin);
  rngHeader.setPredefinedCellStyle("Percent");

  rng.getFormat().getFill().setColor(strColor);
  rng = rng.getOffsetRange(1, 0); // ひとつ下へ
  colClients(rng, arClients);
}

//既存シートをリネームしてバックアップとする
function sheetRename(wb: ExcelScript.Workbook, strSheet: string): ExcelScript.Worksheet {
  let dtNow = getNowDateWithString();
  let sheet = wb.getWorksheet(strSheet);
  let rename: string;
  let counter = 0;

  if (sheet == void 0) { //生成シートが存在しない場合はundefinedが返ってくるので、空シートを追加
    //生成シートが存在しない場合(Backup処理は不要)
    sheet = wb.addWorksheet(strSheet);  //空シートを生成
  } else {
    //生成シートが存在する場合(空シートを作る必要は無い)
    //backupのシートを探す
    let sheetBackup:ExcelScript.Worksheet = void 0;
    do {
      //シートが取得出来なくなるまで繰り返す(=使えるシート名の取得)
      rename = `${strSheet}${dtNow}${("000" + counter.toString()).match(/....$/)}`
      sheetBackup = wb.getWorksheet(rename); //backupシートを取得
      counter++;
    }while(sheetBackup != void 0);
    
    //既存シートをrenameして空シートを追加
    sheet.setName(rename); //既存シートをrenameしてバックアップ
    sheet = wb.addWorksheet(strSheet);  //空シートを生成
    arLatehiding.push(rename); //隠すシート名をメモ(転記できなくなるので先に隠しちゃダメ)
  }
  sheet.setPosition(0);
  return sheet;
}

//Subscriptionテーブルの中身を全件読み込んで返す(抽出は変数から行う)
function parseSubscription(sheet: ExcelScript.Worksheet, arClients: objClient[]): { [key: string]: Subscription[] } {
  let subsc: { [key: string]: Subscription[] } = {};

  sheet.getAutoFilter().remove();
  // Toggle auto filter on selectedSheet
  sheet.getAutoFilter().apply(sheet.getRange("A1"));
  // Clear auto filter on selectedSheet
  sheet.getAutoFilter().clearCriteria();
  // Apply values filter on selectedSheet

  for (const client of arClients) {            //顧客企業ごとに処理
    // Apply values filter on selectedSheet
    sheet.getAutoFilter().apply(sheet.getAutoFilter().getRange(), 0, { filterOn: ExcelScript.FilterOn.values, values: [client.strName as string] });   // 0:A列

    let rngUsed = sheet.getRange().getUsedRange(true);
    let rngView = rngUsed.getVisibleView();
    let values = rngView.getValues();

    if (rngView.getRowCount() == 0) {  //フィルタの結果が空の場合、登録をスキップ
      continue;
    }

    let rows: ExcelScript.RangeView[] = rngView.getRows();  //rngViewで指定されたフィルタ後に表示されている行へのRange配列
    let index: number = 0;
    let arSubsc: Subscription[] = new Array;

    for (const el of values) {          // フィルタされたサブスクリプション行単位
      if (el[0] != "顧客") {
        let dateTo: Date = undefined;
        if (el[4] != "") {
          dateTo = new Date(Math.round(((el[4] as number) - 25569) * 86400 * 1000));  //ExcelDateSerialからJSDateに変換する公式
        }
        let contract: Subscription = {
          customer: client.strName,                                                   // customer: 
          service: el[1] as String,                                                   // service: 
          plan: el[2] as String,                                                      // plan: 
          fromDate: new Date(Math.round(((el[3] as number) - 25569) * 86400 * 1000)), // fromDate: 
          toDate: dateTo,                                                             // toDate: 
          sales: el[6] as number,                                                     // sales: 
          cost: el[7] as number
        };
        log(`el[3]:${typeof el[3]},${el[3] == undefined},${el[3] == null} : ${el[3].toString.length} :  ${el[3]}`);
        if (typeof el[3] == 'string') {
          log(`adjust fromDate`);
          if (el[3].toString.length == 0) {
            contract.fromDate = undefined;
          }
        }
        if (typeof el[6] == 'string') {
          contract.sales = parseFloat(el[6]);
          log(`el[6]:[${el[6]}]`)
          if (Number.isNaN(contract.sales)) {
            contract.sales = 0;
          }
        }
        if (typeof el[7] == 'string') {
          contract.cost = parseFloat(el[7]);
          log(`el[7]:[${el[7]}]`)
          if (Number.isNaN(contract.cost)){
            contract.cost = 0;
          }
        }
        log(`PARSE:${client.strName}.${contract.service}(${contract.plan}) / From:${contract.fromDate} / To:${contract.toDate} / Sales:${contract.sales}(${contract.sales==null}) / Cost:${contract.cost}(${(contract.cost==null)})`);
        log(`PARSE:${typeof el[6]}(${el[6] == undefined},${el[6] == null}) / ${typeof el[7]}(${el[7] == undefined},${el[7] == null})`)
          arSubsc.push(contract);
        //log(`PUSH:${ contract.customer } - ${ contract.service } ${ contract.plan } : ${ fDate(contract.fromDate)} <-> ${fDate(contract.toDate)} : ${contract.sales}/${contract.cost} <${el[4]}>`);
      }
    }
    subsc[client.strName + ""] = arSubsc;
  }

  sheet.getAutoFilter().clearCriteria();
  return subsc;
}

//format用の関数を準備
function fDate(d: Date): string {
  const options = { year: 'numeric', month: 'numeric', day: 'numeric' };
  if (d == undefined) {
    return "[undefined]";
  }
  return d.toLocaleDateString('ja-JP', options);
}

// 指定年月・クライアントのサブスク額を算出する
function calcSubscription(ar: { [key: string]: Subscription[] }, client: string, year: number, month: number): SubscriptionSales {
    let subsc: number = 0;

    let subs: Subscription[];
    let sales: SubscriptionSales[];
    let ss: SubscriptionSales = {
      customer: client,
      sales: 0,
      cost: 0,
      salesStr: "",
      costStr: ""
    }

    //log(`CALCURATE:${client}(${year}/${month})`);

    //指定クライアントのサブスクデータがあるか
    if (client in ar) {  //全てのサブスクデータから指定クライアントの物がある場合のみ
      subs = ar[client];  //指定クライアントのサブスクデータを取り出し
      let dt = new Date(year, month, 1, 0, 0, 0); //指定年月の月初のDate型オブジェクトをを作成
      for (const el of subs) {    // サブスクデータ全てに対して
        var newCond:boolean=true;
        if (newCond){
          if (el.sales != undefined && el.cost != undefined){   //売上と原価が入力されていること
            if (((el.toDate == undefined) && ((el.fromDate == undefined) || (el.fromDate < dt))) || 
              ((el.toDate > dt) && ((el.fromDate == undefined) || (el.fromDate < dt)))){
//            log(`NEWCOND:${el.customer} : ${fDate(el.fromDate)} < ${fDate(dt)} < ${fDate(el.toDate)} ${el.customer} (${year}/${month + 1}) add:${ss.sales}/${ss.cost}=>${ss.sales + el.sales}/${ss.cost + el.cost}`);
              ss.sales += el.sales; //積算
              ss.cost += el.cost;   //積算
              ss.salesStr += `+${el.sales}`;
              ss.costStr += `+${el.cost}`;
//            log(`NEWCOND:--- ${ss.salesStr} / ${ss.costStr}`);
            }
          }
        }
        else{
          log(`${el.customer}-${el.service} : ${fDate(el.fromDate)} < ${fDate(dt)} < ${fDate(el.toDate)}:${el.sales}/${el.cost}`);
          if (                                                  
            (el.fromDate <= dt) && (                        //開始日が fromDate 以降である事
            (el.toDate == undefined) || (dt < el.toDate)   //終了日が undefined もしくは toDate以前である事
            )
          ) { // year/monthの月が契約期間に含まれるかチェック
            log(`${fDate(el.fromDate)} < ${fDate(dt)} < ${fDate(el.toDate)} ${el.customer} (${year}/${month + 1}) add:${ss.sales}/${ss.cost}=>${ss.sales + el.sales}/${ss.cost + el.cost}`);
            ss.sales += el.sales; //積算
            ss.cost += el.cost;   //積算
            ss.salesStr += `+${el.sales}`;
            ss.costStr += `+${el.cost}`;
          }
          else {
            //log(`${fDate(el.fromDate)} < ${fDate(dt)} < ${fDate(el.toDate)} ${el.customer} (${year}/${month + 1}) nop:${ss.sales}/${ss.cost}`);
          }
        }
      }
    }
    else {
      // nothing to do.
    }
    if (ss.salesStr.length == 0) { ss.salesStr = "0" }
    if (ss.costStr.length == 0) { ss.costStr = "0" }
    //log(`RETURN:${ss.customer}(${year}/${month}) ${ss.salesStr}/${ss.costStr}`);
    return ss;
}

//[Subsc]シート生成部
function generateSubsc(sheetSubscNew: ExcelScript.Worksheet, sheetSubscription: ExcelScript.Worksheet, year: number, arClients: objClient[]) {
    let row: number = 0;
    const nHeight: number = arClients.length;
    let strFormulaGross = `=R[${(nHeight + 1) * -2}]C-R[${(nHeight + 1) * -1}]C`;
    let strFormulaRate = `=IF(R[${(nHeight + 1) * -3}]C,R[${(nHeight + 1) * -1}]C/R[${(nHeight + 1) * -3}]C,"-")`;
    //全クライアント・全サブスクリプションを読み込む
    const subsc: { [key: string]: Subscription[] } = parseSubscription(sheetSubscription, arClients);

    let arStringSales: string[][] = new Array();
    let arStringCosts: string[][] = new Array();

    //集計データ作成(各クライアント12ヶ月分)
    for (const client of arClients) {                   //全クライアント分
      let rowSales: string[] = new Array;             //売上
      let rowCosts: string[] = new Array;             //原価
      let ss: SubscriptionSales;
      for (let loopM = 0; loopM < 12; loopM++) {      //12ヶ月分集計、スタック
        ss = calcSubscription(subsc, client.strName + "", year, loopM); //当該年月のサブスク費用を計算式文字列で取得
        if (ss.salesStr.length > 0){
          rowSales.push(`=${ss.salesStr}`);
        }else{
          rowSales.push("0");
        }
        if (ss.costStr.length > 0) {
          rowCosts.push(`=${ss.costStr}`);
        } else {
          rowCosts.push("0");
        }
      }
      arStringSales.push(rowSales);
      arStringCosts.push(rowCosts);
      if  (false){
        //cellデータ作成
        let arStringSales: string[] = new Array();
        let arStringCosts: string[] = new Array();

        for (let loopH = 0; loopH < 2; loopH++) {
          for (let loopQ = 0; loopQ < 2; loopQ++) {
            for (let loopM = 0; loopM < 3; loopM++) {
              let sales:string = ss[loopM + ((loopQ * 3) + (loopH * 6))].salesStr;
              let cost:string = ss[loopM + ((loopQ * 3) + (loopH * 6))].costStr;
              if (sales.length == 0) {        //対象となる売上がないと空文字
                arStringSales.push("0");    //「0」を入れておく
              } else {
                arStringSales.push(`=${sales}`);
              }
              if (cost.length == 0) {        //対象となる売上がないと空文字
                arStringCosts.push("0");    //「0」を入れておく
              } else {
                arStringCosts.push(`=${cost}`);
              }
            }
            arStringSales.push(strFormulaQ);    //四半期計算式(売上)
            arStringCosts.push(strFormulaQ);    //四半期計算式(原価)
          }
          arStringSales.push(strFormulaH);        //半期集計式
          arStringCosts.push(strFormulaH);        //半期集計式
        }
        { //Total
            arStringSales.push(strFormulaTotal);    //1クライアント分(売上)
            arStringCosts.push(strFormulaTotal);    //1クライアント分(原価)
        }
        let ssSales: string[][] = [arStringSales];
        let ssCosts: string[][] = [arStringCosts];

        let rngSales = sheetSubscNew.getRangeByIndexes(row + 2 + (nHeight + 1) * 0, 1, 1, arStringSales.length);
        let rngCosts = sheetSubscNew.getRangeByIndexes(row + 2 + (nHeight + 1) * 1, 1, 1, arStringCosts.length);

        rngSales.setFormulasR1C1(ssSales);
        rngSales.setNumberFormatLocal(strNumberFormat);
        rngCosts.setFormulasR1C1(ssCosts);
        rngCosts.setNumberFormatLocal(strNumberFormat);

        row++;
      }
    }
    if (false){
      //データのdebug用ダンプ
      for (const eachRow of arStringSales){
        var index: number = 0;
        let line: string = "";
        for (const eachCol of eachRow) {
          line += (eachCol + " / ");
        }
        log(`${index++}行目:${line}`);
      }
    }
  setupFormula(sheetSubscNew, arStringSales, arStringCosts);
}

function setupFormula(sheet:ExcelScript.Worksheet, tableSales:string[][], tableCosts:string[][]){
  const nHeight: number = tableSales.length;
  let rowIndex: number[] = [1, 5, 10, 14];
  let formulaSales: string[][] = new Array();
  let formulaCosts: string[][] = new Array();
  if (false) {
  } else {
    //12ヶ月分のデータに、Q集計・H集計・Total集計の計算式を挿入する
    for (let loopClients = 0; loopClients < tableSales.length; loopClients++) {
      let rowSales: string[] = new Array();
      let rowCosts: string[] = new Array();
      for (let loopH = 0; loopH < 2; loopH++) {           //H分の周回(2)
        for (let loopQ = 0; loopQ < 2; loopQ++) {       //Q分の周回(4/2(H)=2)
          for (let loopM = 0; loopM < 3; loopM++) {   //M分の周回(1Q=3Month)
            rowSales.push(tableSales[loopClients][loopH * 2 + loopQ * 3 + loopM]);
            rowCosts.push(tableCosts[loopClients][loopH * 2 + loopQ * 3 + loopM]);
          }
          //Q集計の計算式を挿入
          rowSales.push(strFormulaQ);
          rowCosts.push(strFormulaQ);
        }
        //H集計の計算式を挿入
        rowSales.push(strFormulaH);
        rowCosts.push(strFormulaH);
      }
      //Total集計の計算式を挿入
      rowSales.push(strFormulaTotal);
      rowCosts.push(strFormulaTotal);

      //1クライアント分(1行分)スタック
      formulaSales.push(rowSales);
      formulaCosts.push(rowCosts);
    }
    let rangeSales: ExcelScript.Range = sheet.getRangeByIndexes(2 + (nHeight + 1) * 0, 1, tableSales.length, 19 /* 12Month+4Q+2H+1Total */ );
    let rangeCosts: ExcelScript.Range = sheet.getRangeByIndexes(2 + (nHeight + 1) * 1, 1, tableSales.length, 19);
    log(`RANGE:${rangeSales.getRowCount()}/${rangeSales.getColumnCount()} : VAR:${formulaSales.length}/${formulaSales[0].length}`)
    //一括で計算式を貼り付ける(書式は設定済)
    log(`(${rangeSales.getLeft()},${rangeSales.getTop()},${rangeSales.getColumnCount()},${rangeSales.getRowCount()}) <= (${formulaSales[0].length}, ${formulaSales.length})`);

    rangeSales.setFormulasR1C1(formulaSales);
    rangeCosts.setFormulasR1C1(formulaCosts);
  }
}

//一番新しいハズの[kaetecYYYYMMDD####]シートを捜す
function detectSource(sheets: ExcelScript.Worksheet[]):ExcelScript.Worksheet{
  const yymmdd =                        /20[234][0-9][01][0-9][0123][0-9][0-9][0-9][0-9][0-9]/
  const regex = /[Kk][Aa][Ee][Tt][Ee][Cc]20[234][0-9][01][0-9][0123][0-9][0-9][0-9][0-9][0-9]/
  let arTarget: string[] = new Array();
  let strSheet: string;
  let sheetTarget: ExcelScript.Worksheet = void 0;
  let maxDate: string ="";

  for (const each of sheets) {
    strSheet = each.getName();

    if (null != strSheet.match(regex)) {  //検索結果は1つの検索方法なのでシート名さえ確定すれば良い
      // log(`MATCH:${strSheet}`);
      let arMatch:RegExpMatchArray;
      arMatch = strSheet.match(yymmdd);   //配列で
      // log(`YYMMDD:${arMatch[0]}`);
      if (maxDate < arMatch[0].toString()){
        maxDate = arMatch[0].toString();
        sheetTarget = each;
        log(`YYMMDD: stack it.`);
      }
    }
  }
  return sheetTarget;
}

function dumpRange(rng: ExcelScript.Range){
  let val: string[][] = rng.getValues();

  for (var row = 0; row < val.length; row++){
    var line = "";
    for (var col = 0; col < val[row].length; col++){
      line += ` - ${val[row][col]}`;
    }
    console.log(line);
  }
}

//過去実績を複写
function copyAchiev(
  rngDst: ExcelScript.Range, 
  sheetSrc: ExcelScript.Worksheet, 
  rngSrc: ExcelScript.Range, 
  arClients: objClient[]
){
  for (const el of arClients) {
    let rngFind: ExcelScript.Range = rngSrc.find(el.strName, { matchCase: false });
    if (rngFind != null){
      log(`found: [${el.strName}] at ${rngFind.getAddress()}`);

      //見付かった場合
      rngFind = (rngFind.getOffsetRange(0, 1)).getAbsoluteResizedRange(1, 1 + 12 + 4 + 2);   //発見された位置からcopy元範囲を作成

      log(`copy ${rngFind.getAddress()} -> ${rngDst.getAddress()}`);

      rngDst.copyFrom(rngFind, ExcelScript.RangeCopyType.all, false, false);         //範囲コピー
      rngDst = rngDst.getOffsetRange(1, 0);
    }
    else{
      log(`not found: ${el.strName}`);
    }
  }
}


//スクリプト基点
function main(wb: ExcelScript.Workbook) {
  //処理に関連するシートの変数定義
  let sheetClients: ExcelScript.Worksheet = wb.getWorksheet("Clients");
  let sheetSubscription: ExcelScript.Worksheet = wb.getWorksheet("Subscription");
  let sheetSubscNew: ExcelScript.Worksheet = sheetRename(wb, "Subsc");
  let sheetKaetecNew: ExcelScript.Worksheet = sheetRename(wb, "Kaetec");
  let sheetUpsellNew: ExcelScript.Worksheet = sheetRename(wb, "Upsell");
  let sheetTotalNew: ExcelScript.Worksheet = sheetRename(wb, "Total");

  let sheets: ExcelScript.Worksheet[] = wb.getWorksheets();
  let sheetSource:ExcelScript.Worksheet = detectSource(sheets);

  //フィルタ結果のコピー先用のワークシートの準備
  let sheetEscape: ExcelScript.Worksheet = wb.getWorksheet("一時集計用");
  if (sheetEscape == void 0) { //日付入りシートが存在しない場合はundefinedが返ってくるので、空シートを追加
    sheetEscape = wb.addWorksheet("一時集計用");
    sheetEscape.setVisibility(ExcelScript.SheetVisibility.hidden);
    //  ExcelScript.SheetVisibility.visible
  }

  // [Clients]シートから顧客データ取得
  let clientsActive: objClient[] = getActiveClients(sheetClients);
  log(`Load ${clientsActive.length} Clients.`);

  //計算式・書式・クライアント名など再利用頻度の高い変数の定義処理
  initVars(clientsActive);

    //このExcelBookの各表は全て同じ表の形式になっており、原価・原価率も各シート間で
    //全て同一なので、initShetsでシートの初期設定(表の作成・計算式の設定)を行ってから、
    //generate***関数で売上・原価のデータを構築してsetupFormulaに渡して売上表・原価表を
    //完成させる

  log("Initialize Sheets"); {
    log("Initialize [Kaetec] Sheets.");
    initSheet(sheetKaetecNew, clientsActive);
    log("Initialize [Upsell] Sheets.");
    initSheet(sheetUpsellNew, clientsActive);
    log("Initialize [Total] Sheets.");
    initSheet(sheetTotalNew, clientsActive);
    log("Initialize [Subsc] Sheets.");
    initSheet(sheetSubscNew, clientsActive);
  }
  log("Generate Sheets."); {
    log("Generate [Kaetec] Sheets."); {         /* Kaetecシート生成 */
      if (sheetSource == undefined){
        // 転記対象となるシートが見付からない
        log("[Kaetec] sheet not found.");
      }
      else{
        log(`[Kaetec] copy source sheet:${sheetSource.getName()}.`);
        //売上部分の転記
        // begin place B3
        let rngSrc: ExcelScript.Range;
        let rngDst: ExcelScript.Range = sheetKaetecNew.getRange("B3:S3");  //コピー先初期値

        let rngClients: ExcelScript.Range = sheetSource.getRange("A3");
        let rngUsed: ExcelScript.Range = sheetSource.getUsedRange();
        rngClients = rngClients.getAbsoluteResizedRange(rngUsed.getRowCount() - rngClients.getRowIndex(), 1);
        log(`Search Area:${rngClients.getAddress()}`);
        log(`Used Area:${rngUsed.getAddress()}`);
        log(`Destination Area: ${rngDst.getAddress()}`);

        copyAchiev(rngDst, sheetSource, rngClients, clientsActive);

        //原価部分の転記
        let rngFind: ExcelScript.Range = rngClients.find("原価", { matchCase: false });
        //A列に「原価」という文字列があったらその次の行のセル空がコピー対象と決め打ち
        if (rngFind != null) {
          rngFind.getRowIndex()
          rngFind = rngFind.getAbsoluteResizedRange(rngUsed.getRowCount() - rngFind.getRowIndex(), 1);
          log(`Search Area:${rngFind.getAddress()}`);
          rngDst = (sheetKaetecNew.getRange("B3:S3")).getOffsetRange(clientsActive.length + 1, 0);  //コピー先初期値
          log(`Destination Area: ${rngDst.getAddress()}`);

          copyAchiev(rngDst, sheetSource, rngFind, clientsActive);
        }
        else{
          log(`「原価」が見付かりません`);
        }
      }
    }
    if (true){
      log("Generate [Total] Sheets."); {          /* Totalシート生成 */
        generateTotal(sheetTotalNew, clientsActive);
      }
      log("Generate [Upsell] Sheet."); {          /* Upsellシート生成 */
        let tables: ExcelScript.Table[] = wb.getTables();
        if (tables.length > 0){
          generateUpsell(sheetUpsellNew, tables[0], clientsActive, sheetEscape);
        }
      }
      log("Generate [Subsc] Sheet."); {           /* Subscシート生成 */
        generateSubsc(sheetSubscNew, sheetSubscription, 2023, clientsActive);
      }
    }
  }

  //バックアップ用シートは非表示に
  for (const fn of arLatehiding) {
    (wb.getWorksheet(fn)).setVisibility(ExcelScript.SheetVisibility.hidden);
  }

  log("epilogue.");
  sheetEscape.delete(); //一時集計用シートの削除
}
