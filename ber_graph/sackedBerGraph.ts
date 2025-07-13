/*基本給、積立金、通勤手当の内訳を積み立て縦棒グラフにする */

//ToDo:X軸は人の名前にならないのか？、

//定義
function main(workbook: ExcelScript.Workbook) { // 現在開いてるシートを取得
                                //「スプレッドシート全体」を操作するためのオブジェクト
  const sourceSheet = workbook.getActiveWorksheet(); 

  const startInfoCol = 5; //各社員情報の開始列。この基準は変えないこと。

  // ラベルの位置を探す関数(再利用性・保守性のためfrontSeetを引数にする)
  function findCellByLabel(frontSheet: ExcelScript.Worksheet, targetLabel: string): { labelRow: number, labelcol: number } {
    // 使用範囲の取得（表がある、ワークシートで使用されている空でない範囲）
    const frontSheetUsedRange = frontSheet.getUsedRange(); //ローカル
    // ラベルのあるセルの位置を取得
    const labelValues = frontSheetUsedRange.getValues(); // 範囲内のセルの値を2次元配列で取得する。
        for (let findRow = 0; findRow < labelValues.length; findRow++) {
            for (let findCol = 0; findCol < labelValues[0].length; findCol++) {
            // 各セルの値を文字列として取得し、空白を除去。ラベルと一致するかを判定。
            if (String(labelValues[findRow][findCol]).trim() === targetLabel) {
                return { labelRow: findRow, labelcol: findCol };
            }
            }
        }
        //見つからなかったらエラー
        throw new Error(`ラベル "${targetLabel}" が見つかりませんでした`);
        }

    // " "指定したラベルのある行を取得（必要な項目を追加していく）
    const rowName = findCellByLabel(sourceSheet, "氏名").labelRow;
    const rowBasic = findCellByLabel(sourceSheet, "基本給").labelRow;
    const rowSavings = findCellByLabel(sourceSheet, "積立金").labelRow;
    const rowCommute = findCellByLabel(sourceSheet, "通勤手当").labelRow;

    // 行番号のリスト
    const rowIndexes = [rowBasic, rowSavings, rowCommute];

    // 使用範囲の値を取得。シート全体のデータを配列として扱う。
    const usedRange = sourceSheet.getUsedRange(); //空でなく、実際に使われているセル範囲を取得
    const usedValues = usedRange.getValues(); //usedrageで取得した範囲のセルの値を2次元配列として取得（配列に入れて高速化）

    //氏名の入力されている（各従業員情報の）最終列を探す
    let detectedLastCol = startInfoCol; //初期化（5行目～）

    //外側：startInfoCol（5列目）～、infoColが氏名のある列数まで、infoCol＋4列
    for (let infoCol = startInfoCol; infoCol < usedValues[rowName].length; infoCol += 4) {
      let hasName = false; // 氏名がなければfalse

      // 各4列ブロック内をすべてチェック
      for (let mergedColumns = 0; mergedColumns < 4; mergedColumns++) {
        //氏名が入っている行(rowName)の、現在の列（infoCol + mergedColumns）を取得
        const cellName = usedValues[rowName][infoCol + mergedColumns];
          if (cellName !== null && cellName !== undefined && String(cellName).trim() !== "") {
            hasName = true;
            break;
          }
       }
      // 氏名がみつかった最後のブロックの開始列を記録。
      if (hasName) {
        detectedLastCol = infoCol;
      }
    }
    //従業員情報が入った最終列（"小計"も文字列のため、氏名とカウントされてしまう）
    const endInfoCol = detectedLastCol - 4; // 小計の4列を除外

    // 行番号リストの中で一番下の行を見つける（+1で、0行目～その行まで）
    const rowIndexesLastCount = Math.max(...rowIndexes) + 1;
    // 従業員の各情報列の最後の列番号(+1で0行目～その列までカバー)
    const colLastCount = endInfoCol + 1;
    // 必要なデータを全部取り出す。
    // 0行目・0列～rowIndexesLastCount×colLastCountのサイズの二次元配列として取得。
    const allValues = sourceSheet.getRangeByIndexes(0, 0, rowIndexesLastCount, colLastCount).getValues();

    // 従業員名を格納する配列
    let employeeList: string[] = [];
    // 各行の数値データを格納する2次元配列（今回は5行）
    let dataLists: number[][] = [[], [], [], [], []];

    //従業員の情報（4列ごとに1人）に処理するループ
    for (let col = startInfoCol; col <= endInfoCol; col += 4) {
        // 氏名の行にある、各列を取得し、文字列として整形⇒従業員名の取得
        const name = String(allValues[rowName][col]).trim();
        // 基本給の行にある、各列を取得し、数値として整形⇒各基本給の取得
        const basic = Number(allValues[rowBasic][col]);
        //基本給が数値である場合、処理を続行。従業員名を配列に格納。
        if (!isNaN(basic)) {
        employeeList.push(name);

        //各行の数値データを格納
        for (let i = 0; i < rowIndexes.length; i++) {
            dataLists[i].push(Number(allValues[rowIndexes[i]][col]) || 0);
        }
        }
    }

    //新しいシートを作成
    const newSheet = workbook.addWorksheet("積み立て縦グラフ");
    //出力開始セルを設定
    const outputStartRow = 0;
    const outputStartCol = 0;
    //シートの左端（縦方向）に出力
    const categories = ["基本給", "積立金", "通勤手当"];

    //ヘッダー行の出力　employeeLi（従業員リスト）を横方向に並べる
    newSheet.getCell(outputStartRow, outputStartCol).setValue("氏名");
    for (let i = 0; i < employeeList.length; i++) {
      newSheet.getCell(outputStartRow, outputStartCol + 1 + i).setValue(employeeList[i]);
    }

    //各カテゴリを縦方向へ並べる
    for (let row = 0; row < categories.length; row++) {
      newSheet.getCell(outputStartRow + 1 + row, outputStartCol).setValue(categories[row]);
      //2次元配列dataListsに、各カテゴリ行に従業員ごとの金額を横方向に並べていく
      for (let col = 0; col < employeeList.length; col++) {
          newSheet.getCell(outputStartRow + 1 + row, outputStartCol + 1 + col).setValue(dataLists[row][col]);
        }
    }

    //指定した範囲のセルを取得
    const chartRange = newSheet.getRangeByIndexes( //指定した範囲のセルを取得
      outputStartRow, //表の左上のセル指定
      outputStartCol,
      categories.length + 1, //カテゴリの行数に氏名ヘッダー行を加えたもの
      employeeList.length + 1 //従業員リストの列数に氏名列を加えたもの
    );

    // グラフ作成
    const createchart = newSheet.addChart(ExcelScript.ChartType.columnStacked, chartRange);

    createchart.setSeriesByRows(false);

    // グラフの位置を設定
    createchart.setPosition(
    newSheet.getCell(0, employeeList.length + 3),
    newSheet.getCell(20, employeeList.length + 10)
    );

    // 横軸ラベル（従業員名）を設定
    createchart.setXAxisCategoryLabels(
    newSheet.getRangeByIndexes(outputStartRow, outputStartCol + 1, 1, employeeList.length)
    );

    // 系列名（カテゴリ）を設定
    const series = createchart.getSeries();
      for (let i = 0; i < series.length; i++) {
        series[i].setName(categories[i]);
      }
}