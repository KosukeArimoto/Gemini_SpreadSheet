


// // --- 設定項目 ---
// const SLIDES_TEMPLATE_ID_TR = '1NYkmHwG4hHm8sadB_n15N6knXNGXtX3ZpLibePXfKS8'; // ★要変更: Googleスライドに変換したテンプレートのID
// const TEMPLATE_SLIDE_INDEX_TR = 1; // 複製するテンプレートスライドのインデックス (0が1枚目, 1が2枚目)

// // --- 代替テキスト(タイトル)とシート列のマッピング ---
// // キー: スライドテンプレート上の要素に設定した『代替テキストのタイトル』
// // 値: スプレッドシートの列インデックス (A列=0, B列=1, ...)
// // *** テンプレートに設定した代替テキストのタイトルに合わせてキーを正確に入力してください ***
// const ALT_TEXT_TITLE_MAP_TR = {
//   "placeholder_line":1,
//   "placeholder_process":2,
//   "placeholder_title":3,
//   "placeholder_point":4,
//   "placeholder_detail":5,
//   "placeholder_check":6,
// };

// // --- イラスト画像の代替テキストタイトル ---
// const IMAGE_ALT_TEXT_TITLE_TR = 'placeholder_image'; // ★要変更: イラスト画像に設定した代替テキストのタイトル
// const ILLUSTRATION_COLUMN_INDEX_TR = 8; // イラストが入っている列のインデックス (L列=11)

// // --- メイン関数 ---
// function createKnowledgeSlidesTR() {
//   const targetSheetName = tokaiPromptSheet.getRange("C12").getValue();
//   if (!targetSheetName) {
//     Logger.log(`エラー: promptシートのC13セルに対象シート名が入力されていません。`);
//     ui.alert(`エラー: promptシートのC13セルに対象シート名が入力されていません。`);
//     return;
//   }

//   const sheet = ss.getSheetByName(targetSheetName);
//   if (!sheet) {
//     Logger.log(`エラー: データシート "${targetSheetName}" が見つかりません。`);
//     ui.alert(`エラー: データシート "${targetSheetName}" が見つかりません。`);
//     return;
//   }

//   let slidesTemplate;
//   try {
//     slidesTemplate = SlidesApp.openById(SLIDES_TEMPLATE_ID_TR);
//   } catch (e) {
//     Logger.log(`エラー: スライドテンプレートが開けません。ID: ${SLIDES_TEMPLATE_ID_TR} - ${e}`);
//     ui.alert(`エラー: スライドテンプレートが開けません。IDが正しいか確認してください。`);
//     return;
//   }

//   const templateSlide = slidesTemplate.getSlides()[TEMPLATE_SLIDE_INDEX_TR];
//   if (!templateSlide) {
//     Logger.log(`エラー: テンプレートスライド (インデックス ${TEMPLATE_SLIDE_INDEX_TR}) が見つかりません。`);
//     ui.alert(`エラー: テンプレートスライド (インデックス ${TEMPLATE_SLIDE_INDEX_TR}) が見つかりません。`);
//     return;
//   }

//   const data = sheet.getDataRange().getValues();
//   const header = data.shift();

//   if (data.length === 0) {
//     ui.alert('シートにデータが見つかりません（ヘッダーを除く）。');
//     return;
//   }

//   // --- ★★★ ここからが修正点：保存先フォルダの指定 ★★★ ---
//   // promptシートslideSaveDir_posセルからフォルダURLを取得
//   const outputFolderUrl = tokaiPromptSheet.getRange("C13").getValue();
//   let outputFolder = null; // デフォルトはマイドライブのルート

//   if (outputFolderUrl) {
//     const folderId = _extractFolderIdFromUrl(outputFolderUrl);
//     if (folderId) {
//       try {
//         outputFolder = DriveApp.getFolderById(folderId);
//         Logger.log(`保存先フォルダを指定しました: ${outputFolder.getName()} (ID: ${folderId})`);
//       } catch (e) {
//         Logger.log(`警告: 指定されたフォルダURL(ID: ${folderId})が見つからないかアクセスできません。マイドライブのルートに保存します。エラー: ${e}`);
//         ui.alert('警告', `指定された保存先フォルダが見つからないかアクセスできません。\nマイドライブのルートに保存します。`, ui.ButtonSet.OK);
//       }
//     } else {
//       Logger.log(`警告: promptシート${slideSaveDir_pos}セルのURLから有効なフォルダIDを取得できませんでした。マイドライブのルートに保存します。URL: ${outputFolderUrl}`);
//       ui.alert('警告', `promptシート${slideSaveDir_pos}セルのURLが正しくありません。\nマイドライブのルートに保存します。`, ui.ButtonSet.OK);
//     }
//   } else {
//     Logger.log("保存先フォルダの指定がないため、マイドライブのルートに保存します。");
//   }

//   // --- プレゼンテーションの作成 ---
//   const newPresentationTitle = `東海理科保全詳細_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd')}`;
//   // SlidesApp.createだけではファイル実体はまだDrive上に確定しない場合があるため、一度IDを取得する
//   const tempPresentation = SlidesApp.create(newPresentationTitle);
//   const presentationId = tempPresentation.getId();
//   const presentationName = tempPresentation.getName();
//   SlidesApp.openById(presentationId); // 一度開いておく（不要かもしれないが一応）
//   Logger.log(`新規プレゼンテーションを作成しました: ${presentationName} (ID: ${presentationId})`);

//   // --- Drive上のファイルを取得し、指定フォルダに移動 ---
//   const presentationFile = DriveApp.getFileById(presentationId);
//   if (outputFolder) {
//     try {
//       presentationFile.moveTo(outputFolder);
//       Logger.log(`プレゼンテーションをフォルダ「${outputFolder.getName()}」に移動しました。`);
//     } catch (moveError) {
//        Logger.log(`警告: プレゼンテーションを指定フォルダへの移動に失敗しました。マイドライブのルートに残ります。エラー: ${moveError}`);
//        ui.alert('警告', `プレゼンテーションを指定フォルダへ移動できませんでした。\nマイドライブのルートに保存されています。`, ui.ButtonSet.OK);
//     }
//   }

//   // --- スライドの処理（ループなど） ---
//   const newPresentation = SlidesApp.openById(presentationId); // 移動後のファイルを開き直す
//   const initialSlide = newPresentation.getSlides()[0];

//   SpreadsheetApp.getActiveSpreadsheet().toast(`処理を開始します。全 ${data.length} 行...`, '開始', -1);

//   data.forEach((row, index) => {
//     const rowNumForLog = index + 2;
//     SpreadsheetApp.getActiveSpreadsheet().toast(`${rowNumForLog} / ${data.length + 1} 行目を処理中...`, '処理中', -1);
//     Logger.log(`行 ${rowNumForLog} を処理中`);

//     try {
//       const newSlide = newPresentation.insertSlide(newPresentation.getSlides().length, templateSlide);
//       const pageElements = newSlide.getPageElements(); // 複製後のスライドの全要素を取得

//       // テキスト要素を代替テキスト(タイトル)で検索して置換
//       for (const altTextTitle in ALT_TEXT_TITLE_MAP_TR) {
//         const colIndex = ALT_TEXT_TITLE_MAP_TR[altTextTitle];
//         if (colIndex >= 0 && colIndex < row.length) {
//           let replacementValue = row[colIndex];
//           if (replacementValue instanceof Date) {
//             replacementValue = Utilities.formatDate(replacementValue, Session.getScriptTimeZone(), 'yyyy/MM/dd');
//           }

//           // 複製後のスライドから、指定された代替テキストタイトルを持つ図形(Shape)を探す
//           const shape = pageElements.find(el => el.getPageElementType() === SlidesApp.PageElementType.SHAPE && el.getTitle() === altTextTitle)?.asShape();

//           if (shape && shape.getText) {
//              shape.getText().setText(String(replacementValue || ''));
//              // Logger.log(`置換 (タイトル: ${altTextTitle}): "${replacementValue}"`);
//           } else {
//              Logger.log(`警告: 行 ${rowNumForLog}: 代替テキストタイトル "${altTextTitle}" を持つテキストボックスが見つかりません。`);
//           }
//         } else if (colIndex !== -1) {
//            Logger.log(`警告: 行 ${rowNumForLog}: 代替テキストタイトル "${altTextTitle}" の列インデックス ${colIndex} が範囲外です。`);
//         }
//       }

//       // イラスト画像を代替テキスト(タイトル)で検索して置換
//        const imageSource = row[ILLUSTRATION_COLUMN_INDEX_TR];
//        let imageBlob = null;

//        // (画像ソースの処理部分は変更なし)
//        if (typeof imageSource === 'string' && imageSource.toLowerCase().startsWith('http')) {
//          const fileId = extractGoogleDriveId_(imageSource);
//          if (fileId) { try { imageBlob = DriveApp.getFileById(fileId).getBlob(); } catch (e) { Logger.log(`警告: 行 ${rowNumForLog}: Driveファイル取得失敗 - ${e}`); } }
//          else { try { imageBlob = UrlFetchApp.fetch(imageSource).getBlob(); } catch (e) { Logger.log(`警告: 行 ${rowNumForLog}: URL画像取得失敗 - ${e}`); } }
//        } else if (typeof imageSource === 'object' && imageSource !== null && imageSource.toString() === 'CellImage') {
//          try { const imageUrl = imageSource.getContentUrl(); if (imageUrl) { imageBlob = UrlFetchApp.fetch(imageUrl).getBlob(); } else { Logger.log(`警告: 行 ${rowNumForLog}: CellImage URL取得不可`); } }
//          catch(e) { Logger.log(`警告: 行 ${rowNumForLog}: CellImage処理エラー - ${e}`); }
//        }

//        if (imageBlob && IMAGE_ALT_TEXT_TITLE_TR) {
//           // 複製後のスライドから、指定された代替テキストタイトルを持つ画像(Image)を探す
//           const imagePlaceholder = pageElements.find(el => el.getPageElementType() === SlidesApp.PageElementType.IMAGE && el.getTitle() === IMAGE_ALT_TEXT_TITLE_TR)?.asImage();

//           if (imagePlaceholder) {
//             imagePlaceholder.replace(imageBlob);
//             Logger.log(`行 ${rowNumForLog}: 画像(タイトル: ${IMAGE_ALT_TEXT_TITLE_TR})を置換しました。`);
//           } else {
//             Logger.log(`警告: 行 ${rowNumForLog}: 代替テキストタイトル "${IMAGE_ALT_TEXT_TITLE_TR}" を持つ画像が見つかりません。`);
//           }
//        } else if (imageSource && IMAGE_ALT_TEXT_TITLE_TR){
//          Logger.log(`警告: 行 ${rowNumForLog}: 列 ${ILLUSTRATION_COLUMN_INDEX_TR + 1} の画像ソースを処理できませんでした。ソース: ${imageSource}`);
//        }

//     } catch (slideError) {
//        Logger.log(`エラー: 行 ${rowNumForLog} のスライド処理中にエラーが発生しました - ${slideError}`);
//     }

//     Utilities.sleep(500);
//   });

//   if (initialSlide && newPresentation.getSlides().length > 1) {
//     try { initialSlide.remove(); Logger.log("最初の空スライドを削除"); }
//     catch (removeError) { Logger.log(`警告: 空スライド削除失敗 - ${removeError}`); }
//   }

//   SpreadsheetApp.getActiveSpreadsheet().toast('完了しました！', '完了', 5);
//   Logger.log(`処理完了。プレゼンテーションURL: ${newPresentation.getUrl()}`);
//   ui.alert('成功', `プレゼンテーションを作成しました: ${newPresentation.getName()}\nURL: ${newPresentation.getUrl()}`, ui.ButtonSet.OK);
// }
