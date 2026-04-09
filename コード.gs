/**
 * Webhook エントリーポイント（HTTP POST）
 * GitHub ActionsなどからHTTP POSTリクエストで呼び出す。
 * デプロイURL: https://script.google.com/macros/s/<SCRIPT_ID>/exec
 *
 * リクエストボディ（JSON）:
 *   { "action": "addNewSupplements" }
 *
 * レスポンス（JSON）:
 *   成功: { "status": "ok", "message": "..." }
 *   エラー: { "status": "error", "message": "..." }
 */
function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents
      ? JSON.parse(e.postData.contents)
      : {};
    const action = body.action || '';

    if (action === 'addNewSupplements') {
      addNewSupplements();
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', message: 'addNewSupplements 完了' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: '不明なaction: ' + action }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Webhook エントリーポイント（HTTP GET）
 * ヘルスチェック用。
 */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Webhook is alive' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function addNewSupplements() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ① 成分一覧・効用・タイミング
  const ws1 = ss.getSheetByName('①成分一覧・効用・タイミング');
  ws1.insertRowsBefore(20, 2);

  const r17 = [
    17, 'ハイチオールCプラス2', 'エスエス製薬（第3類医薬品）',
    'L-システイン 240mg/日、アスコルビン酸(ビタミンC) 500mg/日、パントテン酸カルシウム 24mg/日',
    '1日3回・各2錠（計6錠/日）',
    '肌代謝(ターンオーバー)促進・シミ/そばかす改善、メラニン生成抑制・無色化、アセトアルデヒド分解(二日酔改善)、全身倦怠・疲れだるさ改善',
    '【朝食後・昼食後・就寝前（各2錠）】食前後どちらでも可。1日3回均等分割で血中VC・L-システイン濃度を安定維持。就寝前分はコラーゲン(No.3)との相乗効果あり。',
    '食前後どちらでも可', 360, 3500, 6, '=J20/I20*K20', '=L20*30', '=INT(I20/K20)',
    '⚠️ 注意', '朝食後・昼食後・就寝前（1日3回）',
    '⚠️ ビタミンC +500mg追加→合計1,250mg/日（マルチ250+リポC500+本剤500）。耐容上限2,000mgに近接。消化器症状が出た場合は1日2回に減量。パントテン酸+24mg→合計74mg/日（安全範囲）。L-システイン240mgは安全範囲。',
    'https://www.amazon.co.jp/dp/B09DNHCR5D'
  ];

  const r18 = [
    18, 'プロバイオティクス Acidophilus Blend', '21st Century',
    '乳酸菌4種（L. acidophilus, L. salivarius, B. bifidum, S. thermophilus）10億CFU＋プレバイオティクス Inulin 175mg',
    '1カプセル/日',
    '腸内フローラバランス維持、消化機能サポート、腸管免疫調節（sIgA産生促進）、栄養素吸収促進、Gut-Brain Axisを介した精神安定・ストレス軽減',
    '【朝食と共に】食事中の胃酸希釈・消化酵素環境で生菌の腸到達率が向上。カプリル酸(No.2)とは2時間以上空けること。',
    '食事中推奨', 300, 2200, 1, '=J21/I21*K21', '=L21*30', '=INT(I21/K21)',
    '✅ 通常', '起床時（朝食中）',
    'カプリル酸(No.2)との同時服用禁止（抗菌作用で生菌が死滅）→2時間以上間隔をあけること。抗生物質服用時も同様。乳製品アレルギー注意(Whey含有)。',
    'https://jp.iherb.com/pr/21st-century-acidophilus-probiotic-blend-300-capsules/130241'
  ];

  ws1.getRange(20, 1, 1, r17.length).setValues([r17]);
  ws1.getRange(21, 1, 1, r18.length).setValues([r18]);
  ws1.getRange(20, 1, 1, 18).setBackground('#FFF2CC').setFontColor('#7F4000');
  ws1.getRange(20, 2).setFontWeight('bold');
  ws1.getRange(21, 1, 1, 18).setBackground('#E8F5E9').setFontColor('#1B5E20');
  ws1.getRange(21, 2).setFontWeight('bold');
  ws1.getRange(20, 1, 1, 18).setBorder(true, true, true, true, true, true);
  ws1.getRange(21, 1, 1, 18).setBorder(true, true, true, true, true, true);

  // ② 重複・過剰分析
  const ws2 = ss.getSheetByName('②重複・過剰分析');
  const lr2 = ws2.getLastRow();
  const issues = [
    ['ビタミンC（NEW追加後）', 'No.5(マルチ:250mg)+No.14(リポC:500mg)+No.17(ハイチオール:500mg)', '合計 1,250mg/日（サプリ由来）', '食事由来: 約100〜150mg/日', '1,350〜1,400mg/日', '2,000mg/日（成人耐容上限）', '🟡 上限近接（NEW追加後）', '現状は許容範囲内(1,400/2,000mg)。下痢・腎結石リスク者は1,000mg以内に制限。症状が出た場合はハイチオールを1日2回に減らすかリポソームCを1カプセルに減量。', 'ハイチオール追加で500mg増加。腸管耐容量を超えると浸透圧性下痢。リポソーム型は実効吸収量が高い点に注意。'],
    ['パントテン酸（B5）', 'No.5(マルチ:25mg)+No.17(ハイチオール:24mg) ※B-50中止後', '合計 49mg/日', '食事由来: 約3〜5mg/日', '52〜54mg/日', '明確な耐容上限なし（1,000mg超でも安全と報告）', '🟢 適正範囲（NEW）', 'B-50中止後のマルチ+ハイチオールで49mg/日。十分安全な範囲。代謝酵素補助として有効。', 'パントテン酸の耐容上限は日本では未設定。欧州では1,000mg/日以内が目安。49mgは全く問題なし。'],
    ['乳酸菌（プロバイオティクス）', 'No.18(Acidophilus Blend:10億CFU)', '10億CFU/日', '食事由来: ヨーグルト等を摂る場合に加算', '10億〜数十億CFU/日（食品含む）', '上限設定なし（臨床研究では最大1,000億CFUで安全）', '🟢 適正範囲（NEW）', '健康目的には1〜100億CFUが適量。カプリル酸(No.2)との服用間隔（2時間以上）を厳守すること。', '乳酸菌の過剰摂取での重篤な副作用報告なし。免疫抑制剤服用者・重篤な腸疾患患者は医師相談を推奨。']
  ];
  for (let i = 0; i < issues.length; i++) {
    const row = lr2 + 1 + i;
    ws2.getRange(row, 1, 1, issues[i].length).setValues([issues[i]]);
    const bg = issues[i][6].includes('近接') ? '#FFE699' : '#C6EFCE';
    ws2.getRange(row, 1, 1, issues[i].length).setBackground(bg).setBorder(true, true, true, true, true, true).setWrap(true);
  }

  // ③ 摂取スケジュール
  const ws3 = ss.getSheetByName('③摂取スケジュール');
  const lr3 = ws3.getLastRow();
  let nightRow = -1;
  for (let r = 1; r <= lr3; r++) {
    const v = ws3.getRange(r, 1).getValue().toString();
    if (v.includes('就寝前') && v.includes('推奨')) { nightRow = r; break; }
  }
  if (nightRow === -1) nightRow = lr3;
  ws3.insertRowsBefore(nightRow, 4);
  const sched = [
    [17, 'ハイチオールCプラス2（朝食後 — 1回目）', 2, '朝食後（食前後どちらでも可）', 'ビタミンCはマルチ+リポCと合計1,250mg/日。消化器症状注意。', 'L-システインがターンオーバー促進・メラニン抑制。ビタミンCとの相乗効果で美白効果最大化。'],
    [18, 'プロバイオティクス Acidophilus Blend（朝食中）', 1, '朝食中（食事と必ず一緒に）', 'カプリル酸(No.2)とは2時間以上空けること', '食事中の胃酸希釈で生菌の腸到達率向上。腸内フローラ維持・腸管免疫調節。'],
    ['☀️ 昼食後 推奨サプリメント', '', '', '', '', ''],
    [17, 'ハイチオールCプラス2（昼食後 — 2回目）', 2, '昼食後（食前後どちらでも可）', '1日3回のうちの2回目', '等間隔（朝・昼・就寝前）で血中濃度を安定維持。']
  ];
  for (let i = 0; i < sched.length; i++) {
    ws3.getRange(nightRow + i, 1, 1, 6).setValues([sched[i]]);
    if (i === 2) {
      ws3.getRange(nightRow + 2, 1, 1, 6).setBackground('#2E75B6').setFontColor('#FFFFFF').setFontWeight('bold');
    } else if (sched[i][0] === 17) {
      ws3.getRange(nightRow + i, 1, 1, 6).setBackground('#FFF2CC').setFontColor('#7F4000');
    } else {
      ws3.getRange(nightRow + i, 1, 1, 6).setBackground('#E8F5E9').setFontColor('#1B5E20');
    }
  }
  const nlr3 = ws3.getLastRow();
  ws3.getRange(nlr3 + 1, 1, 1, 6).setValues([[17, 'ハイチオールCプラス2（就寝前 — 3回目）', 2, '就寝直前（就寝30〜60分前）', '1日3回のうちの3回目', '就寝中の細胞修復にビタミンC+L-システインが作用。コラーゲン(No.3)との相乗効果。']]);
  ws3.getRange(nlr3 + 1, 1, 1, 6).setBackground('#FFF2CC').setFontColor('#7F4000');

  // ④ 価格サマリー・継続管理
  const ws4 = ss.getSheetByName('④価格サマリー・継続管理');
  const lr4 = ws4.getLastRow();
  let totalRow = -1;
  for (let r = 1; r <= lr4; r++) {
    if (ws4.getRange(r, 1).getValue().toString().includes('合計')) { totalRow = r; break; }
  }
  if (totalRow === -1) totalRow = lr4 + 1;
  ws4.insertRowsBefore(totalRow, 2);
  const t17 = totalRow, t18 = totalRow + 1;
  ws4.getRange(t17, 1, 1, 12).setValues([[17, 'ハイチオールCプラス2', '⚠️ VC重複注意', 360, 3500, 6, '=J' + t17 + '/I' + t17 + '*K' + t17, '=L' + t17 + '*30', '=INT(I' + t17 + '/K' + t17 + ')', '', '=IF(J' + t17 + '="","",J' + t17 + '+I' + t17 + ')', 'Amazon.co.jp等で購入可（iHerb非取扱）']]);
  ws4.getRange(t18, 1, 1, 12).setValues([[18, 'プロバイオティクス Acidophilus Blend', '✅ 通常摂取', 300, 2200, 1, '=J' + t18 + '/I' + t18 + '*K' + t18, '=L' + t18 + '*30', '=INT(I' + t18 + '/K' + t18 + ')', '', '=IF(J' + t18 + '="","",J' + t18 + '+I' + t18 + ')', 'iHerb: 21st-century-acidophilus-probiotic-blend-300-capsules']]);
  ws4.getRange(t17, 1, 1, 12).setBackground('#FFF2CC').setFontColor('#7F4000').setBorder(true, true, true, true, true, true);
  ws4.getRange(t18, 1, 1, 12).setBackground('#E8F5E9').setFontColor('#1B5E20').setBorder(true, true, true, true, true, true);
  ws4.getRange(t17, 7).setNumberFormat('#,##0');
  ws4.getRange(t17, 8).setNumberFormat('#,##0');
  ws4.getRange(t18, 7).setNumberFormat('#,##0');
  ws4.getRange(t18, 8).setNumberFormat('#,##0');

  // ⑤ 飲み合わせ注意
  const ws5 = ss.getSheetByName('⑤飲み合わせ注意');
  const lr5 = ws5.getLastRow();
  const ints = [
    ['🟡 中程度', 'ハイチオールCプラス2 (No.17)', 'リポソームVC (No.14) + マルチビタミン (No.5)', '3製品合算でビタミンC 1,250mg/日。各製品単独は問題ないが合算で耐容上限2,000mgに近接。リポソーム型は実効吸収量が高いため実質的な負荷はさらに大きい可能性。', '浸透圧性下痢、腎結石リスク（シュウ酸生成増加）', '消化器症状が出た場合はハイチオールを1日2回へ減らすか、リポソームCを1カプセルに削減。腎結石既往者は総量1,000mg以内を推奨。'],
    ['🟡 中程度', 'プロバイオティクス (No.18)', 'カプリル酸 (No.2)', 'カプリル酸の抗菌・抗真菌作用が同時服用でプロバイオティクスの生菌を殺菌。同時摂取で相互の効果が完全に相殺される。', 'プロバイオティクスの有効性が大幅に低下または消失', '必ず2時間以上間隔を空けること。推奨: プロバイオティクス→朝食中、カプリル酸→朝食から2時間後 or 夕食中。'],
    ['🟢 相乗効果', 'ハイチオールCプラス2 (No.17)', '加水分解コラーゲンペプチド (No.3)', 'L-システインはケラチン構成アミノ酸。ビタミンCはコラーゲン合成律速酵素の必須補因子。両者で皮膚・爪・毛髪の合成を相補的に促進。', '特になし（推奨される組み合わせ）', '就寝前にコラーゲン+ハイチオールを同時服用することで成長ホルモン分泌ピーク時の皮膚再生を最大化。'],
    ['🟢 相乗効果', 'プロバイオティクス (No.18)', 'カプリル酸 (No.2) ※時間差服用', 'カプリル酸で腸内カンジダを除菌後、プロバイオティクスで善玉菌を定着補充。腸内環境改善のシーケンシャルアプローチとして理想的。', '特になし（2時間以上の間隔が前提）', 'カプリル酸服用2時間後にプロバイオティクス、または就寝前にプロバイオティクスのみ単独服用でも可。']
  ];
  for (let i = 0; i < ints.length; i++) {
    const row = lr5 + 1 + i;
    ws5.getRange(row, 1, 1, ints[i].length).setValues([ints[i]]);
    let bg = '#C6EFCE';
    if (ints[i][0].includes('重大')) bg = '#FF7F7F';
    else if (ints[i][0].includes('中程度')) bg = '#FFE699';
    ws5.getRange(row, 1, 1, ints[i].length).setBackground(bg).setBorder(true, true, true, true, true, true).setWrap(true);
  }

  // ⑥ テンプレート更新履歴
  const ws6 = ss.getSheetByName('⑥新規サプリ追加テンプレ');
  const lr6 = ws6.getLastRow();
  ws6.getRange(lr6 + 2, 1).setValue('📅 最終更新').setFontWeight('bold');
  ws6.getRange(lr6 + 2, 2).setValue('2026年4月9日 — No.17 ハイチオールCプラス2・No.18 プロバイオティクス Acidophilus Blend を追加（全6シート更新済み）').setFontColor('#595959');

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('✅ 完了！\nNo.17 ハイチオールCプラス2\nNo.18 プロバイオティクス Acidophilus Blend\n全6シートへの追記が完了しました。');
}
